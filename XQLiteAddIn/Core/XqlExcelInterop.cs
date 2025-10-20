// XqlExcelInterop.cs
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace XQLite.AddIn
{
    /// <summary>
    /// Excel 개체모델과 Add-in 내부 모듈(XqlSync/XqlCollab/XqlMetaRegistry)을 연결한다.
    /// - 리본/메뉴 → 명령 라우팅(Cmd_*)
    /// - Excel 이벤트(시트 변경/선택 변경/통합문서 열기·닫기) 핸들링
    /// - 헤더 범위 탐색, 타입 툴팁/주석 표시
    /// - Presence/락 하트비트 전송(선택 변경 시)
    /// - 셀 편집 → 2초 디바운스 업서트 큐 적재(XqlSync)
    /// </summary>
    internal sealed class XqlExcelInterop(Excel.Application app, XqlSync sync, XqlCollab collab, XqlSheet sheet, XqlBackup backup) : IDisposable
    {
        private readonly Excel.Application _app = app ?? throw new ArgumentNullException(nameof(app));
        private readonly XqlSync _sync = sync ?? throw new ArgumentNullException(nameof(sync));
        private readonly XqlCollab _collab = collab ?? throw new ArgumentNullException(nameof(collab));
        private readonly XqlSheet _sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
        private readonly XqlBackup _backup = backup ?? throw new ArgumentNullException(nameof(backup));

        private bool _started;

        // ========= 수명 주기 =========

        public void Start()
        {
            if (_started) return;
            _started = true;

            try
            {
                var app = ExcelDnaUtil.Application as Excel.Application;
                var wb = app?.ActiveWorkbook;
                if (wb != null)
                {
                    var full = wb.FullName; // c:\path\file.xlsx
                    _sync.InitPersistentState(full, XqlConfig.Project);
                }
            }
            catch { /* 무시 */ }

            _app.SheetChange += App_SheetChange;
            _app.SheetSelectionChange += App_SheetSelectionChange;
            _app.WorkbookOpen += App_WorkbookOpen;
            _app.WorkbookBeforeClose += App_WorkbookBeforeClose;
        }

        public void Stop()
        {
            if (!_started) return;
            _started = false;

            _app.SheetChange -= App_SheetChange;
            _app.SheetSelectionChange -= App_SheetSelectionChange;
            _app.WorkbookOpen -= App_WorkbookOpen;
            _app.WorkbookBeforeClose -= App_WorkbookBeforeClose;
        }

        public void Dispose()
        {
            Stop();
        }

        // ========= 명령(리본/메뉴) =========
        public async void Cmd_CommitSync()
        {
            try
            {
                await EnsureActiveSheetSchema();             // 헤더 → 서버 스키마 동기화
                await _sync.FlushUpsertsNow(force: true);    // 변경된 셀만 즉시 업서트
            }
            catch (Exception ex)
            {
                XqlLog.Warn("CommitSmart failed: " + ex.Message);
            }
        }

        public async void Cmd_PullOnly()
        {
            try
            {
                var app = ExcelDnaUtil.Application as Excel.Application;
                if (app == null) { await _sync.PullSince(); return; }

                var ws = app.ActiveSheet as Excel.Worksheet;
                if (ws == null) { await _sync.PullSince(); return; }

                // ✅ 부트스트랩 필요 판단
                bool needsBootstrap = XqlSheet.NeedsBootstrap(ws);
                await _sync.PullSince(needsBootstrap ? 0 : (long?)null);
            }
            catch (Exception ex)
            {
                XqlLog.Warn("PullOnly failed: " + ex.Message);
            }
        }

        public void Cmd_RecoverFromExcel()
        {
            var _ = _backup.RecoverFromExcel();
        }

        // ========= Excel 이벤트 =========

        private void App_WorkbookOpen(Excel.Workbook wb)
        {
            try
            {
                if (wb != null)
                    _sync.InitPersistentState(wb.FullName, XqlConfig.Project);
            }
            catch { /* ignore */ }
            finally { XqlCommon.ReleaseCom(wb); }
        }

        private void App_WorkbookBeforeClose(Excel.Workbook wb, ref bool Cancel)
        {
            var _ = _collab.ReleaseByMe();
            try { _sync.Stop(); } catch { }
            XqlCommon.ReleaseCom(wb);
        }

        /// <summary>시트에서 헤더와 데이터 범위를 일관되게 구한다.</summary>
        private static (Excel.Range? header, Excel.Range? data, Excel.ListObject? lo) ResolveHeaderAndData(Excel.Worksheet sh)
        {
            Excel.Range? header = XqlSheetView.ResolveHeader(sh, null, XqlAddIn.Sheet!)
                          ?? (XqlSheet.TryGetHeaderMarker(sh, out var mk) ? mk : XqlSheet.GetHeaderRange(sh));
            if (header == null) return (null, null, null);
            var lo = XqlSheet.FindListObjectContaining(sh, header);
            if (lo?.DataBodyRange != null) return (header, lo.DataBodyRange, lo);
            var first = (Excel.Range)header.Offset[1, 0];
            var last = sh.Cells[sh.Rows.Count, header.Column + header.Columns.Count - 1];
            var data = sh.Range[first, last];
            XqlCommon.ReleaseCom(first, last);
            return (header, data, lo);
        }

        // 변경 이벤트에서 호출
        private void App_SheetChange(object Sh, Excel.Range target)
        {
            Excel.Worksheet? sh = null;
            Excel.Range? header = null;
            Excel.Range? data = null;
            Excel.Range? intersect = null;
            Excel.ListObject? lo = null;

            try
            {
                sh = Sh as Excel.Worksheet;
                if (sh == null) return;

                var sm = _sheet.GetOrCreateSheet(sh.Name);
                (header, data, lo) = ResolveHeaderAndData(sh);
                if (header == null || data == null) return;

                // 헤더 편집이면 캐시 무효화 + 커밋 가능 알림
                var hitHeader = sh.Application.Intersect(target, header) as Excel.Range;
                if (hitHeader != null)
                {
                    XqlSheetView.InvalidateHeaderCache(sh.Name);
                    XqlEvents.RaiseSchemaChanged(); // 리본 상태 갱신
                    XqlCommon.ReleaseCom(hitHeader);
                    return;
                }

                intersect = sh.Application.Intersect(target, data) as Excel.Range;
                if (intersect == null) return;

                var table = string.IsNullOrWhiteSpace(sm.TableName) ? sh.Name : sm.TableName!;
                var keyColName = string.IsNullOrWhiteSpace(sm.KeyColumn) ? "id" : sm.KeyColumn!;

                var areas = intersect.Areas;
                try
                {
                    for (int ai = 1; ai <= areas.Count; ai++)
                    {
                        Range? area = null;
                        try
                        {
                            area = (Excel.Range)areas[ai];
                            foreach (Excel.Range cell in area.Cells)
                            {
                                Excel.Range? hdrCell = null;
                                Excel.Range? keyCell = null;
                                try
                                {
                                    int hdrIdx = cell.Column - header.Column + 1;
                                    if (hdrIdx < 1 || hdrIdx > header.Columns.Count) continue;

                                    hdrCell = (Excel.Range)header.Cells[1, hdrIdx];
                                    var colName = (hdrCell.Value2 as string)?.Trim();
                                    // 편집 이벤트에서는 빈 헤더라도 열문자 사용(사용자 피드백용)
                                    if (string.IsNullOrWhiteSpace(colName))
                                        colName = XqlCommon.ColumnIndexToLetter(hdrCell.Column);

                                    int keyAbsCol = XqlSheet.FindKeyColumnAbsolute(header, sm.KeyColumn);
                                    keyCell = sh.Cells[cell.Row, keyAbsCol] as Excel.Range;

                                    var rowKeyObj = keyCell?.Value2;
                                    string? rowKey = rowKeyObj?.ToString();

                                    object? value = cell.Value2;
                                    _sync.EnqueueIfChanged(table, rowKey!, colName!, value);
                                }
                                finally
                                {
                                    XqlCommon.ReleaseCom(keyCell, hdrCell, cell);
                                }
                            }
                        }
                        finally { XqlCommon.ReleaseCom(area); }
                    }
                }
                finally { XqlCommon.ReleaseCom(areas); }
            }
            catch (Exception ex)
            {
                XqlLog.Warn("OnWorksheetChange: " + ex.Message);
            }
            finally { XqlCommon.ReleaseCom(intersect, data, lo, header, sh); }
        }

        private void App_SheetSelectionChange(object Sh, Excel.Range Target)
        {
            try
            {
                var ws = Sh as Excel.Worksheet; if (ws == null) return;
                string sheet = ws.Name;
                string cell = Target?.Address[false, false] ?? "";
                XqlAddIn.Collab?.SelectionChanged(sheet, cell);
            }
            catch { /* non-fatal */ }
            finally { XqlCommon.ReleaseCom(Target); }
        }

        /// <summary>
        /// 활성 시트의 헤더/메타를 읽어 서버 스키마(테이블/컬럼)와 동기화.
        /// - Rename → Alter → Add → Drop 순으로 항상 실행(드랍 자동).
        /// - 빈 헤더는 "삭제 의도"로 해석(추가/유지 대상에서 제외).
        /// </summary>
        private async Task EnsureActiveSheetSchema()
        {
            if (XqlAddIn.Backend is not IXqlBackend be) return;

            var app = ExcelDnaUtil.Application as Excel.Application;
            if (app == null) return;

            Excel.Worksheet? ws = null; Excel.Range? header = null;
            try
            {
                ws = app.ActiveSheet as Excel.Worksheet;
                if (ws == null) return;

                var sm = _sheet.GetOrCreateSheet(ws.Name);

                // ① 헤더/컬럼명 — 빈 칸은 그대로(삭제 의도 반영)
                var (hdr, headerNamesRaw) = XqlSheet.GetHeaderAndNames(ws);
                header = hdr;
                if (header == null || headerNamesRaw is not { Count: > 0 }) return;

                var headerNames = NormalizeHeaderNamesKeepBlanks(headerNamesRaw); // ★ 빈 이름 유지
                var headerNamesForMeta = headerNames.Where(n => !string.IsNullOrWhiteSpace(n)).ToList();
                _sheet.EnsureColumns(ws.Name, headerNamesForMeta);

                var table = string.IsNullOrWhiteSpace(sm.TableName) ? ws.Name : sm.TableName!;
                var key = string.IsNullOrWhiteSpace(sm.KeyColumn) ? "id" : sm.KeyColumn!;

                // ② 테이블 보장
                await be.TryCreateTable(table, key);
                XqlSheetView.InvalidateHeaderCache(ws.Name);

                // ③ 서버 컬럼 조회
                var serverCols = await be.GetTableColumns(table);
                var serverSet = new HashSet<string>(serverCols.Select(c => c.name), StringComparer.OrdinalIgnoreCase);

                // ④ Rename 추론 → 적용
                var renames = InferRenamesByIndex(serverCols, headerNames);
                if (renames.Count > 0)
                {
                    try { await be.TryRenameColumns(table, renames); }
                    catch (Exception ex) { XqlLog.Warn("RenameColumns skipped: " + ex.Message); }
                    finally
                    {
                        serverCols = await be.GetTableColumns(table); // 갱신
                        serverSet = new HashSet<string>(serverCols.Select(c => c.name), StringComparer.OrdinalIgnoreCase);
                        XqlSheetView.InvalidateHeaderCache(ws.Name);
                    }
                }

                // ⑤ Alter 후보(타입/Null)
                var desired = BuildDesiredColumnSpec(sm, headerNamesForMeta);
                var alters = InferAlters(serverCols, desired);
                if (alters.Count > 0)
                {
                    try { await be.TryAlterColumns(table, alters); }
                    catch (Exception ex) { XqlLog.Warn("AlterColumns skipped: " + ex.Message); }
                }

                // ⑥ Add 후보 — 빈 이름 제외 + rename 타겟 제외
                var renamedTargets = new HashSet<string>(renames.Select(r => r.To), StringComparer.OrdinalIgnoreCase);
                var addTargets = headerNamesForMeta
                    .Where(n => !serverSet.Contains(n))
                    .Where(n => !renamedTargets.Contains(n))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (addTargets.Count > 0)
                {
                    var defs = addTargets.Select(name =>
                    {
                        if (!sm.Columns.TryGetValue(name, out var ct))
                        {
                            ct = new XqlSheet.ColumnType { Kind = XqlSheet.ColumnKind.Text, Nullable = true };
                            sm.SetColumn(name, ct);
                        }
                        return new ColumnDef
                        {
                            Name = name,
                            Kind = ct.Kind switch
                            {
                                XqlSheet.ColumnKind.Int => "integer",
                                XqlSheet.ColumnKind.Real => "real",
                                XqlSheet.ColumnKind.Bool => "bool",
                                XqlSheet.ColumnKind.Date => "integer",  // epoch ms
                                XqlSheet.ColumnKind.Json => "json",
                                _ => "text"
                            },
                            NotNull = !ct.Nullable,
                            Check = null
                        };
                    }).ToList();

                    try { await be.TryAddColumns(table, defs); }
                    catch (Exception ex) { XqlLog.Warn("AddColumns skipped: " + ex.Message); }
                    finally { XqlSheetView.InvalidateHeaderCache(ws.Name); }
                }

                // ⑦ Drop 후보 — **항상 실행** (PK/예약 제외, 헤더에 없는 컬럼 드랍)
                {
                    var metaCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { key, "row_version", "updated_at", "deleted" };
                    var headerSet = new HashSet<string>(headerNamesForMeta, StringComparer.OrdinalIgnoreCase);

                    var drop = serverCols
                        .Where(c => !c.pk && !metaCols.Contains(c.name))
                        .Select(c => c.name?.Trim() ?? "")
                        .Where(n => n.Length > 0)
                        .Where(n => !headerSet.Contains(n))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    if (drop.Count > 0)
                    {
                        try { await be.TryDropColumns(table, drop); }
                        catch (Exception ex) { XqlLog.Warn("DropColumns skipped: " + ex.Message); }
                        finally { XqlSheetView.InvalidateHeaderCache(ws.Name); }
                    }
                }

                try { XqlSheetView.RegisterTableSheet(table, ws.Name); } catch { /* ignore */ }
            }
            catch (Exception ex) { XqlLog.Warn("EnsureActiveSheetSchema: " + ex.Message); }
            finally { XqlCommon.ReleaseCom(header, ws); }
        }

        // ======== 내부 유틸 ========

        /// <summary>헤더 이름을 Trim만 하고, 빈 칸은 그대로 유지한다(삭제 의도 반영).</summary>
        private static List<string> NormalizeHeaderNamesKeepBlanks(IList<string> names)
        {
            var list = new List<string>(names.Count);
            for (int i = 0; i < names.Count; i++)
                list.Add((names[i] ?? "").Trim());
            return list;
        }

        /// <summary>
        /// 서버 컬럼과 헤더 컬럼을 비교하여 (인덱스 기반) rename 후보를 추론.
        /// 같은 위치에서 이름만 달라졌으면 rename으로 간주(PK/예약 컬럼 제외).
        /// </summary>
        private static List<RenameDef> InferRenamesByIndex(IReadOnlyList<ColumnInfo> serverCols, IList<string> header)
        {
            var renames = new List<RenameDef>();
            if (serverCols.Count == 0 || header.Count == 0) return renames;

            var bizCols = serverCols.Where(c => !c.pk && !IsReserved(c.name)).ToList();
            int lim = Math.Min(bizCols.Count, header.Count);

            for (int i = 0; i < lim; i++)
            {
                var oldName = (bizCols[i].name ?? "").Trim();
                var newName = (header[i] ?? "").Trim();

                if (string.IsNullOrEmpty(oldName) || string.IsNullOrEmpty(newName)) continue;
                if (oldName.Equals(newName, StringComparison.OrdinalIgnoreCase)) continue;

                renames.Add(new RenameDef { From = oldName, To = newName });
            }

            return renames
                .Where(r => !string.IsNullOrWhiteSpace(r.From) && !string.IsNullOrWhiteSpace(r.To)
                         && !r.From.Equals(r.To, StringComparison.OrdinalIgnoreCase))
                .GroupBy(r => (From: r.From.ToLowerInvariant(), To: r.To.ToLowerInvariant()))
                .Select(g => g.First())
                .ToList();
        }

        private static bool IsReserved(string name)
        {
            return string.Equals(name, "row_version", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "updated_at", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "deleted", StringComparison.OrdinalIgnoreCase);
        }

        private static Dictionary<string, (string Type, bool NotNull, string? Check)> BuildDesiredColumnSpec(XqlSheet.Meta sm, IEnumerable<string> headerNames)
        {
            var result = new Dictionary<string, (string, bool, string?)>(StringComparer.OrdinalIgnoreCase);
            foreach (var n in headerNames)
            {
                if (string.IsNullOrWhiteSpace(n)) continue;
                if (!sm.Columns.TryGetValue(n, out var ct))
                {
                    ct = new XqlSheet.ColumnType { Kind = XqlSheet.ColumnKind.Text, Nullable = true };
                    sm.SetColumn(n, ct);
                }
                string type = ct.Kind switch
                {
                    XqlSheet.ColumnKind.Int => "integer",
                    XqlSheet.ColumnKind.Real => "real",
                    XqlSheet.ColumnKind.Bool => "bool",
                    XqlSheet.ColumnKind.Date => "integer",
                    XqlSheet.ColumnKind.Json => "json",
                    _ => "text"
                };
                bool notNull = !ct.Nullable;
                result[n] = (type, notNull, null);
            }
            return result;
        }

        private static List<AlterDef> InferAlters(IEnumerable<ColumnInfo> serverCols, Dictionary<string, (string Type, bool NotNull, string? Check)> desired)
        {
            var list = new List<AlterDef>();
            foreach (var sc in serverCols)
            {
                if (sc.pk) continue;
                if (IsReserved(sc.name)) continue;
                if (!desired.TryGetValue(sc.name, out var want)) continue;

                var serverType = (sc.type ?? "").Trim();
                var wantType = want.Type.Trim();

                bool typeDiff = !serverType.Equals(wantType, StringComparison.OrdinalIgnoreCase);
                bool nnDiff = sc.notnull != want.NotNull;

                if (typeDiff || nnDiff)
                {
                    list.Add(new AlterDef
                    {
                        Name = sc.name,
                        ToType = typeDiff ? wantType : null,
                        ToNotNull = nnDiff ? want.NotNull : null,
                        ToCheck = null
                    });
                }
            }
            return list;
        }
    }
}
