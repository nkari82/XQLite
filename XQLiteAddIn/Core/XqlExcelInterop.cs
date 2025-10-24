// XqlExcelInterop.cs  — SmartCom<T> 적용 버전 (lastDataRow 계산 보강 + Commit 전/후 스키마 보장)
using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using static XQLite.AddIn.XqlCommon;

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
                if (app != null)
                {
                    using var wb = SmartCom<Excel.Workbook>.Wrap(app.ActiveWorkbook);
                    var full = wb?.Value?.FullName;
                    if (!string.IsNullOrEmpty(full))
                        _sync.InitPersistentState(full!, XqlConfig.Project);
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

        // 베스트: 행 단위 커밋
        // - 기존행(id 있음) → upsertRows
        // - 신규행(id 없음) → upsertCells(임시키 포함) → assigned로 id 반영
        public async void Cmd_CommitSync()
        {
            try
            {
                var app = ExcelDnaUtil.Application as Excel.Application;
                if (app == null) return;

                using var ws = SmartCom<Excel.Worksheet>.Wrap(app.ActiveSheet as Excel.Worksheet);
                if (ws?.Value == null) return;

                // ✅ 0) 커밋 전에 항상 스키마부터(타입/이름 정규화 포함) 서버에 확정
                await EnsureActiveSheetSchema().ConfigureAwait(false);

                var sm = _sheet.GetOrCreateSheet(ws.Value.Name);

                using var header = SmartCom<Excel.Range>.Wrap(XqlSheetView.GetHeaderOrFallback(ws.Value));
                if (header?.Value == null) return;

                var headers = XqlSheet.ComputeHeaderNames(header.Value);
                string keyName = string.IsNullOrWhiteSpace(sm.KeyColumn) ? "id" : sm.KeyColumn!;
                int keyIdx1 = XqlSheet.FindKeyColumnIndex(headers, keyName);
                if (keyIdx1 <= 0) keyIdx1 = 1;

                int firstDataRow = header.Value.Row + 1;

                // ⚠ UsedRange만으로는 첫 입력 직후 갱신이 늦을 수 있어 보강 로직으로 최종 데이터 행을 산출한다.
                int lastDataRow = GetLastDataRow(ws.Value, header.Value, firstDataRow, headers.Count);

                if (lastDataRow < firstDataRow)
                {
                    // 데이터가 전혀 없으면 커밋할 것이 없다.
                    return;
                }

                // 분기: 기존행/신규행
                var rowsForUpsertRows = new List<Dictionary<string, object?>>();
                var cellsForUpsertCells = new List<EditCell>();
                var tempRowKeyToExcelRow = new Dictionary<string, int>(StringComparer.Ordinal);

                for (int r = firstDataRow; r <= lastDataRow; r++)
                {
                    object? idVal = GetCell(ws.Value, r, header.Value.Column + keyIdx1 - 1);
                    string idStr = XqlCommon.Canonicalize(idVal) ?? "";

                    var obj = new Dictionary<string, object?>(StringComparer.Ordinal);
                    for (int i = 0; i < headers.Count; i++)
                    {
                        var col = headers[i];
                        if (string.IsNullOrWhiteSpace(col)) continue;
                        object? v = GetCell(ws.Value, r, header.Value.Column + i);
                        if (string.Equals(col, keyName, StringComparison.OrdinalIgnoreCase))
                        {
                            if (!string.IsNullOrWhiteSpace(idStr))
                                obj[keyName] = idStr;
                            continue;
                        }
                        obj[col] = v is DateTime dt ? dt : v;
                    }

                    bool anyData = obj.Any(kv =>
                        !kv.Key.Equals(keyName, StringComparison.OrdinalIgnoreCase) &&
                        kv.Value != null &&
                        !string.IsNullOrWhiteSpace(Convert.ToString(kv.Value)));
                    if (!anyData && string.IsNullOrWhiteSpace(idStr)) continue;

                    if (!string.IsNullOrWhiteSpace(idStr))
                    {
                        obj[keyName] = idStr;
                        rowsForUpsertRows.Add(obj);
                    }
                    else
                    {
                        string tempKey = "-" + r.ToString();
                        tempRowKeyToExcelRow[tempKey] = r;

                        foreach (var kv in obj)
                        {
                            if (string.Equals(kv.Key, keyName, StringComparison.OrdinalIgnoreCase)) continue; // PK 제외
                            var val = kv.Value;
                            cellsForUpsertCells.Add(new EditCell(
                                ws.Value.Name,
                                tempKey,
                                kv.Key,
                                val
                            ));
                        }
                    }
                }

                // 서버 호출 (⚠ await 전에는 RCW를 더 이상 들고있지 않음 — 모두 값으로만 유지)
                if (XqlAddIn.Backend is IXqlBackend be)
                {
                    // 3-a) 기존행: upsertRows
                    if (rowsForUpsertRows.Count > 0)
                    {
                        var resp = await be.UpsertRows(ws.Value.Name, rowsForUpsertRows).ConfigureAwait(false);
                        if (resp?.Errors is { Count: > 0 })
                            XqlLog.Warn("Commit errors (upsertRows): " + string.Join("; ", resp.Errors));
                    }

                    // 3-b) 신규행: upsertCells (assigned 처리)
                    if (cellsForUpsertCells.Count > 0)
                    {
                        var resp2 = await be.UpsertCells(cellsForUpsertCells).ConfigureAwait(false);
                        if (resp2?.Errors is { Count: > 0 })
                            XqlLog.Warn("Commit errors (upsertCells): " + string.Join("; ", resp2.Errors));

                        // 서버가 발급한 id를 시트에 반영
                        if (resp2?.Assigned is { Count: > 0 })
                        {
                            foreach (var a in resp2.Assigned)
                            {
                                if (a == null) continue;
                                if (!string.Equals(a.Table, ws.Value.Name, StringComparison.Ordinal)) continue;
                                if (string.IsNullOrWhiteSpace(a.NewId)) continue;
                                if (a.TempRowKey == null) continue;

                                if (tempRowKeyToExcelRow.TryGetValue(a.TempRowKey, out var rowIdx))
                                {
                                    using var keyCell = SmartCom<Excel.Range>.Acquire(() => (Excel.Range)ws.Value.Cells[rowIdx, header.Value.Column + keyIdx1 - 1]);
                                    try { if (keyCell.Value != null) keyCell.Value.Value2 = a.NewId; } catch { }
                                }
                            }
                        }
                    }
                }

                // 4) 최신 반영을 위해 Pull
#pragma warning disable CS4014
                _sync.PullSince(null);
#pragma warning restore CS4014

                // ✅ 5) 혹시라도 동시성으로 TEXT가 잠깐 만들어졌다면 여기서 한 번 더 스키마 보정
                await EnsureActiveSheetSchema().ConfigureAwait(false);
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

                using var ws = SmartCom<Excel.Worksheet>.Wrap(app.ActiveSheet as Excel.Worksheet);
                if (ws?.Value == null) { await _sync.PullSince(); return; }

                bool needsBootstrap = XqlSheet.NeedsBootstrap(ws.Value);
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
            // 이벤트 파라미터 RCW는 Excel이 관리하므로 해제하지 않음.
        }

        private void App_WorkbookBeforeClose(Excel.Workbook wb, ref bool Cancel)
        {
            var _ = _collab.ReleaseByMe();
            try { _sync.Stop(); } catch { }
            // 이벤트 파라미터 RCW 해제는 생략(Excel이 소유)
        }

        /// <summary>시트에서 헤더와 데이터 범위를 일관되게 구한다.</summary>
        private static (Excel.Range? header, Excel.Range? data, Excel.ListObject? lo) ResolveHeaderAndData(Excel.Worksheet sh)
        {
            var header = XqlSheetView.ResolveHeader(sh, null, XqlAddIn.Sheet!)
                         ?? (XqlSheet.TryGetHeaderMarker(sh, out var mk) ? mk : XqlSheet.GetHeaderRange(sh));
            if (header == null) return (null, null, null);

            var lo = XqlSheet.FindListObjectContaining(sh, header);
            if (lo?.DataBodyRange != null) return (header, lo.DataBodyRange, lo);

            using var first = SmartCom<Excel.Range>.Acquire(() => (Excel.Range)header.Offset[1, 0]);
            using var last = SmartCom<Excel.Range>.Acquire(() => (Excel.Range)sh.Cells[sh.Rows.Count, header.Column + header.Columns.Count - 1]);
            var data = sh.Range[first.Value, last.Value];
            return (header, data, lo);
        }

        // 변경 이벤트에서 호출
        private void App_SheetChange(object Sh, Excel.Range target)
        {
            try
            {
                var sh = Sh as Excel.Worksheet;
                if (sh == null) return;

                var sm = _sheet.GetOrCreateSheet(sh.Name);
                var (header, data, lo) = ResolveHeaderAndData(sh);
                if (header == null || data == null) return;

                // 헤더 편집이면 캐시 무효화 + 커밋 가능 알림
                using (var hitHeader = SmartCom<Excel.Range>.Wrap(SafeIntersect(sh.Application, target, header)))
                {
                    if (hitHeader?.Value != null)
                    {
                        XqlSheetView.InvalidateHeaderCache(sh.Name);
                        XqlEvents.RaiseSchemaChanged(); // 리본 상태 갱신
                        return;
                    }
                }

                using var intersect = SmartCom<Excel.Range>.Wrap(SafeIntersect(sh.Application, target, data));
                if (intersect?.Value == null) return;

                var table = string.IsNullOrWhiteSpace(sm.TableName) ? sh.Name : sm.TableName!;
                var keyColName = string.IsNullOrWhiteSpace(sm.KeyColumn) ? "id" : sm.KeyColumn!;

                using var areas = SmartCom<Excel.Range>.Wrap(intersect.Value.Areas);
                int areaCount = areas?.Value?.Count ?? 0;
                for (int ai = 1; ai <= areaCount; ai++)
                {
                    using var area = SmartCom<Excel.Range>.Acquire(() => (Excel.Range)areas!.Value![ai]);
                    if (area?.Value == null) continue;

                    foreach (Excel.Range rawCell in area.Value.Cells)
                    {
                        using var cell = SmartCom<Excel.Range>.Wrap(rawCell);
                        using var hdrCell = SmartCom<Excel.Range>.Acquire(() =>
                        {
                            int hdrIdx = cell!.Value!.Column - header!.Column + 1;
                            if (hdrIdx < 1 || hdrIdx > header!.Columns.Count) return null;
                            return (Excel.Range)header!.Cells[1, hdrIdx];
                        });

                        if (hdrCell?.Value == null) continue;

                        string? colName = (hdrCell.Value.Value2 as string)?.Trim();
                        if (string.IsNullOrWhiteSpace(colName))
                            colName = XqlCommon.ColumnIndexToLetter(hdrCell.Value.Column);

                        int keyAbsCol = XqlSheet.FindKeyColumnAbsolute(header, sm.KeyColumn);

                        using var keyCell = SmartCom<Excel.Range>.Acquire(() => (Excel.Range)sh.Cells[cell!.Value!.Row, keyAbsCol]);
                        string? rowKey = keyCell?.Value?.Value2?.ToString();

                        object? value = cell?.Value?.Value2;
                        _sync.EnqueueIfChanged(table, rowKey!, colName!, value);
                    }
                }
            }
            catch (Exception ex)
            {
                XqlLog.Warn("OnWorksheetChange: " + ex.Message);
            }
        }

        private static Excel.Range? SafeIntersect(Excel.Application app, Excel.Range a, Excel.Range b)
        {
            try { return app.Intersect(a, b) as Excel.Range; } catch { return null; }
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
        }

        /// <summary>
        /// 활성 시트의 헤더/메타를 읽어 서버 스키마(테이블/컬럼)와 동기화.
        /// - Rename → Alter → Add → Drop 순으로 항상 실행(드랍 자동).
        /// - 빈 헤더는 "삭제 의도"로 해석(추가/유지 대상에서 제외).
        /// </summary>
        private async Task EnsureActiveSheetSchema()
        {
            if (XqlAddIn.Backend is not IXqlBackend be) return;

            // 1) UI 스냅샷 (Excel UI 스레드에서만 COM 접근)
            var snap = await XqlCommon.OnExcelThreadAsync(() =>
            {
                var result = new SchemaSnapshot();
                var app = ExcelDnaUtil.Application as Excel.Application;
                if (app == null) return result;

                using var ws = SmartCom<Excel.Worksheet>.Wrap(app.ActiveSheet as Excel.Worksheet);
                if (ws?.Value == null) return result;

                var sm = _sheet.GetOrCreateSheet(ws.Value.Name);

                var (hdr, names0) = XqlSheet.GetHeaderAndNames(ws.Value);
                using var header = SmartCom<Excel.Range>.Wrap(hdr);
                if (header?.Value == null || names0 is not { Count: > 0 }) return result;

                // 빈/머지 헤더 보정
                var normalizedHeader = NormalizeHeaderNamesWithLetters(header.Value, names0);

                // 메타 레지스트리 최신화
                _sheet.EnsureColumns(ws.Value.Name, normalizedHeader);

                result.SheetName = ws.Value.Name;
                result.Table = string.IsNullOrWhiteSpace(sm.TableName) ? ws.Value.Name : sm.TableName!;
                result.Key = string.IsNullOrWhiteSpace(sm.KeyColumn) ? "id" : sm.KeyColumn!;
                result.HeaderNames = normalizedHeader;
                result.Meta = sm;
                result.HasHeader = true;
                return result;
            });

            if (!snap.HasHeader) return;

            // 2) 서버 스키마 작업 (COM 접근 없음 — 자유로운 async)
            await be.TryCreateTable(snap.Table!, snap.Key!).ConfigureAwait(false);
            XqlSheetView.InvalidateHeaderCache(snap.SheetName!);

            var serverCols = await be.GetTableColumns(snap.Table!).ConfigureAwait(false);
            var serverSet = new HashSet<string>(serverCols.Select(c => c.name), StringComparer.OrdinalIgnoreCase);

            // 2-1) Rename 추론(인덱스 기반) → 수행
            var renames = InferRenamesByIndex(serverCols, snap.HeaderNames!);
            if (renames.Count > 0)
            {
                try { await be.TryRenameColumns(snap.Table!, renames).ConfigureAwait(false); }
                catch (Exception ex) { XqlLog.Warn("RenameColumns skipped: " + ex.Message); }
                finally
                {
                    serverCols = await be.GetTableColumns(snap.Table!).ConfigureAwait(false);
                    serverSet = new HashSet<string>(serverCols.Select(c => c.name), StringComparer.OrdinalIgnoreCase);
                    XqlSheetView.InvalidateHeaderCache(snap.SheetName!);
                }
            }

            // 2-2) Alter 후보(type / NOT NULL / CHECK) → 수행
            var desired = BuildDesiredColumnSpec(snap.Meta!, snap.HeaderNames!);
            var alters = InferAlters(serverCols, desired);
            if (alters.Count > 0)
            {
                try { await be.TryAlterColumns(snap.Table!, alters).ConfigureAwait(false); }
                catch (Exception ex) { XqlLog.Warn("AlterColumns skipped: " + ex.Message); }
            }

            // 2-3) Add
            var addTargets = snap.HeaderNames!
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .Where(n => !serverSet.Contains(n))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (addTargets.Count > 0)
            {
                var defs = addTargets.Select(name =>
                {
                    if (!snap.Meta!.Columns.TryGetValue(name, out var ct))
                    {
                        ct = new XqlSheet.ColumnType { Kind = XqlSheet.ColumnKind.Text, Nullable = true };
                        snap.Meta!.SetColumn(name, ct);
                    }
                    return new ColumnDef
                    {
                        Name = name,
                        Kind = ct.Kind switch
                        {
                            XqlSheet.ColumnKind.Int => "integer",
                            XqlSheet.ColumnKind.Real => "real",
                            XqlSheet.ColumnKind.Bool => "bool",
                            XqlSheet.ColumnKind.Date => "integer", // epoch ms
                            XqlSheet.ColumnKind.Json => "json",
                            _ => "text"
                        },
                        NotNull = !ct.Nullable,
                        Check = null
                    };
                }).ToList();

                try { await be.TryAddColumns(snap.Table!, defs).ConfigureAwait(false); }
                catch (Exception ex) { XqlLog.Warn("AddColumns skipped: " + ex.Message); }
                finally { XqlSheetView.InvalidateHeaderCache(snap.SheetName!); }
            }

            // 2-4) Drop
            {
                var metaCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    snap.Key!, "row_version", "updated_at", "deleted"
                };
                var headerSet = new HashSet<string>(snap.HeaderNames!, StringComparer.OrdinalIgnoreCase);

                var drop = serverCols
                    .Where(c => !c.pk && !metaCols.Contains(c.name))
                    .Select(c => c.name?.Trim() ?? "")
                    .Where(n => n.Length > 0)
                    .Where(n => !headerSet.Contains(n))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (drop.Count > 0)
                {
                    try { await be.TryDropColumns(snap.Table!, drop).ConfigureAwait(false); }
                    catch (Exception ex) { XqlLog.Warn("DropColumns skipped: " + ex.Message); }
                    finally { XqlSheetView.InvalidateHeaderCache(snap.SheetName!); }
                }
            }

            // 3) 마지막 매핑 시도 (예외 무음)
            try { XqlSheetView.RegisterTableSheet(snap.Table!, snap.SheetName!); } catch { /* ignore */ }
        }

        // XqlExcelInterop.cs 내부 지역 보조 DTO
        private sealed class SchemaSnapshot
        {
            public bool HasHeader;
            public string? SheetName;
            public string? Table;
            public string? Key;
            public List<string>? HeaderNames;
            public XqlSheet.Meta? Meta;
        }

        // ======== 내부 유틸 ========

        /// <summary>
        /// UsedRange가 갱신되지 않아도 신뢰할 수 있게 마지막 데이터 행을 계산.
        /// - 테이블(ListObject)이 있으면 DataBodyRange 기준
        /// - 없으면 헤더 각 컬럼의 끝에서 위로(End[xlUp]) 스캔하여 최댓값
        /// </summary>
        private static int GetLastDataRow(Excel.Worksheet ws, Excel.Range header, int firstDataRow, int headerColCount)
        {
            if (ws == null || header == null) return firstDataRow - 1;

            // 1) ListObject 우선
            try
            {
                var lo = XqlSheet.FindListObjectContaining(ws, header);
                var body = lo?.DataBodyRange;
                if (body != null)
                {
                    int r = body.Row + body.Rows.Count - 1;
                    return Math.Max(r, firstDataRow - 1);
                }
            }
            catch { /* ignore */ }

            // 2) 각 컬럼별로 End(xlUp) 스캔
            int last = firstDataRow - 1;
            for (int i = 0; i < Math.Max(1, headerColCount); i++)
            {
                try
                {
                    int absCol = header.Column + i;
                    using var lastCell = SmartCom<Excel.Range>.Acquire(() => (Excel.Range)ws.Cells[ws.Rows.Count, absCol]);
                    if (lastCell?.Value == null) continue;

                    using var hit = SmartCom<Excel.Range>.Acquire(() => (Excel.Range)lastCell.Value.End[Excel.XlDirection.xlUp]);
                    if (hit?.Value == null) continue;

                    int candidate = hit.Value.Row;
                    // 헤더행보다 위쪽이면 데이터가 없다는 뜻
                    if (candidate < firstDataRow) continue;

                    if (candidate > last) last = candidate;
                }
                catch { /* ignore per column */ }
            }

            // 3) 그래도 감지 못했으면 UsedRange 보조
            if (last < firstDataRow)
            {
                try
                {
                    using var used = SmartCom<Excel.Range>.Wrap(ws.UsedRange);
                    try { _ = used?.Value?.Address[true, true, Excel.XlReferenceStyle.xlA1, false]; } catch { }
                    int usedLast = (used?.Value?.Row ?? 1) + (used?.Value?.Rows.Count ?? 1) - 1;
                    last = Math.Max(last, usedLast);
                }
                catch { /* ignore */ }
            }

            return last;
        }

        private static List<string> NormalizeHeaderNamesWithLetters(Excel.Range header, IList<string> names)
        {
            if (header == null) throw new ArgumentNullException(nameof(header));
            if (names == null) throw new ArgumentNullException(nameof(names));

            var result = new List<string>(names.Count);
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < names.Count; i++)
            {
                using var cell = SmartCom<Excel.Range>.Acquire(() => (Excel.Range)header.Cells[1, i + 1]);

                // 머지 헤더면 대표셀로 이동
                try
                {
                    if (cell.Value?.MergeCells is bool m && m)
                    {
                        using var ma = SmartCom<Excel.Range>.Wrap(cell.Value.MergeArea);
                        using var rep = SmartCom<Excel.Range>.Acquire(() => (Excel.Range)ma!.Value!.Cells[1, 1]);
                        if (rep?.Value != null)
                        {
                            cell.Detach();
                        }
                    }
                }
                catch { /* ignore */ }

                // ① 기본: 입력값
                string s = (names[i] ?? string.Empty).Trim();

                // ② 값이 비면 표시 텍스트 → 값 → 열문자 순
                if (string.IsNullOrWhiteSpace(s))
                {
                    try { s = Convert.ToString(cell.Value?.Text) ?? ""; } catch { }
                    if (string.IsNullOrWhiteSpace(s))
                    {
                        try { s = Convert.ToString(cell.Value?.Value2) ?? ""; } catch { }
                    }
                    if (string.IsNullOrWhiteSpace(s))
                    {
                        if (cell.Value != null)
                            s = XqlCommon.ColumnIndexToLetter(cell.Value.Column);
                    }
                }

                // ③ 완전 공백 방지
                if (string.IsNullOrWhiteSpace(s) && cell.Value != null)
                    s = XqlCommon.ColumnIndexToLetter(cell.Value.Column);

                // ④ 중복 방지
                var name = (s ?? "").Trim();
                if (used.Contains(name))
                {
                    int n = 2;
                    string candidate;
                    do { candidate = $"{name}_{n++}"; }
                    while (used.Contains(candidate));
                    name = candidate;
                }

                result.Add(name);
                used.Add(name);
            }

            return result;
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

        // 셀 값을 안전하게 가져오는 헬퍼(Value2 → Date/숫자/문자 정규화)
        private static object? GetCell(Excel.Worksheet w, int row, int col)
        {
            using var c = SmartCom<Excel.Range>.Acquire(() => (Excel.Range)w.Cells[row, col]);
            try
            {
                var v = c.Value?.Value2;
                if (v == null) return null;
                if (v is double d && XqlCommon.IsExcelDateTimeLikely(c.Value!))
                    return DateTime.FromOADate(d);
                return v;
            }
            catch { return null; }
        }
    }
}
