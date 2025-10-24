// XqlBackup.cs — SmartCom<T> 적용
using ExcelDna.Integration;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using static XQLite.AddIn.XqlCommon;

namespace XQLite.AddIn
{
    /// <summary>
    /// Export / Recover / Diagnostics 통합 모듈
    /// - ExportDiagnostics: 메타/시트 CSV/서버 메타 & 감사로그를 zip으로 내보내기
    /// - RecoverFromExcel: 현재 워크북을 원본으로 서버 DB 재생성(스키마 보강 + 대량 업서트)
    /// - ExportDb: (가능하면) 서버 풀 덤프 생성 후 zip (스키마 의존, 미지원시 CSV 기반으로 대체)
    /// </summary>
    internal sealed class XqlBackup : IDisposable
    {
        private readonly XqlSheet _sheet;
        private readonly IXqlBackend _backend;

        public XqlBackup(IXqlBackend backend, XqlSheet sheet)
        {
            _sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
            _backend = backend ?? throw new ArgumentNullException(nameof(backend));
        }

        public void Dispose() { /* backend is owned by AddIn */ }

        // ============================================================
        // 1) Diagnostics Export
        //   - meta.json
        //   - sheets/*.csv
        //   - server/meta.json, server/audit_log.json (가능시)
        // ============================================================
        public async Task ExportDiagnostics(string outZipPath)
        {
            try
            {
                var tmp = CreateTempDir("xql_diag_");
                try
                {
                    // meta.json
                    var metaJson = SerializeMeta();
                    File.WriteAllText(Path.Combine(tmp, "meta.json"), metaJson, new UTF8Encoding(false));

                    // sheets/*.csv
                    var sheetsDir = Path.Combine(tmp, "sheets");
                    Directory.CreateDirectory(sheetsDir);
                    ExportAllSheetsCsv(sheetsDir);

                    // server meta/audit (선택)
                    try
                    {
                        var serverDir = Path.Combine(tmp, "server");
                        Directory.CreateDirectory(serverDir);

                        // 1) 서버 메타
                        var sMeta = await _backend.TryFetchServerMeta().ConfigureAwait(false);
                        if (sMeta != null)
                        {
                            var metaText = sMeta.ToString(Formatting.Indented);
                            File.WriteAllText(Path.Combine(serverDir, "meta.json"), metaText, new UTF8Encoding(false));
                        }

                        // 2) 감사 로그 (전체)
                        var audit = await _backend.TryFetchAuditLog(null).ConfigureAwait(false);
                        if (audit != null)
                        {
                            var auditText = audit.ToString(Formatting.Indented);
                            File.WriteAllText(Path.Combine(serverDir, "audit_log.json"), auditText, new UTF8Encoding(false));
                        }
                    }
                    catch { /* 서버 미구현/일시 오류 무시 */ }

                    // zip
                    SafeZipDirectory(tmp, outZipPath);
                }
                finally
                {
                    TryDeleteDir(tmp);
                }
            }
            catch (Exception ex)
            {
                XqlLog.Warn("ExportDiagnostics failed: " + ex.Message);
            }
        }

        // ============================================================
        // 2) Recover
        //   - 원칙: Excel 파일 = 동기화된 DB 원본
        //   - 절차: 스키마 생성/보강 → 배치 업서트 → (신규) id 배정 반영 → 요약
        // ============================================================
        public async Task RecoverFromExcel(int batchSize = 500)
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;

                // 요약 집계 시작
                XqlSheetView.RecoverSummaryBegin();

                foreach (var sheetName in GetWorkbookSheets(app))
                {
                    if (!_sheet.TryGetSheet(sheetName, out var sm)) continue;

                    await EnsureTableSchema(sheetName, sm).ConfigureAwait(false);

                    // 표/헤더 스냅샷 + 행 로딩 (Excel 스레드 접근은 내부에서만)
                    var snap = ReadSheetSnapshot(app, sheetName, sm);
                    if (snap.Rows.Count == 0) continue;

                    var table = string.IsNullOrWhiteSpace(sm.TableName) ? sheetName : sm.TableName!;

                    foreach (var chunkRows in XqlCommon.Chunk(snap.Rows, batchSize))
                    {
                        // 행들을 Cell 단위 편평화(서버 upsertCells 사용)
                        var (edits, tempMap) = RowsToCellEdits(table, sm, chunkRows);

                        var res = await _backend.UpsertCells(edits).ConfigureAwait(false);

                        // 충돌 로그 적재
                        if (res?.Conflicts is { Count: > 0 })
                            XqlSheetView.AppendConflicts(res.Conflicts.Cast<object>());

                        // 신규 id 배정 반영
                        if (res?.Assigned is { Count: > 0 })
                        {
                            try
                            {
                                await OnExcelThreadAsync(() =>
                                {
                                    using var ws = SmartCom<Excel.Worksheet>.Wrap(XqlSheet.FindWorksheet(app, sheetName));
                                    if (ws?.Value == null) return 0;

                                    foreach (var a in res.Assigned)
                                    {
                                        if (!string.Equals(a.Table, table, StringComparison.Ordinal)) continue;
                                        if (string.IsNullOrEmpty(a.NewId) || string.IsNullOrEmpty(a.TempRowKey)) continue;
                                        if (!tempMap.TryGetValue(a.TempRowKey!, out var excelRow)) continue;

                                        var keyAbsCol = snap.HeaderColumn + snap.KeyIndex1 - 1;
                                        using var keyCell = SmartCom<Excel.Range>.Acquire(() => (Excel.Range)ws.Value.Cells[excelRow, keyAbsCol]);
                                        if (keyCell?.Value != null) keyCell.Value.Value2 = a.NewId;
                                    }
                                    return 0;
                                });
                            }
                            catch { /* 시트가 닫혔거나 사용자가 이동한 경우 무시 */ }
                        }

                        // 요약
                        int conflicts = res?.Conflicts?.Count ?? 0;
                        int errors = res?.Errors?.Count ?? 0;
                        int affected = edits.Count; // 업서트한 '셀' 개수
                        XqlSheetView.RecoverSummaryPush(table, affected, conflicts, errors);
                    }
                }

                XqlSheetView.RecoverSummaryShow();
            }
            catch
            {
                // 무음(사용자는 Inspector/Diag로 확인 가능)
            }
        }

        // ============================================================
        // 3) Export DB (가능한 경우)
        //   - 서버 풀 덤프가 없다면 CSV 기반 진단 zip과 동일하게 동작
        // ============================================================
        public async Task ExportDb(string outZipPath)
        {
            try
            {
                var tmp = CreateTempDir("xql_export_");
                try
                {
                    // 1) 서버가 풀 덤프를 지원하면 그 결과를 그대로 보관
                    var dbBytes = await _backend.TryExportDatabase().ConfigureAwait(false);
                    if (dbBytes != null)
                    {
                        var dbPath = Path.Combine(tmp, "database.sqlite");
                        File.WriteAllBytes(dbPath, dbBytes);
                    }

                    // 2) 항상 CSV/메타도 함께 내보내기(사람이 열람 가능)
                    var metaJson = SerializeMeta();
                    File.WriteAllText(Path.Combine(tmp, "meta.json"), metaJson, new UTF8Encoding(false));
                    var sheetsDir = Path.Combine(tmp, "sheets");
                    Directory.CreateDirectory(sheetsDir);
                    ExportAllSheetsCsv(sheetsDir);

                    SafeZipDirectory(tmp, outZipPath);
                }
                finally
                {
                    TryDeleteDir(tmp);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("ExportDb failed: " + ex.Message);
            }
        }

        // ============================================================
        // 내부: 스키마/행/셀 변환 유틸
        // ============================================================
        private static string MapKind(XqlSheet.ColumnKind k) => k switch
        {
            XqlSheet.ColumnKind.Int => "integer",
            XqlSheet.ColumnKind.Real => "real",
            XqlSheet.ColumnKind.Bool => "bool",
            XqlSheet.ColumnKind.Date => "integer", // epoch ms
            XqlSheet.ColumnKind.Json => "json",
            _ => "text"
        };

        private async Task EnsureTableSchema(string sheetName, XqlSheet.Meta sm)
        {
            // 테이블명/키 기본값 보강
            var table = string.IsNullOrWhiteSpace(sm.TableName) ? sheetName : sm.TableName!;
            var key = string.IsNullOrWhiteSpace(sm.KeyColumn) ? "id" : sm.KeyColumn!;
            sm.TableName = table;
            sm.KeyColumn = key;

            await _backend.TryCreateTable(table, key).ConfigureAwait(false);

            // 현재 메타 기준 컬럼 정의(없으면 TEXT NULL)
            var defs = sm.Columns.Select(kv => new ColumnDef
            {
                Name = kv.Key,
                Kind = MapKind(kv.Value.Kind),
                NotNull = !kv.Value.Nullable,
                Check = null
            }).ToList();

            // 빈 정의일 수 있어도 호출은 안전(서버에서 무시)
            await _backend.TryAddColumns(table, defs).ConfigureAwait(false);
        }

        // 스냅샷 DTO(헤더 좌표/키 인덱스 포함)
        private sealed class SheetSnapshot
        {
            public int HeaderRow;
            public int HeaderColumn;
            public int HeaderCols;
            public int KeyIndex1;
            public List<string> HeaderNames = new();
            public List<RowData> Rows = new();
        }

        private sealed class RowData
        {
            public int ExcelRow;
            public Dictionary<string, object?> Values = new(StringComparer.Ordinal);
        }

        /// <summary>
        /// 시트 헤더/키 인덱스/행 데이터를 한 번에 스냅샷(모두 순수 데이터만 반환; COM 미보관)
        /// </summary>
        private SheetSnapshot ReadSheetSnapshot(Excel.Application app, string sheetName, XqlSheet.Meta sm)
        {
            var snap = new SheetSnapshot();

            using var ws = SmartCom<Excel.Worksheet>.Wrap(XqlSheet.FindWorksheet(app, sheetName));
            if (ws?.Value == null) return snap;

            using var used = SmartCom<Excel.Range>.Wrap(ws.Value.UsedRange);

            // 헤더 탐색: 마커 → 표 헤더 → 폴백
            using var header = SmartCom<Excel.Range>.Acquire(() =>
            {
                if (XqlSheet.TryGetHeaderMarker(ws.Value, out var mk)) return mk;
                Excel.ListObject? lo = null;
                try { lo = ws.Value.ListObjects?.Count > 0 ? ws.Value.ListObjects[1] : null; }
                catch { lo = null; }
                if (lo?.HeaderRowRange != null) return lo.HeaderRowRange;
                return XqlSheet.GetHeaderRange(ws.Value);
            });

            if (header?.Value == null) return snap;

            // UsedRange 끝 행 계산 (Value 접근 전 Address-touch는 SmartCom 필요 없음)
            int lastRow = 0;
            try { lastRow = used?.Value != null ? (used.Value.Row + used.Value.Rows.Count - 1) : (ws.Value.UsedRange.Row + ws.Value.UsedRange.Rows.Count - 1); }
            catch { lastRow = ws.Value.UsedRange.Row + ws.Value.UsedRange.Rows.Count - 1; }

            snap.HeaderRow = header.Value.Row;
            snap.HeaderColumn = header.Value.Column;
            snap.HeaderCols = header.Value.Columns.Count;

            // 헤더명 (Value2 배열 우선)
            var hv = header.Value.Value2 as object[,];
            if (hv != null)
            {
                for (int i = 1; i <= snap.HeaderCols; i++)
                {
                    var nm = (Convert.ToString(hv[1, i]) ?? "").Trim();
                    snap.HeaderNames.Add(string.IsNullOrEmpty(nm)
                        ? ColumnIndexToLetter(snap.HeaderColumn + i - 1)
                        : nm);
                }
            }
            else
            {
                for (int i = 1; i <= snap.HeaderCols; i++)
                {
                    using var hc = SmartCom<Excel.Range>.Acquire(() => (Excel.Range)header.Value.Cells[1, i]);
                    var nm = (hc?.Value?.Value2 as string)?.Trim();
                    snap.HeaderNames.Add(string.IsNullOrEmpty(nm)
                        ? ColumnIndexToLetter(snap.HeaderColumn + i - 1)
                        : nm!);
                }
            }

            // 키 인덱스(1-based)
            var keyName = string.IsNullOrWhiteSpace(sm.KeyColumn) ? "id" : sm.KeyColumn!;
            snap.KeyIndex1 = XqlSheet.FindKeyColumnIndex(snap.HeaderNames, keyName);
            if (snap.KeyIndex1 <= 0) snap.KeyIndex1 = 1;

            // 데이터 행
            int firstDataRow = snap.HeaderRow + 1;
            int last = Math.Max(firstDataRow, lastRow);

            for (int r = firstDataRow; r <= last; r++)
            {
                var row = new RowData { ExcelRow = r };
                bool any = false;

                for (int i = 1; i <= snap.HeaderCols; i++)
                {
                    var colName = snap.HeaderNames[i - 1];
                    if (string.IsNullOrWhiteSpace(colName)) continue;
                    if (!sm.Columns.ContainsKey(colName)) continue; // 메타에 없는 컬럼은 무시

                    using var cell = SmartCom<Excel.Range>.Acquire(() => (Excel.Range)ws.Value.Cells[r, snap.HeaderColumn + i - 1]);
                    object? v = null;
                    try { v = cell?.Value?.Value2; } catch { }

                    row.Values[colName] = v;
                    if (v != null && !(v is string s && string.IsNullOrWhiteSpace(s))) any = true;
                }

                if (any) snap.Rows.Add(row);
            }

            return snap;
        }

        /// <summary>
        /// 행 묶음을 Cell 편평화로 변환. tempKey("-<ExcelRow>") 매핑을 반환해서
        /// 서버 Assigned → 시트 id 반영에 사용.
        /// </summary>
        private static (List<EditCell> Cells, Dictionary<string, int> TempKeyToRow)
            RowsToCellEdits(string table, XqlSheet.Meta sm, IEnumerable<RowData> rows)
        {
            var cells = new List<EditCell>();
            var map = new Dictionary<string, int>(StringComparer.Ordinal);

            string keyName = string.IsNullOrWhiteSpace(sm.KeyColumn) ? "id" : sm.KeyColumn!;

            foreach (var rd in rows)
            {
                var r = rd.Values;

                // RowKey 결정: key 컬럼 값 있으면 그걸, 없으면 tempKey 사용
                object? keyVal = null;
                if (!string.IsNullOrWhiteSpace(keyName) && r.TryGetValue(keyName, out var pk) && pk != null)
                    keyVal = pk;

                string? tempKey = null;
                if (keyVal == null)
                {
                    tempKey = "-" + rd.ExcelRow.ToString();
                    keyVal = tempKey;
                    map[tempKey] = rd.ExcelRow;
                }

                foreach (var kv in r)
                {
                    var col = kv.Key;
                    if (string.Equals(col, keyName, StringComparison.OrdinalIgnoreCase))
                        continue; // PK 컬럼은 업서트 셀에서 제외

                    cells.Add(new EditCell(table, keyVal!, col, kv.Value));
                }
            }

            return (cells, map);
        }

        private string SerializeMeta()
        {
            var meta = new Dictionary<string, object?>(StringComparer.Ordinal);
            // 시트별 메타
            var sheets = new List<object>();
            foreach (var name in GetAllRegisteredSheets())
            {
                if (!_sheet.TryGetSheet(name, out var sm)) continue;
                var cols = sm.Columns.Select(kv => new
                {
                    name = kv.Key,
                    kind = kv.Value.Kind.ToString(),
                    nullable = kv.Value.Nullable,
                }).ToList();

                sheets.Add(new
                {
                    sheet = name,
                    table = string.IsNullOrWhiteSpace(sm.TableName) ? name : sm.TableName,
                    key = string.IsNullOrWhiteSpace(sm.KeyColumn) ? "id" : sm.KeyColumn,
                    columns = cols
                });
            }
            meta["sheets"] = sheets;
            return JsonConvert.SerializeObject(meta, Formatting.Indented);
        }

        private void ExportAllSheetsCsv(string outDir)
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            foreach (var sheetName in GetWorkbookSheets(app))
            {
                if (!_sheet.TryGetSheet(sheetName, out var sm)) continue;
                var snap = ReadSheetSnapshot(app, sheetName, sm);
                var rows = snap.Rows.Select(r => r.Values).ToList();
                var outPath = Path.Combine(outDir, $"{SafeFileName(sheetName)}.csv");
                WriteCsv(outPath, rows);
            }
        }

        private static IEnumerable<string> GetWorkbookSheets(Excel.Application app)
        {
            var list = new List<string>();
            using var sheets = SmartCom<Excel.Sheets>.Wrap(app.Worksheets);
            try
            {
                int count = sheets?.Value?.Count ?? 0;
                for (int i = 1; i <= count; i++)
                {
                    using var w = SmartCom<Excel.Worksheet>.Acquire(() => (Excel.Worksheet?)sheets!.Value![i]);
                    if (w?.Value != null) list.Add(w.Value.Name);
                }
            }
            catch { /* ignore */ }
            return list;
        }

        private IEnumerable<string> GetAllRegisteredSheets()
        {
            // 워크북 시트명과 메타 매칭
            var app = (Excel.Application)ExcelDnaUtil.Application;
            foreach (var name in GetWorkbookSheets(app))
                if (_sheet.TryGetSheet(name, out _)) yield return name;
        }

        private static void WriteCsv(string path, List<Dictionary<string, object?>> rows)
        {
            try
            {
                if (rows.Count == 0) { File.WriteAllText(path, "", new UTF8Encoding(false)); return; }

                // 헤더: 모든 키의 합집합
                var headers = rows.SelectMany(r => r.Keys).Distinct(StringComparer.Ordinal).ToList();
                using var sw = new StreamWriter(path, false, new UTF8Encoding(false));
                sw.WriteLine(string.Join(",", headers.Select(CsvEscape)));

                foreach (var r in rows)
                {
                    var line = string.Join(",", headers.Select(h => CsvEscape(ValueToString(r.TryGetValue(h, out var v) ? v : null))));
                    sw.WriteLine(line);
                }
            }
            catch { /* ignore */ }
        }

        private static string SafeFileName(string name)
        {
            var invalid = Path.GetInvalidFileNameChars();
            var sb = new StringBuilder(name.Length);
            foreach (var ch in name)
                sb.Append(invalid.Contains(ch) ? '_' : ch);
            return sb.ToString();
        }
    }
}
