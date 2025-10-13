// XqlBackup.cs
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;
using Excel = Microsoft.Office.Interop.Excel;

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
            _backend = backend;
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
                var tmp = XqlCommon.CreateTempDir("xql_diag_");
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

                        // 2) 감사 로그 (전체 스냅샷)
                        var audit = await _backend.TryFetchAuditLog(null).ConfigureAwait(false);
                        if (audit != null)
                        {
                            var auditText = audit.ToString(Formatting.Indented);
                            File.WriteAllText(Path.Combine(serverDir, "audit_log.json"), auditText, new UTF8Encoding(false));
                        }
                    }
                    catch { /* 서버가 미구현이어도 무시 */ }

                    // zip
                    XqlCommon.SafeZipDirectory(tmp, outZipPath);
                }
                finally
                {
                    XqlCommon.TryDeleteDir(tmp);
                }
            }
            catch (Exception ex)
            {
                // 실패는 무음 처리(Excel 안정성 우선), 로그 시트에만 남김
                XqlLog.Warn("ExportDiagnostics failed: " + ex.Message);
            }
        }

        // ============================================================
        // 2) Recover
        //   - 원칙: Excel 파일 = 동기화된 DB 원본
        //   - 절차: 스키마 생성/보강 → 배치 업서트 → 무결성 검사(서버 쪽) → 완료
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
                    await EnsureTableSchema(sm).ConfigureAwait(false);

                    var rows = ReadSheetRows(app, sheetName, sm);
                    if (rows.Count == 0) continue;
                    var table = sm.TableName ?? sheetName; // 요약/충돌 기록에 사용

                    foreach (var chunk in XqlCommon.Chunk(RowsToCellEdits(table, sm, rows), batchSize))
                    {
                        var res = await _backend.UpsertCells(chunk).ConfigureAwait(false);
                        // 충돌이 있으면 Conflict 워크시트에 적재
                        if (res?.Conflicts != null && res.Conflicts.Count > 0)
                            XqlSheetView.AppendConflicts(res.Conflicts.Cast<object>());
                        // 요약 집계(행/셀 기준은 운영 목적에 맞춰 조정 가능. 여기선 '셀 건수'로 반영)
                        var conflicts = res?.Conflicts?.Count ?? 0;
                        var errors = res?.Errors?.Count ?? 0;
                        var affected = chunk.Count; // 업서트한 셀 개수
                        XqlSheetView.RecoverSummaryPush(table, affected, conflicts, errors);
                    }
                }

                XqlSheetView.RecoverSummaryShow();
            }
            catch { /* 무음 실패 (UI는 Inspector/Diag로 확인) */ }
        }

        // ============================================================
        // 3) Export DB (가능한 경우)
        //   - 서버 풀 덤프가 없다면 CSV 기반 진단 zip과 동일하게 동작
        // ============================================================
        public async Task ExportDb(string outZipPath)
        {
            try
            {
                var tmp = XqlCommon.CreateTempDir("xql_export_");
                try
                {
                    // 1) 서버가 풀 덤프를 지원하면 그 결과를 그대로 보관
                    var dbBytes = await _backend.TryExportDatabase();
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

                    XqlCommon.SafeZipDirectory(tmp, outZipPath);
                }
                finally
                {
                    XqlCommon.TryDeleteDir(tmp);
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
        private async Task EnsureTableSchema(XqlSheet.Meta sm)
        {
            await _backend.TryCreateTable(sm.TableName, sm.KeyColumn);
            var defs = sm.Columns.Select(kv => new ColumnDef
            {
                Name = kv.Key,
                Kind = kv.Value.Kind.ToString().ToLowerInvariant(),
                NotNull = !kv.Value.Nullable,
                Check = null
            }).ToList();
            await _backend.TryAddColumns(sm.TableName, defs);
        }

        private static List<Dictionary<string, object?>> ReadSheetRows(Excel.Application app, string sheetName, XqlSheet.Meta sm)
        {
            var list = new List<Dictionary<string, object?>>();
            Excel.Worksheet? ws = null; Excel.Range? used = null; Excel.Range? header = null; Excel.ListObject? lo = null;
            try
            {
                ws = app.Worksheets.Cast<Excel.Worksheet>()
                    .FirstOrDefault(s => string.Equals(s.Name, sheetName, StringComparison.Ordinal));
                if (ws == null) return list;

                used = ws.UsedRange;

                int usedLastRow = used.Row + used.Rows.Count - 1;
                XqlCommon.ReleaseCom(used); used = null;

                // 1) 헤더 탐색: 마커 → 표 헤더 → Fallback
                if (!XqlSheet.TryGetHeaderMarker(ws, out header))
                {
                    lo = ws.ListObjects?.Count > 0 ? ws.ListObjects[1] : null;
                    if (lo?.HeaderRowRange != null) header = lo.HeaderRowRange;
                    if (header == null) header = XqlSheet.GetHeaderRange(ws);
                }
                if (header == null) return list;


                // 2) 헤더 수집 (배열 경로)
                var headers = new List<string>(header.Columns.Count);
                var hv = header.Value2 as object[,];
                if (hv != null)
                {
                    int cols = header.Columns.Count;
                    for (int i = 1; i <= cols; i++)
                    {
                        var nm = (Convert.ToString(hv[1, i]) ?? "").Trim();
                        headers.Add(string.IsNullOrEmpty(nm) ? XqlCommon.ColumnIndexToLetter(header.Column + i - 1) : nm);
                    }
                }
                else
                {
                    for (int i = 1; i <= header.Columns.Count; i++)
                    {
                        Excel.Range? hc = null;
                        try
                        {
                            hc = (Excel.Range)header.Cells[1, i];
                            var nm = (hc.Value2 as string)?.Trim();
                            headers.Add(string.IsNullOrEmpty(nm) ? XqlCommon.ColumnIndexToLetter(header.Column + i - 1) : nm!);
                        }
                        finally { XqlCommon.ReleaseCom(hc); }
                    }
                }

                // 3) 데이터 행: header 바로 아래부터 UsedRange 끝까지
                int firstDataRow = header.Row + 1;
                int lastRow = Math.Max(firstDataRow, usedLastRow);
                for (int r = firstDataRow; r <= lastRow; r++)
                {
                    var row = new Dictionary<string, object?>(StringComparer.Ordinal);
                    bool any = false;
                    for (int i = 1; i <= header.Columns.Count; i++)
                    {
                        var key = headers[i - 1];
                        if (string.IsNullOrWhiteSpace(key)) continue;
                        if (!sm.Columns.ContainsKey(key)) continue; // 메타에 없는 컬럼 skip

                        Excel.Range? cell = null;
                        object? v = null;
                        try { cell = (Excel.Range)ws.Cells[r, header.Column + i - 1]; v = cell.Value2; }
                        catch { }
                        finally { XqlCommon.ReleaseCom(cell); }
                        // Excel 정수=double → 그대로 보관 (서버가 형변환)
                        row[key] = v;
                        if (v != null && !(v is string s && string.IsNullOrWhiteSpace(s))) any = true;
                    }
                    if (any) list.Add(row);
                }
            }
            catch { }
            finally
            {
                XqlCommon.ReleaseCom(header, lo, ws);
            }
            return list;
        }

        private static List<EditCell> RowsToCellEdits(string table, XqlSheet.Meta sm, List<Dictionary<string, object?>> rows)
        {
            var cells = new List<EditCell>(rows.Count * 4);
            for (int i = 0; i < rows.Count; i++)
            {
                var r = rows[i];
                // 행 키는 "id" 또는 "key" 또는 1열 값을 우선 사용 (여기선 id/key 우선)
                object rowKey =
                (sm.KeyColumn is { Length: > 0 } && r.TryGetValue(sm.KeyColumn, out var pk) && pk != null) ? pk :
                (r.TryGetValue("id", out var idv) && idv != null ? idv :
                r.TryGetValue("key", out var kv) && kv != null ? kv :
                i + 1);

                foreach (var kvp in r)
                {
                    cells.Add(new EditCell(table, rowKey, kvp.Key, kvp.Value));
                }
            }
            return cells;
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
                    table = sm.TableName,
                    key = sm.KeyColumn,
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
                var rows = ReadSheetRows(app, sheetName, sm);
                var outPath = Path.Combine(outDir, $"{SafeFileName(sheetName)}.csv");
                WriteCsv(outPath, rows);
            }
        }

        private static IEnumerable<string> GetWorkbookSheets(Excel.Application app)
        {
            var list = new List<string>();
            try
            {
                foreach (Excel.Worksheet w in app.Worksheets)
                {
                    try { list.Add(w.Name); }
                    finally { XqlCommon.ReleaseCom(w); }
                }
            }
            catch { }
            return list;
        }

        private IEnumerable<string> GetAllRegisteredSheets()
        {
            // XqlMetaRegistry 내부 사전을 직접 노출하지 않으므로, CSV 내보내기는 워크북 기준으로,
            // 메타 serialize 는 TryGetSheet로 가능한 이름만 포함.
            // 여기서는 워크북 시트명과 메타 매칭을 시도.
            var app = (Excel.Application)ExcelDnaUtil.Application;
            foreach (var name in GetWorkbookSheets(app))
                if (_sheet.TryGetSheet(name, out _)) yield return name;
        }

        private static void WriteCsv(string path, List<Dictionary<string, object?>> rows)
        {
            try
            {
                if (rows.Count == 0) { File.WriteAllText(path, "", new UTF8Encoding(false)); return; }

                // 헤더: 모든 키의 합집합 (메타로 제한되어 있지만 혹시 모를 차이를 위해 합집합)
                var headers = rows.SelectMany(r => r.Keys).Distinct(StringComparer.Ordinal).ToList();
                using var sw = new StreamWriter(path, false, new UTF8Encoding(false));
                sw.WriteLine(string.Join(",", headers.Select(XqlCommon.CsvEscape)));

                foreach (var r in rows)
                {
                    var line = string.Join(",", headers.Select(h => XqlCommon.CsvEscape(XqlCommon.ValueToString(r.TryGetValue(h, out var v) ? v : null))));
                    sw.WriteLine(line);
                }
            }
            catch { }
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
