// RowSyncService.cs
// - 기존 XqlUpsert(디바운스 큐 + NDJSON 아웃박스)를 직접 사용
// - 흐름: 헤더/메타 보강 → 데이터 범위 → 행 수집(Dictionary) → Enqueue → FlushAsync
// - 서버가 row_version/updated_at을 즉시 돌려주지 않으므로, 커밋 직후 셀 갱신은 하지 않음
//   (다음 스텝: Pull(since_version)에서 최신 메타를 병합)

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

#if false
namespace XQLite.AddIn
{
    internal static class XqlRowSyncService
    {
        private const string COL_ID = "id";
        private const string COL_ROW_VERSION = "row_version";
        private const string COL_UPDATED_AT = "updated_at";
        private const string COL_DELETED = "deleted";

        /// <summary>리본 Commit 버튼에서 호출</summary>
        public static void CommitActiveSheet()
        {
            var app = (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
            var ws = app?.ActiveSheet as Excel.Worksheet;
            if (ws == null) return;

            var meta = XqlSheetMetaRegistry.Get(ws);
            if (meta == null)
            {
                System.Windows.Forms.MessageBox.Show("메타 헤더가 없습니다. 먼저 헤더를 설치하세요.");
                return;
            }

            // 헤더 Range & 헤더명 로드
            var header = ws.Range[
                ws.Cells[meta.TopRow, meta.LeftCol],
                ws.Cells[meta.TopRow, meta.LeftCol + Math.Max(1, meta.ColCount) - 1]
            ];
            var headers = ReadHeaderNames(header);

            // 메타 컬럼 보강(id,row_version,updated_at,deleted)
            EnsureMetaColumns(ws, header, headers);

            // 데이터 범위 찾기
            var dataRange = GetDataRange(ws, header, headers.Count);
            if (dataRange == null)
            {
                System.Windows.Forms.MessageBox.Show("데이터가 없습니다. 편집 후 Commit 하세요.");
                return;
            }

            // 새 행 id / updated_at 기본값 채우기(비어있을 때만)
            EnsureRowIdentity(ws, dataRange, headers);

            // 행 수집
            var rows = CollectRows(ws, dataRange, headers);
            if (rows.Count == 0)
            {
                System.Windows.Forms.MessageBox.Show("업서트할 데이터가 없습니다.");
                return;
            }

            // 큐 적재 + 즉시 Flush
            string table = ws.Name;
            try
            {
                foreach (var r in rows)
                    XqlUpsert.Enqueue(table, r);

                // 리본/COM 컨텍스트에서 안전하게 동기 대기
                XqlUpsert.FlushAsync().GetAwaiter().GetResult();

                // 여기서는 서버 메타 즉시 반영 X (다음 스텝에서 Pull로 메타 병합)
                System.Windows.Forms.MessageBox.Show($"Commit 큐 전송 완료: {rows.Count} row(s).");
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Commit 실패: " + ex.Message);
            }
        }

        // ---------- Header & Meta ----------

        private static List<string> ReadHeaderNames(Excel.Range header)
        {
            int cols = 1; try { cols = Math.Max(1, Convert.ToInt32(header.Columns.Count)); } catch { cols = 1; }
            var list = new List<string>(cols);
            for (int i = 1; i <= cols; i++)
            {
                string name = "";
                try { name = Convert.ToString((header.Cells[1, i] as Excel.Range)?.Value2) ?? ""; } catch { }
                list.Add((name ?? "").Trim());
            }
            return list;
        }

        /// <summary>필수 메타 컬럼이 없으면 헤더 우측에 추가하고, 헤더 스타일/검증을 갱신</summary>
        private static void EnsureMetaColumns(Excel.Worksheet ws, Excel.Range header, List<string> headers)
        {
            bool hasId = headers.Any(h => string.Equals(h, COL_ID, StringComparison.OrdinalIgnoreCase));
            bool hasRowVersion = headers.Any(h => string.Equals(h, COL_ROW_VERSION, StringComparison.OrdinalIgnoreCase));
            bool hasUpdatedAt = headers.Any(h => string.Equals(h, COL_UPDATED_AT, StringComparison.OrdinalIgnoreCase));
            bool hasDeleted = headers.Any(h => string.Equals(h, COL_DELETED, StringComparison.OrdinalIgnoreCase));

            var baseRow = header.Row;
            var startCol = header.Column;
            int colCount = headers.Count;

            int added = 0;
            void AddCol(string name)
            {
                var c = startCol + colCount + added;
                var cell = ws.Cells[baseRow, c] as Excel.Range;
                cell?.Value2 = name;
                added++;
            }

            if (!hasId) AddCol(COL_ID);
            if (!hasRowVersion) AddCol(COL_ROW_VERSION);
            if (!hasUpdatedAt) AddCol(COL_UPDATED_AT);
            if (!hasDeleted) AddCol(COL_DELETED);

            if (added > 0)
            {
                // 헤더 스타일/폭 재계산
                XqlSheetMetaRegistry.RefreshHeaderBorders(ws);

                // 새로 추가한 메타 열 타입/검증(선택)
                try
                {
                    var newHeader = ws.Range[ws.Cells[baseRow, startCol], ws.Cells[baseRow, startCol + colCount + added - 1]];
                    for (int i = colCount + 1; i <= colCount + added; i++)
                    {
                        var cell = newHeader.Cells[1, i] as Excel.Range; if (cell == null) continue;
                        string nm = Convert.ToString(cell.Value2)?.Trim().ToLowerInvariant() ?? "";
                        if (nm == COL_ID) XqlColumnTypeRegistry.SetColumnType(ws, cell, "TEXT");
                        else if (nm == COL_ROW_VERSION) XqlColumnTypeRegistry.SetColumnType(ws, cell, "INT");
                        else if (nm == COL_UPDATED_AT) XqlColumnTypeRegistry.SetColumnType(ws, cell, "DATE");
                        else if (nm == COL_DELETED) XqlColumnTypeRegistry.SetColumnType(ws, cell, "BOOL");
                    }
                }
                catch { }

                // headers 최신화
                headers.Clear();
                var hdr2 = ws.Range[ws.Cells[baseRow, startCol], ws.Cells[baseRow, startCol + colCount + added - 1]];
                headers.AddRange(ReadHeaderNames(hdr2));
            }
        }

        // ---------- Data range ----------

        private static Excel.Range? GetDataRange(Excel.Worksheet ws, Excel.Range header, int headerColCount)
        {
            int firstRow = header.Row + 1;
            int firstCol = header.Column;
            int lastCol = firstCol + headerColCount - 1;

            int lastRow = firstRow - 1;

            int usedLastRow;
            try { usedLastRow = Math.Max(ws.UsedRange?.Row + ws.UsedRange?.Rows?.Count - 1 ?? firstRow, firstRow); }
            catch { usedLastRow = firstRow + 5000; }

            for (int r = usedLastRow; r >= firstRow; r--)
            {
                bool any = false;
                for (int c = firstCol; c <= lastCol; c++)
                {
                    var cell = ws.Cells[r, c] as Excel.Range;
                    var v = cell?.Value2;
                    if (v != null && !(v is string s && string.IsNullOrWhiteSpace(s)))
                    {
                        any = true; break;
                    }
                }
                if (any) { lastRow = r; break; }
            }

            if (lastRow < firstRow) return null;
            return ws.Range[ws.Cells[firstRow, firstCol], ws.Cells[lastRow, lastCol]];
        }

        // ---------- Identity defaults ----------

        private static void EnsureRowIdentity(Excel.Worksheet ws, Excel.Range data, List<string> headers)
        {
            int rows = 1; try { rows = Math.Max(1, Convert.ToInt32(data.Rows.Count)); } catch { }
            int colId = IndexOf(headers, COL_ID);
            int colUpdatedAt = IndexOf(headers, COL_UPDATED_AT);

            for (int r = 1; r <= rows; r++)
            {
                if (colId >= 0)
                {
                    var cell = data.Cells[r, colId + 1] as Excel.Range;
                    try
                    {
                        if (cell != null)
                        {
                            var v = cell.Value2;
                            if (v == null || (v is string s && string.IsNullOrWhiteSpace(s)))
                                cell.Value2 = Guid.NewGuid().ToString("N");
                        }
                    }
                    catch { }
                }

                if (colUpdatedAt >= 0)
                {
                    var cell = data.Cells[r, colUpdatedAt + 1] as Excel.Range;
                    try
                    {
                        if (cell != null)
                        {
                            var v = cell.Value2;
                            if (v == null || (v is string s && string.IsNullOrWhiteSpace(s)))
                                cell.Value2 = DateTime.UtcNow;
                        }
                    }
                    catch { }
                }
            }
        }

        // ---------- Collect ----------

        private static List<Dictionary<string, object?>> CollectRows(Excel.Worksheet ws, Excel.Range data, List<string> headers)
        {
            int rows = 1, cols = 1;
            try { rows = Math.Max(1, Convert.ToInt32(data.Rows.Count)); } catch { }
            try { cols = Math.Max(1, Convert.ToInt32(data.Columns.Count)); } catch { }

            var list = new List<Dictionary<string, object?>>(rows);

            for (int r = 1; r <= rows; r++)
            {
                var obj = new Dictionary<string, object?>(cols, StringComparer.OrdinalIgnoreCase);
                bool any = false;

                for (int c = 1; c <= cols; c++)
                {
                    string key = headers[c - 1];
                    if (string.IsNullOrWhiteSpace(key)) continue;

                    var cell = data.Cells[r, c] as Excel.Range;
                    object? val = null;
                    try { val = cell?.Value2; } catch { }

                    // 날짜형 열 추정 → OADate 변환 시도
                    if (val is double d && LooksLikeDateColumn(key))
                    {
                        try { val = DateTime.FromOADate(d); } catch { }
                    }

                    // 공백 문자열 → null
                    if (val is string s && string.IsNullOrWhiteSpace(s)) val = null;

                    if (val != null) any = true;
                    obj[key] = val;
                }

                if (any)
                    list.Add(obj);
            }

            return list;
        }

        private static bool LooksLikeDateColumn(string key)
        {
            key = (key ?? "").ToLowerInvariant();
            return key.Contains("date") || key.Contains("at") || key.EndsWith("_ts");
        }

        private static int IndexOf(List<string> headers, string name)
        {
            for (int i = 0; i < headers.Count; i++)
                if (string.Equals(headers[i], name, StringComparison.OrdinalIgnoreCase)) return i;
            return -1;
        }
    }
}
#endif