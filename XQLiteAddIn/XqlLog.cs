// XqlLog.cs
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Diagnostics;
using System.Net.NetworkInformation;
using System.Text.RegularExpressions;
using XQLite.AddIn;
using Excel = Microsoft.Office.Interop.Excel;

namespace XQLite.AddIn
{
    internal static class XqlLog
    {
        public static void Info(string msg, string? sheet = "*", string? address = null)
        {
            Log("INFO", msg, sheet, address);
            try { XqlFileLogger.Write("INFO", sheet!, msg); } catch { }
        }
        public static void Warn(string msg, string? sheet = "*", string? address = null)
        {
            Log("WARN", msg, sheet, address);
            try { XqlFileLogger.Write("WARN", sheet!, msg); } catch { }
        }
        public static void Error(string msg, string? sheet = "*", string? address = null)
        {
            Log("ERROR", msg, sheet, address);
            try { XqlFileLogger.Write("ERR", sheet!, msg); } catch { }
        }

        // ─────────────────────────────────────────────────────────────────
        //  Log: 워크북에 "_XQL_Log" 시트를 만들어 기록 (UI 스레드에서만 작업)
        //  컬럼: Timestamp | Level | Sheet | Address | Message
        // ─────────────────────────────────────────────────────────────────

        public static void Log(string level, string msg, string? sheet, string? address)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
                Excel.Workbook? wb = null; Excel.Worksheet? ws = null;
                Excel.Range? cell = null; Excel.Range? row = null;
                try
                {
                    wb = app.ActiveWorkbook;
                    if (wb == null) return;

                    // 시트 찾기/생성
                    ws = FindOrCreateLogSheet(wb, "_XQL_Log");

                    // 헤더

                    // UsedRange는 호출 시점 스냅샷. 헤더 작성 전 기준으로 next를 계산해도
                    // 최소 2행부터 쓰도록 Max 처리되어 안전함.
                    Excel.Range? ur = ws.UsedRange as Excel.Range;
                    if (((ur?.Rows?.Count) ?? 0) <= 1 && ((ws.Cells[1, 1] as Excel.Range)?.Value2 == null))
                    {
                        (ws.Cells[1, 1] as Excel.Range)!.Value2 = "Timestamp";
                        (ws.Cells[1, 2] as Excel.Range)!.Value2 = "Level";
                        (ws.Cells[1, 3] as Excel.Range)!.Value2 = "Sheet";
                        (ws.Cells[1, 4] as Excel.Range)!.Value2 = "Address";
                        (ws.Cells[1, 5] as Excel.Range)!.Value2 = "Message";
                    }
                    int last = (ur?.Row ?? 1) + Math.Max(0, (ur?.Rows?.Count ?? 1) - 1);
                    int next = Math.Max(2, last + 1);
                    XqlCommon.ReleaseCom(ur); ur = null;

#pragma warning disable CS8602 // null 가능 참조에 대한 역참조입니다.
                    row = ws.Range[ws.Cells[next, 1], ws.Cells[next, 5]];
                    (row.Cells[1, 1] as Excel.Range).Value2 = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                    (row.Cells[1, 2] as Excel.Range).Value2 = level;
                    (row.Cells[1, 3] as Excel.Range).Value2 = sheet ?? "";
                    (row.Cells[1, 4] as Excel.Range).Value2 = address ?? "";
                    (row.Cells[1, 5] as Excel.Range).Value2 = msg ?? "";
#pragma warning restore CS8602 // null 가능 참조에 대한 역참조입니다.

                }
                catch { /* logging failure is non-fatal */ }
                finally
                {
                    XqlCommon.ReleaseCom(row); XqlCommon.ReleaseCom(cell); XqlCommon.ReleaseCom(ws); XqlCommon.ReleaseCom(wb);
                }
            });
        }

        private static Excel.Worksheet FindOrCreateLogSheet(Excel.Workbook wb, string name)
        {
            Excel.Worksheet? match = null;
            foreach (Excel.Worksheet s in wb.Worksheets)
            {
                bool keep = false;
                try
                {
                    if (string.Equals(s.Name, name, StringComparison.Ordinal))
                    {
                        match = s; keep = true; break;
                    }
                }
                finally
                {
                    if (!keep) XqlCommon.ReleaseCom(s);
                }
            }
            if (match != null) return match;

            var created = (Excel.Worksheet)wb.Worksheets.Add();
            created.Name = name;
            created.Move(After: wb.Worksheets[wb.Worksheets.Count]); // 맨 뒤
            return created;
        }
    }
}
