// XqlLog.cs
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Globalization;
using XQLite.AddIn;
using Excel = Microsoft.Office.Interop.Excel;

namespace XQLite.AddIn
{
    internal static class XqlLog
    {
        public static void Info(string msg, string? sheet = "*", string? address = null)
        {
            Log("INFO", msg, sheet, address);
            try { XqlFileLogger.Write("INFO", sheet ?? "*", msg ?? ""); } catch { }
        }

        public static void Warn(string msg, string? sheet = "*", string? address = null)
        {
            Log("WARN", msg, sheet, address);
            try { XqlFileLogger.Write("WARN", sheet ?? "*", msg ?? ""); } catch { }
        }

        public static void Error(string msg, string? sheet = "*", string? address = null)
        {
            Log("ERROR", msg, sheet, address);
            try { XqlFileLogger.Write("ERR", sheet ?? "*", msg ?? ""); } catch { }
        }

        // ─────────────────────────────────────────────────────────────────
        //  Log: 워크북에 "_XQL_Log" 시트를 만들어 기록 (UI 스레드에서만 작업)
        //  컬럼: Timestamp | Level | Sheet | Address | Message
        // ─────────────────────────────────────────────────────────────────
        public static void Log(string level, string msg, string? sheet, string? address)
        {
            // UI 스레드에서만 Excel COM 호출
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
                Excel.Workbook? wb = null;
                Excel.Worksheet? ws = null;
                Excel.Range? ur = null;
                Excel.Range? rowRange = null;

                try
                {
                    wb = app.ActiveWorkbook;
                    if (wb == null) return;

                    // 시트 찾기/생성
                    ws = FindOrCreateLogSheet(wb, "_XQL_Log");
                    if (ws == null) return;

                    // 1) 헤더 유무 확인(빠르고 안전)
                    bool needHeader = false;
                    try
                    {
                        var h11 = (Excel.Range)ws.Cells[1, 1];
                        var v = h11.Value2;
                        needHeader = v == null || string.IsNullOrWhiteSpace(Convert.ToString(v, CultureInfo.InvariantCulture));
                        XqlCommon.ReleaseCom(h11);
                    }
                    catch { /* fallback 아래에서 처리 */ }

                    // 2) 헤더가 없으면 한 번에 채우기(배열로)
                    if (needHeader)
                    {
                        var hdr = new object[1, 5]
                        {
                            {
                                "Timestamp", "Level", "Sheet", "Address", "Message"
                            }
                        };
                        var hdrRange = ws.Range[ws.Cells[1, 1], ws.Cells[1, 5]];
                        try { hdrRange.Value2 = hdr; }
                        catch { /* ignore */ }
                        finally { XqlCommon.ReleaseCom(hdrRange); }
                    }

                    // 3) 마지막 행 계산(UsedRange 재조회)
                    int lastRow = 1;
                    try
                    {
                        ur = ws.UsedRange as Excel.Range;
                        if (ur != null)
                            lastRow = ur.Row + ur.Rows.Count - 1;
                    }
                    catch { /* ignore */ }
                    finally { XqlCommon.ReleaseCom(ur); ur = null; }

                    int next = Math.Max(2, lastRow + 1);

                    // 4) 로그 한 줄을 배열로 한번에 기록
                    var data = new object[1, 5]
                    {
                        {
                            DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture),
                            level ?? "",
                            sheet ?? "",
                            address ?? "",
                            msg ?? ""
                        }
                    };

                    rowRange = ws.Range[ws.Cells[next, 1], ws.Cells[next, 5]];
                    rowRange.Value2 = data;
                }
                catch
                {
                    // 로깅 실패는 치명적이지 않음
                }
                finally
                {
                    XqlCommon.ReleaseCom(rowRange);
                    XqlCommon.ReleaseCom(ur);
                    XqlCommon.ReleaseCom(ws);
                    XqlCommon.ReleaseCom(wb);
                }
            });
        }

        // "_XQL_Log" 시트 검색/생성 (이름으로 매칭)
        private static Excel.Worksheet? FindOrCreateLogSheet(Excel.Workbook wb, string name)
        {
            Excel.Worksheet? match = null;
            try
            {
                foreach (Excel.Worksheet s in wb.Worksheets)
                {
                    bool keep = false;
                    try
                    {
                        if (string.Equals(s.Name, name, StringComparison.Ordinal))
                        {
                            match = s;
                            keep = true;
                            break;
                        }
                    }
                    finally
                    {
                        if (!keep) XqlCommon.ReleaseCom(s);
                    }
                }

                if (match != null) return match;

                match = (Excel.Worksheet)wb.Worksheets.Add();
                try { match.Name = name; } catch { /* Excel이 고유 이름으로 바꿀 수 있음 */ }
                try { match.Move(After: wb.Worksheets[wb.Worksheets.Count]); } catch { /* ignore */ }
                return match;
            }
            catch
            {
                XqlCommon.ReleaseCom(match);
                return null;
            }
        }
    }
}
