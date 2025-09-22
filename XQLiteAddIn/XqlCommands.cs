using System;
using System.IO;
using System.Windows.Forms;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;


namespace XQLite.AddIn
{
    internal static class XqlCommands
    {
        [ExcelCommand(Name = "XQL.CmdPalette", Description = "Open XQLite command palette", ShortCut = "Ctrl-Shift-K")]
        internal static void CmdPalette()
        {
            XqlCommandPaletteForm.ShowSingleton();
        }

        [ExcelCommand(Name = "XQL.Config", Description = "Open XQLite Config")]
        internal static void ConfigCommand()
        {
            try
            {
                XqlConfigForm.ShowSingleton();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "XQLite", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        [ExcelCommand(Name = "XQL.Commit", Description = "Commit pending changes", ShortCut = "Ctrl-Shift-C")]
        internal static void CommitCommand()
        {
            try { _ = XqlUpsert.FlushAsync(); }
            catch (Exception ex) { MessageBox.Show("Commit failed: " + ex.Message, "XQLite"); }
        }

        [ExcelCommand(Name = "XQL.Recover", Description = "Recover (batch upsert from current workbook)", ShortCut = "Ctrl-Shift-R")]
        internal static void RecoverCommand()
        {
            try { XqlRecoverForm.ShowSingleton(); }
            catch (Exception ex) { MessageBox.Show("Recover UI failed: " + ex.Message, "XQLite"); }
        }

        [ExcelCommand(Name = "XQL.Inspector", Description = "Open Inspector", ShortCut = "Ctrl-Shift-I")]
        internal static void InspectorCommand()
        {
            try { XqlInspectorForm.ShowSingleton(); }
            catch (Exception ex) { MessageBox.Show("Inspector failed: " + ex.Message, "XQLite"); }
        }

        [ExcelCommand(Name = "XQL.ExportSnapshot", Description = "Export rows snapshot as JSON/CSV", ShortCut = "Ctrl-Shift-E")]
        internal static async void ExportSnapshotCommand()
        {
            try
            {
                // 1) 폴더 선택
                string? targetDir = null;
                using (var dlg = new FolderBrowserDialog
                {
                    Description = "Select folder to save XQLite snapshot"
                })
                {
                    if (dlg.ShowDialog() != DialogResult.OK)
                        return; // 취소
                    targetDir = dlg.SelectedPath;
                }

                // 2) 포맷 선택 (Yes = CSV, No = JSON)
                var fmt = MessageBox.Show(
                    "Export as CSV?\r\nYes = CSV, No = JSON",
                    "XQLite Export",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question);

                if (fmt == DialogResult.Cancel) return;

                bool csv = (fmt == DialogResult.Yes);

                // 3) since_version 선택 (간단히 전체 스냅샷: 0)
                long since = 0L;
                Cursor.Current = Cursors.WaitCursor;

                // 4) 실행 (동기 래핑)
                await XqlExportService.ExportSnapshotAsync(since, targetDir, csv);

                Cursor.Current = Cursors.Default;
                MessageBox.Show(
                    $"Export completed to:\r\n{targetDir}",
                    "XQLite",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Cursor.Current = Cursors.Default;
                MessageBox.Show("Export failed: " + ex.Message, "XQLite", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        [ExcelCommand(Name = "XQL.PresenceHUD", Description = "Show presence HUD", ShortCut = "Ctrl-Shift-P")]
        internal static void PresenceCommand()
        {
            try
            {
                XqlPresenceHudForm.ShowSingleton();
            }
            catch (Exception ex) { MessageBox.Show("Presence HUD failed: " + ex.Message, "XQLite"); }
        }

        [ExcelCommand(Name = "XQL.Schema", Description = "Open schema explorer", ShortCut = "Ctrl-Shift-S")]
        internal static void SchemaCommand()
        {
            try
            {
                XqlSchemaForm.ShowSingleton();
            }
            catch (Exception ex) { MessageBox.Show("Schema explorer failed: " + ex.Message, "XQLite"); }
        }

        [ExcelCommand(Name = "XQL.ExportDiagnostics", Description = "Export diagnostics zip", ShortCut = "Ctrl-Shift-D")]
        internal static async void ExportDiagnosticsCommand()
        {

            using (var dlg = new SaveFileDialog
            {
                Title = "Save XQLite Diagnostic Bundle",
                Filter = "ZIP Archive (*.zip)|*.zip",
                FileName = Path.GetFileName(XqlDiagExport.DefaultDiagZipPath()),
                AddExtension = true,
                OverwritePrompt = true
            })
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {

                    // UI 스레드 블로킹 없이 동작하도록 간단 래퍼
                    try
                    {
                        await XqlDiagExport.ExportAsync(dlg.FileName);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Diag export failed: " + ex.Message, "XQLite");

                    }
                }
            }
        }

        [ExcelCommand(Name = "XQL.Lock", Description = "Lock")]
        internal static void LockCommand()
        {
            try
            {
                XqlLockForm.ShowSingleton();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lock explorer failed: " + ex.Message, "XQLite");
            }
        }


        /// <summary>현재 워크시트의 Selection을 이용해 메타 헤더를 설치</summary>
        public static void InsertMetaHeaderFromSelection()
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                var ws = (Excel.Worksheet)app.ActiveSheet;
                var sel = (Excel.Range)app.Selection;

                if (sel == null)
                {
                    MessageBox.Show("헤더로 사용할 셀(한 줄)을 선택한 후 다시 실행하세요.", "XQLite");
                    return;
                }

                // 한 줄이 아니어도 동작: 첫 번째 행만 사용
                if (sel.Columns.Count < 1)
                {
                    MessageBox.Show("선택된 열이 없습니다.", "XQLite");
                    return;
                }

                if (SheetMetaRegistry.Exists(ws))
                {
                    MessageBox.Show("이 시트에는 이미 메타 헤더가 있습니다.", "XQLite");
                    return;
                }

                // 생성
                SheetMetaRegistry.CreateFromSelection(ws, sel, false);
                MessageBox.Show("메타 헤더가 설치되었습니다. 헤더 아래 행들이 데이터 행으로 취급됩니다.", "XQLite");
            }
            catch (Exception ex)
            {
                MessageBox.Show("메타 헤더 설치에 실패했습니다:\r\n" + ex.Message, "XQLite");
            }
        }

        /// <summary>현재 시트의 메타 헤더 정보를 확인</summary>
        public static void ShowMetaHeaderInfo()
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            var ws = (Excel.Worksheet)app.ActiveSheet;
            var meta = SheetMetaRegistry.Get(ws);
            if (meta == null)
            {
                MessageBox.Show("이 시트에는 메타 헤더가 없습니다.", "XQLite");
                return;
            }
            MessageBox.Show(
                $"시트: {ws.Name}\nTopRow: {meta.TopRow}\nLeftCol: {meta.LeftCol}\nColCount: {meta.ColCount}",
                "XQLite");
        }

        /// <summary>메타 헤더 제거(보호된 기능, 필요 시만 사용)</summary>
        public static void RemoveMetaHeader()
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            var ws = (Excel.Worksheet)app.ActiveSheet;
            if (!SheetMetaRegistry.Exists(ws))
            {
                MessageBox.Show("이 시트에는 메타 헤더가 없습니다.", "XQLite");
                return;
            }
            if (DialogResult.Yes == MessageBox.Show("메타 헤더를 제거할까요?", "XQLite", MessageBoxButtons.YesNo))
            {
                SheetMetaRegistry.Remove(ws);
                MessageBox.Show("메타 헤더가 제거되었습니다.", "XQLite");
            }
        }

        public static void RefreshMetaHeader()
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application; // 또는 여러분 환경에 맞는 Excel.Application 참조
                var ws = app.ActiveSheet as Excel.Worksheet;
                if (ws == null) return;
                SheetMetaRegistry.RefreshHeaderBorders(ws);
            }
            catch (Exception ex)
            {
                MessageBox.Show("메타 새로고침 중 오류: " + ex.Message);
            }
        }
        }
}