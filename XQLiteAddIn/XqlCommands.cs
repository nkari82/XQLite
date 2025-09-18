using ExcelDna.Integration;
using System;
using System.Windows.Forms;

namespace XQLite.AddIn
{
#if true
    public static class XqlCommands
    {
        [ExcelCommand(Name = "XQL.CmdPalette", Description = "Open XQLite command palette", ShortCut = "Ctrl-Shift-K")]
        public static void CmdPalette()
        {
            XqlCommandPaletteForm.ShowSingleton();
        }

        [ExcelCommand(Name = "XQL.Config", Description = "Open XQLite Config")]
        public static void ConfigCommand()
        {
            // 1
            try 
            { 
                XqlConfigForm.ShowSingleton(); 
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "XQLite", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        [ExcelCommand(Name = "XQL.Commit", Description = "Commit pending changes", ShortCut = "Ctrl-Shift-C")]
        public static void CommitCommand()
        {
            try { _ = XqlUpsert.FlushAsync(); }
            catch (Exception ex) { MessageBox.Show("Commit failed: " + ex.Message, "XQLite"); }
        }

        [ExcelCommand(Name = "XQL.Recover", Description = "Recover (batch upsert from current workbook)", ShortCut = "Ctrl-Shift-R")]
        public static void RecoverCommand()
        {
            // 2
            try { XqlRecoverForm.ShowSingleton(); }
            catch (Exception ex) { MessageBox.Show("Recover UI failed: " + ex.Message, "XQLite"); }
        }

        [ExcelCommand(Name = "XQL.Inspector", Description = "Open Inspector", ShortCut = "Ctrl-Shift-I")]
        public static void InspectorCommand()
        {
            // 3
            try { XqlInspectorForm.ShowSingleton(); }
            catch (Exception ex) { MessageBox.Show("Inspector failed: " + ex.Message, "XQLite"); }
        }


        [ExcelCommand(Name = "XQL.ExportSnapshot", Description = "Export rows snapshot as JSON/CSV", ShortCut = "Ctrl-Shift-E")]
        public static void ExportSnapshotCommand()
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
                XqlExportService.ExportSnapshotAsync(since, targetDir, csv)
                                .GetAwaiter().GetResult();

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
        public static void PresenceCommand()
        {
            // 4
            try 
            { 
                XqlPresenceHudForm.ShowSingleton(); 
            }
            catch (Exception ex) { MessageBox.Show("Presence HUD failed: " + ex.Message, "XQLite"); }
        }

        [ExcelCommand(Name = "XQL.Schema", Description = "Open schema explorer", ShortCut = "Ctrl-Shift-S")]
        public static void SchemaCommand()
        {
            // 5
            try 
            { 
                XqlSchemaForm.ShowSingleton(); 
            }
            catch (Exception ex) { MessageBox.Show("Schema explorer failed: " + ex.Message, "XQLite"); }
        }

        [ExcelCommand(Name = "XQL.ExportDiagnostics", Description = "Export diagnostics zip", ShortCut = "Ctrl-Shift-D")]
        public static void ExportDiagnosticsCommand()
        {
            try
            {
                var path = XqlDiagExport.ExportZip();
                MessageBox.Show("Saved:\r\n" + path, "XQLite", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { MessageBox.Show("Diagnostics export failed: " + ex.Message, "XQLite"); }
        }

        [ExcelCommand(Name = "XQL.Lock", Description = "Lock")]
        public static void LockCommand()
        {
            // 6
            try 
            { 
                XqlLockForm.ShowSingleton(); 
            }
            catch (Exception ex) { MessageBox.Show("Lock explorer failed: " + ex.Message, "XQLite"); }
        }

        internal static void ExportDiagCommand()
        {
            
        }
    }
#endif
}