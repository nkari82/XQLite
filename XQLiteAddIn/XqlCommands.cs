using ExcelDna.Integration;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

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
    }
}