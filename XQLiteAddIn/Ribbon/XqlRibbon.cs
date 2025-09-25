using ExcelDna.Integration.CustomUI;
using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration; // ExcelDnaUtil.Application 사용


namespace XQLite.AddIn
{
    // https://bert-toolkit.com/imagemso-list.html
    public sealed class XqlRibbon : ExcelRibbon
    {
        public override string GetCustomUI(string ribbonId) => @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='tabXQL' label='XQLite'>
        <group id='grpGeneral' label='General'>
          <button id='btnConfig'    label='Config'       size='large' onAction='OnConfig'    imageMso='PropertySheet'/>
          <button id='btnCommit'    label='Commit'                     onAction='OnCommit'    imageMso='SaveAll'/>
          <button id='btnRecover'   label='Recover'                    onAction='OnRecover'   imageMso='FileCompactAndRepairDatabase'/>
          <button id='btnInspector' label='Inspector'                  onAction='OnInspector' imageMso='Search'/>
          <button id='btnExport'    label='Export'                     onAction='OnExport'    imageMso='ExportTextFile'/>
          <button id='btnPresence'  label='Presence'                   onAction='OnPresence'  imageMso='Piggy'/>
          <button id='btnSchema'    label='Schema'                     onAction='OnSchema'    imageMso='TableDesign'/>
          <button id='btnLockMgr'   label='Locks'                      onAction='OnLockMgr'   imageMso='Lock'/>
          <button id='btnDiag'      label='Export Diag'                onAction='OnDiag'      imageMso='FileSaveAs'/>
        </group>
        <group id='grpMeta' label='메타/테이블'>
          <button id='btnInsertMeta' label='메타 헤더 삽입'
                  size='large' onAction='OnInsertMeta' imageMso='TableInsertDialog'/>
          <button id='btnMetaInfo' label='메타 정보'
                  onAction='OnMetaInfo' imageMso='ZoomPrintPreviewExcel'/>
          <button id='btnMetaRemove' label='메타 제거'
                  onAction='OnMetaRemove' imageMso='TableDelete'/>
          <button id='btnRefreshMeta' label='새로고침'
                  onAction='OnRefreshMeta' imageMso='Refresh'/>
          <menu id='menuColType' label='Column Type' imageMso='TableInsertDialog'>
            <button id='typeInt'  label='INT (정수)'  onAction='OnSetType' tag='INT' />
            <button id='typeReal' label='REAL (실수)' onAction='OnSetType' tag='REAL' />
            <button id='typeText' label='TEXT (문자)' onAction='OnSetType' tag='TEXT' />
            <button id='typeBool' label='BOOL'       onAction='OnSetType' tag='BOOL' />
            <button id='typeDate' label='DATE (날짜)' onAction='OnSetType' tag='DATE' />
          </menu>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        // ===== General =====
        public void OnConfig(IRibbonControl _)
        {
            try
            {
                XqlConfigForm.ShowSingleton();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "XQLite", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnCommit(IRibbonControl _)
            => XqlAddIn.ExcelInterop?.Cmd_CommitSync();

        public void OnRecover(IRibbonControl _)
        {
            try
            {
                XqlRecoverForm.ShowSingleton();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Recover UI failed: " + ex.Message, "XQLite");
            }
        }

        public void OnInspector(IRibbonControl _)
        {
            try
            {
                XqlInspectorForm.ShowSingleton();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Inspector failed: " + ex.Message, "XQLite");
            }
        }

        public void OnExport(IRibbonControl _)
        {
            if (XqlAddIn.Backup == null)
            {
                MessageBox.Show("Backup 모듈이 초기화되지 않았습니다.", "XQLite");
                return;
            }

            using (var sfd = new SaveFileDialog
            {
                Title = "Export (zip)",
                Filter = "Zip (*.zip)|*.zip",
                FileName = $"xql_export_{DateTime.Now:yyyyMMdd_HHmm}.zip",
                OverwritePrompt = true
            })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    XqlAddIn.Backup?.ExportDb(sfd.FileName);
                }
            }
        }

        public void OnPresence(IRibbonControl _)
        {
            try
            {
                XqlPresenceHudForm.ShowSingleton();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Presence HUD failed: " + ex.Message, "XQLite");
            }
        }

        public void OnSchema(IRibbonControl _)
        {
            try
            {
                XqlSchemaForm.ShowSingleton();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Schema explorer failed: " + ex.Message, "XQLite");
            }
        }

        public void OnLockMgr(IRibbonControl _)
        {
            try
            {
                XqlLockForm.ShowSingleton();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Schema explorer failed: " + ex.Message, "XQLite");
            }
        }

        public void OnDiag(IRibbonControl _)
        {
            if (XqlAddIn.Backup == null)
            {
                MessageBox.Show("Backup 모듈이 초기화되지 않았습니다.", "XQLite");
                return;
            }

            using var sfd = new SaveFileDialog
            {
                Title = "Export Diagnostics (zip)",
                Filter = "Zip (*.zip)|*.zip",
                FileName = $"xql_diag_{DateTime.Now:yyyyMMdd_HHmm}.zip",
                OverwritePrompt = true
            };
            if (sfd.ShowDialog() == DialogResult.OK)
                XqlAddIn.Backup?.ExportDiagnostics(sfd.FileName);
        }

        // ===== Meta =====

        public void OnInsertMeta(IRibbonControl _) => XqlSheetView.InsertMetaHeaderFromSelection();
        public void OnMetaInfo(IRibbonControl _) => XqlSheetView.ShowMetaHeaderInfo();
        public void OnMetaRemove(IRibbonControl _) => XqlSheetView.RemoveMetaHeader();
        public void OnRefreshMeta(IRibbonControl _) => XqlSheetView.RefreshMetaHeader();

        // 드롭다운 항목 공통 핸들러
        public void OnSetType(IRibbonControl c)
        {
            try
            {
                var tag = (c?.Tag as string ?? "").Trim().ToUpperInvariant();
                if (string.IsNullOrEmpty(tag)) return;

                ColumnKind kind = tag switch
                {
                    "INT" or "INTEGER" => ColumnKind.Int,
                    "REAL" => ColumnKind.Real,
                    "TEXT" => ColumnKind.Text,
                    "BOOL" or "BOOLEAN" => ColumnKind.Bool,
                    "DATE" => ColumnKind.Date,
                    _ => ColumnKind.Text,
                };

                var app = (Excel.Application)ExcelDnaUtil.Application;
                if (app.ActiveSheet is not Excel.Worksheet ws) return;
                var sheet = XqlAddIn.Sheet;
                if (sheet == null) { MessageBox.Show("MetaRegistry not ready.", "XQLite"); return; }

                Excel.Range? sel = app.Selection as Excel.Range;
                var header = XqlSheetView.ResolveHeader(ws, sel, sheet);
                if (header == null) { MessageBox.Show("헤더를 선택하거나 표 헤더에서 실행하세요.", "XQLite"); return; }

                var hit = sel != null ? ws.Application.Intersect(header, sel) : null;
                Excel.Range? cell = (Excel.Range)((hit != null && hit.Cells.Count >= 1) ? hit.Cells[1, 1] : header.Cells[1, 1]);

                try
                {
                    var colName = (cell.Value2 as string)?.Trim();
                    if (string.IsNullOrEmpty(colName))
                        colName = XqlCommon.ColumnIndexToLetter(cell.Column); // 폴백

                    if (string.IsNullOrEmpty(colName))
                    {
                        MessageBox.Show("컬럼명을 찾을 수 없습니다.", "XQLite");
                        return;
                    }

                    var sm = sheet.GetOrCreateSheet(ws.Name);
#pragma warning disable CS8604 // 가능한 null 참조 인수입니다.
                    var ct = sm.Columns.TryGetValue(colName, out var cur) ? cur : new ColumnType();
#pragma warning restore CS8604 // 가능한 null 참조 인수입니다.
                    ct.Kind = kind;
                    sm.SetColumn(colName, ct);

                    // 주석/툴팁 갱신
                    var dict = sheet.BuildTooltipsForSheet(ws.Name);
                    XqlSheetView.SetHeaderTooltips(header, dict);
                }
                finally
                {
                    XqlCommon.ReleaseCom(cell);
                    XqlCommon.ReleaseCom(hit);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Set Type failed: " + ex.Message, "XQLite");
            }
        }

    }
}
