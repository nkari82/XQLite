using ExcelDna.Integration.CustomUI;
using System;
using System.Windows.Forms;

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
            if (XqlBackup.Instance == null)
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
                    XqlBackup.Instance?.ExportDb(sfd.FileName);
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
                XqlSchemaForm.ShowSingleton();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Schema explorer failed: " + ex.Message, "XQLite");
            }
        }

        public void OnDiag(IRibbonControl _)
        {
            if (XqlBackup.Instance == null)
            {
                MessageBox.Show("Backup 모듈이 초기화되지 않았습니다.", "XQLite");
                return;
            }

            using (var sfd = new SaveFileDialog
            {
                Title = "Export Diagnostics (zip)",
                Filter = "Zip (*.zip)|*.zip",
                FileName = $"xql_diag_{DateTime.Now:yyyyMMdd_HHmm}.zip",
                OverwritePrompt = true
            })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                    XqlBackup.Instance?.ExportDiagnostics(sfd.FileName);
            }
        }

        // ===== Meta =====
#if false
        public void OnInsertMeta(IRibbonControl _) => XqlSheetUtil.InsertMetaHeaderFromSelection();
        public void OnMetaInfo(IRibbonControl _) => XqlSheetUtil.ShowMetaHeaderInfo();
        public void OnMetaRemove(IRibbonControl _) => XqlSheetUtil.RemoveMetaHeader();
        public void OnRefreshMeta(IRibbonControl _) => XqlSheetUtil.RefreshMetaHeader();
#endif
        // 드롭다운 항목 공통 핸들러
        public void OnSetType(IRibbonControl c)
        {
#if false
            var type = (c.Tag ?? "").Trim().ToUpperInvariant();

            var app = (Excel.Application)ExcelDnaUtil.Application;
            if (app.ActiveSheet is not Excel.Worksheet ws || app.Selection is not Excel.Range sel) return;

            // 헤더 한 칸만 기준: 사용자가 범위 선택해도 첫 셀만 취급
            var cell = (Excel.Range)sel.Cells[1, 1];

            // 현재 셀이 메타헤더 “행”에 있는지 간단 검증(선택 사항: 스킵 가능)
            var meta = XqlSheetMetaRegistry.Get(ws);
            if (meta == null || cell.Row != meta.TopRow)
            {
                MessageBox.Show("메타 헤더의 셀을 선택한 후 타입을 지정하세요.");
                return;
            }

            XqlColumnTypeRegistry.SetColumnType(ws, cell, type);
            XqlSheetMetaRegistry.RefreshHeaderBorders(ws);
#endif
        }
    }
}
