using ExcelDna.Integration.CustomUI;

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

        // General
        public void OnConfig(IRibbonControl _) => XqlCommands.ConfigCommand();
        public void OnCommit(IRibbonControl _) => XqlCommands.CommitCommand();
        public void OnRecover(IRibbonControl _) => XqlCommands.RecoverCommand();
        public void OnInspector(IRibbonControl _) => XqlCommands.InspectorCommand();
        public void OnExport(IRibbonControl _) => XqlCommands.ExportSnapshotCommand();
        public void OnPresence(IRibbonControl _) => XqlCommands.PresenceCommand();
        public void OnSchema(IRibbonControl _) => XqlCommands.SchemaCommand();
        public void OnLockMgr(IRibbonControl _) => XqlCommands.LockCommand();
        public void OnDiag(IRibbonControl _) => XqlCommands.ExportDiagnosticsCommand();

        // Meta
        public void OnInsertMeta(IRibbonControl _) => XqlCommands.InsertMetaHeaderFromSelection();
        public void OnMetaInfo(IRibbonControl _) => XqlCommands.ShowMetaHeaderInfo();
        public void OnMetaRemove(IRibbonControl _) => XqlCommands.RemoveMetaHeader();
        public void OnRefreshMeta(IRibbonControl _) => XqlCommands.RefreshMetaHeader();

        // 드롭다운 항목 공통 핸들러
        public void OnSetType(IRibbonControl c)
        {
            var type = (c.Tag ?? "").Trim().ToUpperInvariant();
            XqlCommands.SetType(type);
        }
    }
}
