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
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        // 콜백
        public void OnConfig(IRibbonControl _) => XqlCommands.ConfigCommand();
        public void OnCommit(IRibbonControl _) => XqlCommands.CommitCommand();
        public void OnRecover(IRibbonControl _) => XqlCommands.RecoverCommand();
        public void OnInspector(IRibbonControl _) => XqlCommands.InspectorCommand();
        public void OnExport(IRibbonControl _) => XqlCommands.ExportSnapshotCommand();
        public void OnPresence(IRibbonControl _) => XqlCommands.PresenceCommand();
        public void OnSchema(IRibbonControl _) => XqlCommands.SchemaCommand();
        public void OnLockMgr(IRibbonControl _) => XqlCommands.LockCommand();
        public void OnDiag(IRibbonControl _) => XqlCommands.ExportDiagCommand();
    }
}
