using ExcelDna.Integration.CustomUI;

namespace XQLite.AddIn
{
    public sealed class XqlRibbon : ExcelRibbon
    {
        public override string GetCustomUI(string ribbonId) => @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='tabXQL' label='XQLite'>
        <group id='grpGeneral' label='General'>
          <button id='btnConfig'   label='Config'    size='large' onAction='OnConfig'    imageMso='HappyFace'/>
          <button id='btnCommit'   label='Commit'               onAction='OnCommit'    imageMso='HappyFace'/>
          <button id='btnRecover'  label='Recover'              onAction='OnRecover'   imageMso='HappyFace'/>
          <button id='btnInspector'label='Inspector'            onAction='OnInspector' imageMso='HappyFace'/>
          <button id='btnExport'   label='Export'               onAction='OnExport'    imageMso='HappyFace'/>
          <button id='btnPresence' label='Presence'             onAction='OnPresence'  imageMso='HappyFace'/>
          <button id='btnSchema'   label='Schema'               onAction='OnSchema'    imageMso='HappyFace'/>
          <button id='btnLockMgr'  label='Locks'                onAction='OnLockMgr'   imageMso='HappyFace'/>
          <button id='btnDiag'     label='Export Diag'          onAction='OnDiag'      imageMso='HappyFace'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        public void OnConfig(IRibbonControl _) => XqlCommands.ConfigCommand();
        public void OnCommit(IRibbonControl _) => XqlCommands.CommitCommand();
        public void OnRecover(IRibbonControl _) => XqlCommands.RecoverCommand();
        public void OnInspector(IRibbonControl _) => XqlCommands.InspectorCommand();
        public void OnExport(IRibbonControl _) => XqlCommands.ExportSnapshotCommand();
        public void OnPresence(IRibbonControl _) => XqlCommands.PresenceCommand();
        public void OnSchema(IRibbonControl _) => XqlCommands.SchemaCommand();
        public void OnLockMgr(IRibbonControl _) => XqlCommands.LockCommand();
    }
}