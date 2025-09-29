using ExcelDna.Integration.CustomUI;
using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration; // ExcelDnaUtil.Application

namespace XQLite.AddIn
{
    // https://bert-toolkit.com/imagemso-list.html
    public sealed class XqlRibbon : ExcelRibbon
    {
        private IRibbonUI? _ribbon;

        // ───────────────────────── Ribbon XML (그룹화/툴팁/동적활성화) ─────────────────────────
        public override string GetCustomUI(string ribbonId) => @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
  <ribbon>
    <tabs>
      <tab id='tabXQL' label='XQLite' insertAfterMso='TabHome'>
        
        <!-- 설정 -->
        <group id='grpConfig' label='설정'>
          <button id='btnConfig' label='환경설정' size='large'
                  onAction='OnConfig' imageMso='PropertySheet'
                  screentip='환경설정'
                  supertip='서버 주소, 계정, 디바운스/풀 주기 등 XQLite 전역 설정을 관리합니다.'/>
        </group>

        <!-- 동기화 -->
        <group id='grpSync' label='동기화'>
          <button id='btnCommit' label='커밋'
                  onAction='OnCommit' imageMso='SaveAll'
                  screentip='Commit'
                  supertip='편집분(2초 디바운스)을 즉시 서버에 업서트합니다.'/>
        </group>

        <!-- 헤더 -->
        <group id='grpMeta' label='헤더'>
          <button id='btnInsertHeader' label='헤더 삽입' size='large'
                  onAction='OnInsertHeader' imageMso='TableInsertDialog'
                  screentip='메타 헤더 삽입'
                  getEnabled='GetEnabled_InsertHeader'
                  supertip='선택 영역을 표 헤더로 지정하고, 컬럼 메타·검증·툴팁을 설치합니다.'/>
          <button id='btnRefreshHeader' label='새로고침'
                  onAction='OnRefreshHeader' imageMso='Refresh'
                  screentip='메타 새로고침'
                  supertip='현재 헤더 기준으로 툴팁/검증 표시를 다시 적용합니다.'/>
          <button id='btnHeaderInfo' label='헤더 정보'
                  onAction='OnHeaderInfo' imageMso='ZoomPrintPreviewExcel'
                  screentip='헤더 정보'
                  supertip='현재 시트의 컬럼 타입/제약 정보를 요약해서 보여줍니다.'/>
          <button id='btnHeaderRemove' label='헤더 제거'
                  onAction='OnHeaderRemove' imageMso='TableDelete'
                  screentip='헤더 제거'
                  supertip='헤더 표시, 검증, 마커(Name)를 제거합니다.'/>
          <menu id='menuColType' label='컬럼 타입' imageMso='TableProperties'
                screentip='Column Type'
                supertip='선택한 헤더 셀의 컬럼 타입을 지정합니다.'>
            <button id='typeInt'  label='INT (정수)'     onAction='OnSetType' tag='INT'/>
            <button id='typeReal' label='REAL (실수)'    onAction='OnSetType' tag='REAL'/>
            <button id='typeText' label='TEXT (문자)'    onAction='OnSetType' tag='TEXT'/>
            <button id='typeBool' label='BOOL (참/거짓)' onAction='OnSetType' tag='BOOL'/>
            <button id='typeDate' label='DATE (날짜)'    onAction='OnSetType' tag='DATE'/>
          </menu>
        </group>

        <!-- 협업 -->
        <group id='grpCollab' label='협업'>
          <button id='btnPresence'  label='프레즌스 HUD' size='large'
                  onAction='OnPresence' imageMso='PeoplePane'
                  screentip='Presence HUD'
                  supertip='동시작업자 위치/셀 잠금을 확인합니다.'/>
          <button id='btnLockMgr'   label='잠금 관리'
                  onAction='OnLockMgr' imageMso='Lock'
                  screentip='Lock Manager'
                  supertip='현재 시트/셀의 잠금 정보를 점검하고 해제할 수 있습니다.'/>
        </group>

        <!-- DB/스키마 -->
        <group id='grpDb' label='DB/스키마'>
          <button id='btnInspector' label='인스펙터' size='large'
                  onAction='OnInspector' imageMso='Search'
                  screentip='Inspector'
                  supertip='동기화 기록/충돌/버전 정보를 빠르게 확인합니다.'/>
          <button id='btnSchema'    label='스키마'
                  onAction='OnSchema' imageMso='TableDesign'
                  screentip='Schema'
                  supertip='SQLite(서버)의 테이블/인덱스/뷰/트리거를 확인합니다.'/>
        </group>

        <!-- 백업/복구/진단 -->
        <group id='grpBackup' label='백업/복구/진단'>
          <button id='btnRecover'   label='복구' size='large'
                  onAction='OnRecover' imageMso='FileCompactAndRepairDatabase'
                  getEnabled='GetEnabled_Backup'
                  screentip='복구'
                  supertip='엑셀 파일을 DB 원본으로 간주하여 서버를 재구성합니다.'/>
          <button id='btnExport'    label='내보내기'
                  onAction='OnExport' imageMso='ExportTextFile'
                  getEnabled='GetEnabled_Backup'
                  screentip='Export'
                  supertip='DB와 메타/로그를 zip으로 내보냅니다.'/>
          <button id='btnDiag'      label='진단 내보내기'
                  onAction='OnDiag' imageMso='FileSaveAs'
                  getEnabled='GetEnabled_Backup'
                  screentip='Diagnostics Export'
                  supertip='문제 분석을 위한 진단 패키지를 zip으로 내보냅니다.'/>
        </group>

      </tab>
    </tabs>
  </ribbon>
</customUI>";

        // ───────────────────────── Ribbon lifecycle / dynamic enable ─────────────────────────
        public void OnRibbonLoad(IRibbonUI ribbon) => _ribbon = ribbon;

        // Backup 모듈이 준비돼야 Export/Recover/Diag 활성화
        public bool GetEnabled_Backup(IRibbonControl _) => XqlAddIn.Backup != null;

        public bool GetEnabled_InsertHeader(IRibbonControl _)
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                if (app.ActiveSheet is not Excel.Worksheet ws) return false;
                // 헤더가 이미 있으면 비활성화
                return !XqlSheet.TryGetHeaderMarker(ws, out var _);
            }
            catch { return true; } // 오류 시에는 일단 노출
        }

        // ───────────────────────── General ─────────────────────────
        public void OnConfig(IRibbonControl _)
        {
            try { XqlConfigForm.ShowSingleton(); }
            catch (Exception ex) { MessageBox.Show(ex.Message, "XQLite", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        public void OnCommit(IRibbonControl _) => XqlAddIn.ExcelInterop?.Cmd_CommitSync();

        // ───────────────────────── Backup / Recover / Diagnostics ─────────────────────────
        public void OnRecover(IRibbonControl _)
        {
            try { XqlRecoverForm.ShowSingleton(); }
            catch (Exception ex) { MessageBox.Show("Recover UI failed: " + ex.Message, "XQLite"); }
        }

        public void OnExport(IRibbonControl _)
        {
            if (XqlAddIn.Backup == null) { MessageBox.Show("Backup 모듈이 초기화되지 않았습니다.", "XQLite"); return; }
            using var sfd = new SaveFileDialog
            {
                Title = "Export (zip)",
                Filter = "Zip (*.zip)|*.zip",
                FileName = $"xql_export_{DateTime.Now:yyyyMMdd_HHmm}.zip",
                OverwritePrompt = true
            };
            if (sfd.ShowDialog() == DialogResult.OK)
                XqlAddIn.Backup?.ExportDb(sfd.FileName);
        }

        public void OnDiag(IRibbonControl _)
        {
            if (XqlAddIn.Backup == null) { MessageBox.Show("Backup 모듈이 초기화되지 않았습니다.", "XQLite"); return; }
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

        // ───────────────────────── Inspector / Schema / Presence / Locks ─────────────────────────
        public void OnInspector(IRibbonControl _)
        {
            try { XqlInspectorForm.ShowSingleton(); }
            catch (Exception ex) { MessageBox.Show("Inspector failed: " + ex.Message, "XQLite"); }
        }

        public void OnSchema(IRibbonControl _)
        {
            try { XqlSchemaForm.ShowSingleton(); }
            catch (Exception ex) { MessageBox.Show("Schema explorer failed: " + ex.Message, "XQLite"); }
        }

        public void OnPresence(IRibbonControl _)
        {
            try { XqlPresenceHudForm.ShowSingleton(); }
            catch (Exception ex) { MessageBox.Show("Presence HUD failed: " + ex.Message, "XQLite"); }
        }

        public void OnLockMgr(IRibbonControl _)
        {
            try { XqlLockForm.ShowSingleton(); }
            catch (Exception ex) { MessageBox.Show("Lock Manager failed: " + ex.Message, "XQLite"); }
        }

        // ───────────────────────── Meta / Header ─────────────────────────
        public void OnInsertHeader(IRibbonControl _) => XqlSheetView.InstallHeader();
        public void OnHeaderInfo(IRibbonControl _) => XqlSheetView.ShowHeaderInfo();
        public void OnHeaderRemove(IRibbonControl _) => XqlSheetView.RemoveHeader();
        public void OnRefreshHeader(IRibbonControl _) => XqlSheetView.RefreshHeader();

        // 컬럼 타입 변경 (메타 반영 + 툴팁/검증 갱신)
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

                // 선택 헤더 셀 → 컬럼명
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
                    var ct = sm.Columns.TryGetValue(colName!, out var cur) ? cur : new ColumnType();
                    ct.Kind = kind;
                    sm.SetColumn(colName!, ct);

                    // 주석/툴팁 갱신: 이름 우선 + 위치(@1,@2..) 폴백으로 일관
                    var tips = XqlSheetView.BuildHeaderTooltips(sm, header);
                    XqlSheetView.SetHeaderTooltips(header, tips);
                    XqlSheetView.ApplyDataValidationForHeader(ws, header, sm);
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
