using ExcelDna.Integration; // ExcelDnaUtil.Application
using ExcelDna.Integration.CustomUI;
using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

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
                  screentip='커밋'
                  supertip='필요하면 테이블/컬럼을 생성하고, 변경된 셀만 효율적으로 업서트합니다.'/>
          <button id='btnPull' label='풀'
                  onAction='OnPull' imageMso='Refresh'
                  screentip='풀'
                  supertip='서버의 최신 증분만 내려받아 반영합니다.'/>
          <toggleButton id='tglDropCols' label='헤더 외 컬럼 DROP'
                        getPressed='DropCols_GetPressed'
                        onAction='DropCols_OnToggle'
                        imageMso='DeleteColumns'
                        screentip='헤더에 없는 컬럼 삭제'
                        supertip='켜면 커밋 시 헤더에 없는 서버 컬럼을 DROP합니다(주의: 되돌릴 수 없습니다). 끄면 컬럼은 추가만 합니다.'/>
        </group>

        <!-- 연결 상태 -->
        <group id='grpConn' label='연결 상태'>
          <button id='btnStatus' size='large'
                  getLabel='XqlStatus_GetLabel'
                  getImage='XqlStatus_GetImage'
                  getSupertip='XqlStatus_GetSupertip'
                  onAction='XqlStatus_OnClick' />
        </group>

        <!-- 헤더 -->
        <group id='grpMeta' label='헤더'>
          <button id='btnInsertHeader' label='헤더 삽입' size='large'
                  onAction='OnInsertHeader' imageMso='TableInsertDialog'
                  screentip='메타 헤더 삽입'
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
                  screentip='복구'
                  supertip='엑셀 파일을 DB 원본으로 간주하여 서버를 재구성합니다.'/>
          <button id='btnExport'    label='내보내기'
                  onAction='OnExport' imageMso='ExportTextFile'
                  screentip='내보내기'
                  supertip='DB와 메타/로그를 zip으로 내보냅니다.'/>
          <button id='btnDiag'      label='진단 내보내기'
                  onAction='OnDiag' imageMso='FileSaveAs'
                  screentip='진단 내보내기'
                  supertip='문제 분석을 위한 진단 패키지를 zip으로 내보냅니다.'/>
        </group>

      </tab>
    </tabs>
  </ribbon>
</customUI>";

        // ───────────────────────── Ribbon lifecycle / dynamic enable ─────────────────────────
        public void OnRibbonLoad(IRibbonUI ribbon)
        {
            _ribbon = ribbon;
            try
            {
                if (XqlAddIn.Backend is XqlGqlBackend be)
                {
                    be.StateChanged += (_, __) =>
                        ExcelAsyncUtil.QueueAsMacro(() =>
                        {
                            try { _ribbon?.InvalidateControl("btnStatus"); } catch { }
                            try
                            {
                                var app = ExcelDnaUtil.Application as Excel.Application;
                                if (app != null) app.StatusBar = XqlStatus_GetLabel(null!);
                            }
                            catch { }
                        });
                }
            }
            catch { }
        }

        // ───────────────────────── General ─────────────────────────
        public void OnConfig(IRibbonControl _)
        {
            try { XqlConfigForm.ShowSingleton(); }
            catch (Exception ex) { MessageBox.Show(ex.Message, "XQLite", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        public void OnPull(IRibbonControl _)
        {
            try { XqlAddIn.ExcelInterop?.Cmd_PullOnly(); }
            catch (Exception ex) { MessageBox.Show("Pull 실패: " + ex.Message, "XQLite"); }
        }

        // 기존 커밋 핸들러는 “스키마 보강 + 푸시 + 필요시 풀”
        public void OnCommit(IRibbonControl _)
        {
            try { XqlAddIn.ExcelInterop?.Cmd_CommitSmart(); }
            catch (Exception ex) { MessageBox.Show("Commit 실패: " + ex.Message, "XQLite"); }
        }

        // 리본 토글 상태 반환
        public bool DropCols_GetPressed(IRibbonControl _)
        {
            try { return XqlConfig.DropColumnsOnCommit; }  // 설정 값 그대로 반영
            catch { return false; }
        }

        // 토글 변경 시 저장
        public void DropCols_OnToggle(IRibbonControl _, bool pressed)
        {
            try
            {
                XqlConfig.DropColumnsOnCommit = pressed; // 설정 반영
                XqlConfig.Save();                        // 영구 저장
                                                         // 즉시 UI 반영(선택): 상태바에 잠깐 안내
                var app = ExcelDnaUtil.Application as Excel.Application;
                if (app != null) app.StatusBar = pressed ? "DROP Columns on Commit: ON" : "DROP Columns on Commit: OFF";
            }
            catch (Exception ex)
            {
                MessageBox.Show("설정 저장 실패: " + ex.Message, "XQLite");
            }
        }

        // ───────────────────────── Connection Status (Label / Icon / Tip / Click) ─────────────────────────
        public string XqlStatus_GetLabel(IRibbonControl _)
        {
            var be = XqlAddIn.Backend as XqlGqlBackend;
            if (be == null) return "오프라인";

            string since = "";
            if (be.LastOkUtc != DateTime.MinValue)
            {
                var diff = DateTime.UtcNow - be.LastOkUtc;
                since = diff.TotalSeconds < 60 ? $"{(int)diff.TotalSeconds}초" : $"{(int)(diff.TotalSeconds / 60)}분";
                since = " · " + since;
            }

            return be.State switch
            {
                IXqlBackend.ConnState.Online => "온라인" + since,
                IXqlBackend.ConnState.Connecting => "연결 중…",
                IXqlBackend.ConnState.Degraded => "품질 저하" + since,
                _ => "오프라인"
            };
        }

        public stdole.IPictureDisp? XqlStatus_GetImage(IRibbonControl _)
        {
            try
            {
                var app = ExcelDnaUtil.Application as Excel.Application;
                if (app == null) return null;
                var msoId = (XqlAddIn.Backend as XqlGqlBackend)?.State switch
                {
                    IXqlBackend.ConnState.Online => "PersonaStatusOnline",
                    IXqlBackend.ConnState.Connecting => "PersonaStatusAway",
                    IXqlBackend.ConnState.Degraded => "PersonaStatusBusy",
                    _ => "PersonaStatusOffline"
                };
                return (stdole.IPictureDisp)app.CommandBars.GetImageMso(msoId, 32, 32);
            }
            catch { return null; }
        }

        public string XqlStatus_GetSupertip(IRibbonControl _)
        {
            var be = XqlAddIn.Backend as XqlGqlBackend;
            if (be == null) return "백엔드가 초기화되지 않았습니다.";
            var detail = string.IsNullOrWhiteSpace(be.StateDetail) ? "" : $"\r\n{be.StateDetail}";
            return $"서버 상태: {be.State}{detail}";
        }

        public async void XqlStatus_OnClick(IRibbonControl _)
        {
            var be = XqlAddIn.Backend as XqlGqlBackend;
            if (be == null) return;
            try { await be.Ping(); }
            catch { /* 실패해도 상태는 이벤트로 갱신됨 */ }
        }

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
