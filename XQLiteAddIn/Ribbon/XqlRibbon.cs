using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Concurrent;
using System.Linq;
using System.Reflection;
using System.Threading; // Interlocked / Volatile
using System.Threading.Tasks;
using System.Windows.Forms;
using static XQLite.AddIn.IXqlBackend;
using Action = System.Action;
using Excel = Microsoft.Office.Interop.Excel;

namespace XQLite.AddIn
{
    public sealed class XqlRibbon : ExcelRibbon
    {
        private IRibbonUI? _ribbon;
        private Excel.Application? _app; // Excel 이벤트 구독용

        // Commit 가드 상태 (서버에 동일 테이블 존재 && 이 워크북이 '처음'이면 Commit 잠금)
        private volatile bool _blockCommit;     // true면 Commit 비활성화
        private long _blockCheckedMs;           // 마지막 점검 시각(ms) - Interlocked로 접근
        private readonly object _blockSync = new();

        // ───────────────────────── Ribbon XML ─────────────────────────
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
          <button id='btnPull' label='풀'
                  onAction='OnPull' imageMso='Refresh'
                  getEnabled='Pull_GetEnabled'
                  screentip='풀'
                  supertip='서버의 최신 증분만 내려받아 반영합니다.'/>
          <button id='btnCommit' label='커밋'
                  onAction='OnCommit' imageMso='SaveAll'
                  getEnabled='Commit_GetEnabled'
                  screentip='커밋'
                  supertip='필요하면 테이블/컬럼을 생성하고, 변경된 셀만 효율적으로 업서트합니다.'/>
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

        // ───────────────────── Ribbon lifecycle / dynamic enable ─────────────────────
        public void OnRibbonLoad(IRibbonUI ribbon)
        {
            _ribbon = ribbon;

            // Excel Application 이벤트 구독: 새 시트 생성 / 시트 전환 시 즉시 재평가
            _app = ExcelDnaUtil.Application as Excel.Application;
            try
            {
                if (_app != null)
                {
                    _app.SheetActivate += App_SheetActivate;
                    _app.WorkbookNewSheet += App_WorkbookNewSheet; // (Excel.Workbook, object)
                }
            }
            catch { /* ignore */ }

            // ⬇️ Pull 진행 상태 변경 시 버튼 갱신
            try
            {
                XqlAddIn.Sync?.PullStateChanged += pulling =>
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    try { _ribbon?.InvalidateControl("btnPull"); } catch { }
                });
            }
            catch { /* ignore */ }

            // 백엔드 상태 변화 시 버튼 재평가
            if (XqlAddIn.Backend is XqlGqlBackend be)
            {
                be.StateChanged += (_, __) =>
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        try { _ribbon?.InvalidateControl("btnStatus"); } catch { }
                        try { _ribbon?.InvalidateControl("btnCommit"); } catch { }
#pragma warning disable CS4014
                        RefreshCommitEnabled(); // 반환값 사용 안 함
#pragma warning restore CS4014
                        try
                        {
                            var app = (Excel.Application)ExcelDnaUtil.Application;
                            app.StatusBar = XqlStatus_GetLabel(null!);
                        }
                        catch { }
                    });
            }

            // ✅ 이벤트 허브 구독: 스키마/커밋 재평가 신호
            XqlEvents.SchemaChanged += () =>
            {
                SetBlock(false);
                try { _ribbon?.InvalidateControl("btnCommit"); } catch { }
            };

            XqlEvents.RequestReevalCommit += () =>
            {
                try { _ribbon?.InvalidateControl("btnCommit"); } catch { }
                _ = RefreshCommitEnabled();
            };
        }

        private void App_WorkbookNewSheet(Excel.Workbook wb, object sh)
        {
            ExcelAsyncUtil.QueueAsMacro(SafeReevalCommit);
        }
        private void App_SheetActivate(object Sh)
        {
            ExcelAsyncUtil.QueueAsMacro(SafeReevalCommit);
        }
        private void SafeReevalCommit()
        {
            try { _ribbon?.InvalidateControl("btnCommit"); } catch { }
#pragma warning disable CS4014
            RefreshCommitEnabled(); // 반환값 무시
#pragma warning restore CS4014
        }

        // 1초에 한 번만 재평가(스레드 안전)
        private bool ShouldRecheck()
        {
            var now = XqlCommon.NowMs();
            var last = Interlocked.Read(ref _blockCheckedMs);
            if (now - last <= 1000) return false;
            Interlocked.Exchange(ref _blockCheckedMs, now);
            return true;
        }

        // Commit 활성/비활성 동적 제어
        public bool Commit_GetEnabled(IRibbonControl _)
        {
            if (ShouldRecheck())
#pragma warning disable CS4014
                RefreshCommitEnabled(); // 비동기 재평가
#pragma warning restore CS4014
            return !_blockCommit;
        }

        // Pull 버튼 활성/비활성
        public bool Pull_GetEnabled(IRibbonControl _)
        {
            var pulling = XqlAddIn.Sync?.IsPulling ?? false;
            return !pulling;
        }

        // ───────────────────── Excel UI 스레드에서 동기 실행 헬퍼 ─────────────────────
        private static Task<T> RunOnExcel<T>(Func<T> fn)
        {
            var tcs = new TaskCompletionSource<T>();
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try { tcs.SetResult(fn()); }
                catch (Exception ex) { tcs.SetException(ex); }
            });
            return tcs.Task;
        }
        private static Task RunOnExcel(Action fn)
        {
            var tcs = new TaskCompletionSource<object?>();
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try { fn(); tcs.SetResult(null); }
                catch (Exception ex) { tcs.SetException(ex); }
            });
            return tcs.Task;
        }

        private sealed class ExcelSideInfo
        {
            public string Table = "";
            public bool HasLocal;
            public bool HasUsableHeader;
        }

        private async Task RefreshCommitEnabled()
        {
            try
            {
                var be = XqlAddIn.Backend as XqlGqlBackend;
                var sheet = XqlAddIn.Sheet;
                if (be == null || sheet == null) { SetBlock(false); return; }

                // ① Excel UI 스레드에서 COM 안전하게 정보 수집
                ExcelSideInfo info = await RunOnExcel(() =>
                {
                    var app = ExcelDnaUtil.Application as Excel.Application;
                    if (app == null) return new ExcelSideInfo();
                    if (app.ActiveSheet is not Excel.Worksheet ws) return new ExcelSideInfo();

                    var wb = ws.Parent as Excel.Workbook;
                    if (wb == null) return new ExcelSideInfo();

                    bool hasLocal = false;
                    try
                    {
                        var st = XqlSheet.EnsureStateSheet(wb);
                        var ur = st.UsedRange;
                        int rows = ur.Row + ur.Rows.Count - 1;
                        hasLocal = rows > 1;
                        XqlCommon.ReleaseCom(ur, st);
                    }
                    catch { hasLocal = false; }

                    // 테이블명
                    var sm = sheet.GetOrCreateSheet(ws.Name);
                    var table = string.IsNullOrWhiteSpace(sm.TableName) ? ws.Name : sm.TableName!;

                    // 헤더 준비도
                    bool hasUsableHeader = false;
                    Excel.Range? header = null;
                    try
                    {
                        if (!XqlSheet.TryGetHeaderMarker(ws, out header))
                            header = XqlSheet.GetHeaderRange(ws);

                        if (header != null)
                        {
                            var names = XqlSheet.ComputeHeaderNames(header);
                            bool allFallback = XqlSheet.IsFallbackLetterHeader(header);
                            hasUsableHeader = !allFallback && names.Any(n => !string.IsNullOrWhiteSpace(n));
                        }
                    }
                    catch { hasUsableHeader = false; }
                    finally { XqlCommon.ReleaseCom(header); }

                    return new ExcelSideInfo { Table = table, HasLocal = hasLocal, HasUsableHeader = hasUsableHeader };
                }).ConfigureAwait(false);

                // ② 서버 측 확인 (백그라운드 OK)
                bool hasServer = false;
                try
                {
                    var cols = await be.GetTableColumns(info.Table).ConfigureAwait(false);
                    hasServer = cols != null && cols.Count > 0;
                }
                catch { hasServer = false; }

                // ③ 최종 판정 반영은 매크로 큐에서
                bool block = hasServer && !info.HasLocal && !info.HasUsableHeader;
                await RunOnExcel(() => SetBlock(block)).ConfigureAwait(false);
            }
            catch
            {
                SetBlock(false); // 실패/미확정이면 활성
            }
        }

        private void SetBlock(bool block)
        {
            bool changed;
            lock (_blockSync)
            {
                changed = (_blockCommit != block);
                _blockCommit = block;
            }
            if (changed)
            {
                // 리본 무효화는 항상 Excel UI 큐에서
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    try { _ribbon?.InvalidateControl("btnCommit"); } catch { }
                });
            }
        }

        // ───────────────────────── 동작(Commands) ─────────────────────────
        public void OnConfig(IRibbonControl _)
        {
            try { XqlConfigForm.ShowSingleton(); }
            catch (Exception ex) { MessageBox.Show(ex.Message, "XQLite", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        public void OnPull(IRibbonControl _)
        {
            try { XqlAddIn.ExcelInterop?.Cmd_PullOnly(); }
            catch (Exception ex) { MessageBox.Show("Pull 실패: " + ex.Message, "XQLite"); }
            finally
            {
                SafeReevalCommit();
                try { _ribbon?.InvalidateControl("btnPull"); } catch { }
            }
        }

        // 커밋은 “스키마 보강 → 업서트 → 필요 시 Pull”
        public void OnCommit(IRibbonControl _)
        {
            try { XqlAddIn.ExcelInterop?.Cmd_CommitSync(); }
            catch (Exception ex) { MessageBox.Show("Commit 실패: " + ex.Message, "XQLite"); }
            finally { SafeReevalCommit(); }
        }

        // ───────────────────────── 상태 표시(라벨/아이콘/툴팁) ─────────────────────────
        public string XqlStatus_GetLabel(IRibbonControl _)
        {
            var be = XqlAddIn.Backend as XqlGqlBackend;
            if (be == null) return "연결: 알 수 없음";
            return be.State switch
            {
                ConnState.Online => "연결: 정상",
                ConnState.Connecting => "연결 중…",
                ConnState.Degraded => "연결: 지연",
                _ => "연결: 끊김"
            };
        }

        public stdole.IPictureDisp? XqlStatus_GetImage(IRibbonControl _)
        {
            try
            {
                string idMso = (XqlAddIn.Backend as XqlGqlBackend)?.State switch
                {
                    ConnState.Online => "PersonaStatusOnline",
                    ConnState.Connecting => "PersonaStatusAway",
                    ConnState.Degraded => "PersonaStatusBusy",
                    _ => "PersonaStatusOffline"
                } ?? "PersonaStatusOffline";
                return MsoImageHelper.Get(idMso, 32);
            }
            catch { return null; }
        }

        public string XqlStatus_GetSupertip(IRibbonControl _)
        {
            var be = XqlAddIn.Backend as XqlGqlBackend;
            if (be == null) return "백엔드가 초기화되지 않았습니다.";
            return be.State switch
            {
                ConnState.Online => "서버 상태가 정상입니다.",
                ConnState.Connecting => "서버에 연결 중입니다.",
                ConnState.Degraded => "지연이 감지되었습니다.",
                _ => "서버 연결이 끊겼습니다."
            };
        }

        public async void XqlStatus_OnClick(IRibbonControl _)
        {
            var be = XqlAddIn.Backend as XqlGqlBackend;
            if (be == null) return;
            try { await be.Ping(); }
            catch { /* 실패해도 상태 이벤트로 갱신됨 */ }
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

                XqlSheet.ColumnKind kind = tag switch
                {
                    "INT" or "INTEGER" => XqlSheet.ColumnKind.Int,
                    "REAL" => XqlSheet.ColumnKind.Real,
                    "TEXT" => XqlSheet.ColumnKind.Text,
                    "BOOL" or "BOOLEAN" => XqlSheet.ColumnKind.Bool,
                    "DATE" => XqlSheet.ColumnKind.Date,
                    _ => XqlSheet.ColumnKind.Text,
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
                    var ct = sm.Columns.TryGetValue(colName!, out var cur) ? cur : new XqlSheet.ColumnType();
                    ct.Kind = kind;
                    sm.SetColumn(colName!, ct);

                    // 툴팁/검증 갱신
                    var tips = XqlSheetView.BuildHeaderTooltips(sm, header);
                    XqlSheetView.SetHeaderTooltips(header, tips);
                    XqlSheetView.ApplyDataValidationForHeader(ws, header, sm);
                }
                finally
                {
                    XqlCommon.ReleaseCom(cell, hit);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Set Type failed: " + ex.Message, "XQLite");
            }
        }

        // ───────────────────────── Mso Image Helper ─────────────────────────
        internal static class MsoImageHelper
        {
            private static readonly ConcurrentDictionary<string, stdole.IPictureDisp> _cache =
                new(StringComparer.OrdinalIgnoreCase);

            /// <summary>Office 버전/PIA 차이를 피하기 위해 리플렉션으로 GetImageMso 호출.</summary>
            public static stdole.IPictureDisp? Get(string idMso, int size = 32)
            {
                if (string.IsNullOrWhiteSpace(idMso)) return null;

                if (_cache.TryGetValue(idMso, out var pic))
                    return pic;

                try
                {
                    object app = ExcelDnaUtil.Application;                // Excel.Application (COM)
                    var appType = app.GetType();

                    // Application.CommandBars
                    object commandBars = appType.InvokeMember(
                        "CommandBars",
                        BindingFlags.GetProperty,
                        binder: null, target: app, args: null);

                    // CommandBars.GetImageMso(string idMso, int width, int height) : IPictureDisp
                    object? ret = commandBars.GetType().InvokeMember(
                        "GetImageMso",
                        BindingFlags.InvokeMethod,
                        binder: null, target: commandBars,
                        args: new object[] { idMso, size, size });

                    pic = ret as stdole.IPictureDisp;
                    if (pic != null)
                        _cache[idMso] = pic;

                    return pic;
                }
                catch
                {
                    // 실패 시 null 반환 → UI는 기본 아이콘/텍스트로 진행
                    return null;
                }
            }
        }
    }
}
