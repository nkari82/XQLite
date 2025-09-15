'Module XQL_Api
Option Explicit

Public Function XQL_GraphQL(ByVal query As String, ByVal variablesJson As String) As String
    Dim xhr As Object: Set xhr = CreateObject("MSXML2.XMLHTTP")
    Dim url$, body$
    url = XQL_GetApiUrl()
    body = "{""query"":""" & Replace(query, """", "\""") & """,""variables"":" & variablesJson & "}"

    xhr.Open "POST", url, False
    xhr.setRequestHeader "Content-Type", "application/json"
    Dim k$: k = XQL_GetApiKey()
    If Len(k) > 0 Then xhr.setRequestHeader "x-api-key", k
    xhr.setRequestHeader "x-actor", XQL_GetNickname()
    xhr.send body
    XQL_GraphQL = CStr(xhr.responseText)
End Function

' 응답에서 max_row_version만 뽑는 간단 헬퍼
Public Function XQL_ExtractMaxRowVersion(ByVal resp As String) As Long
    Dim p&, q&, s$
    p = InStr(1, resp, "max_row_version"":", vbTextCompare)
    If p = 0 Then Exit Function
    q = InStr(p + 16, resp, AnyOf(",}"))
    s = Mid$(resp, p + 16, q - (p + 16))
    XQL_ExtractMaxRowVersion = CLng(val(s))
End Function

Private Function AnyOf(ByVal chars As String) As String
    AnyOf = ",}"
End Function

' 간단 오류 감지
Public Function XQL_HasErrors(ByVal resp As String) As Boolean
    XQL_HasErrors = (InStr(1, resp, """errors"":", vbTextCompare) > 0)
End Function

'Module XQL_AuditView
Option Explicit

Public Function XQL_GraphQL(ByVal query As String, ByVal variablesJson As String) As String
    Dim xhr As Object: Set xhr = CreateObject("MSXML2.XMLHTTP")
    Dim url$, body$
    url = XQL_GetApiUrl()
    body = "{""query"":""" & Replace(query, """", "\""") & """,""variables"":" & variablesJson & "}"

    xhr.Open "POST", url, False
    xhr.setRequestHeader "Content-Type", "application/json"
    Dim k$: k = XQL_GetApiKey()
    If Len(k) > 0 Then xhr.setRequestHeader "x-api-key", k
    xhr.setRequestHeader "x-actor", XQL_GetNickname()
    xhr.send body
    XQL_GraphQL = CStr(xhr.responseText)
End Function

' 응답에서 max_row_version만 뽑는 간단 헬퍼
Public Function XQL_ExtractMaxRowVersion(ByVal resp As String) As Long
    Dim p&, q&, s$
    p = InStr(1, resp, "max_row_version"":", vbTextCompare)
    If p = 0 Then Exit Function
    q = InStr(p + 16, resp, AnyOf(",}"))
    s = Mid$(resp, p + 16, q - (p + 16))
    XQL_ExtractMaxRowVersion = CLng(val(s))
End Function

Private Function AnyOf(ByVal chars As String) As String
    AnyOf = ",}"
End Function

' 간단 오류 감지
Public Function XQL_HasErrors(ByVal resp As String) As Boolean
    XQL_HasErrors = (InStr(1, resp, """errors"":", vbTextCompare) > 0)
End Function

'Module XQL_AuditView
' === Module: XQL_AuditView ===
Option Explicit

Private Const SHEET_AUD As String = "XQLite_Audit"

Public Sub XQL_Audit_Setup()
    Dim ws As Worksheet
    On Error Resume Next: Set ws = Sheets(SHEET_AUD): On Error GoTo 0
    If ws Is Nothing Then
        Set ws = Sheets.add(After:=Sheets(Sheets.count))
        ws.name = SHEET_AUD
    Else
        ws.Cells.Clear
    End If

    With ws
        .Range("A1:F1").Value = Array("actor", "action", "table", "since(YYYY-MM-DD)", "until(YYYY-MM-DD)", "limit")
        .Range("A2:F2").Value = Array("", "", "", "", "", 100)
        .Range("A4").Value = "▶ Run"
        .Range("B4").Value = "◀ Prev"
        .Range("C4").Value = "Next ▶"
        .Range("D4").Value = "? Reset"
        .rows(1).Font.Bold = True
        .Columns("A:H").AutoFit
    End With

    MsgBox SHEET_AUD & " 시트를 생성했습니다." & vbCrLf & _
           "- A2~F2 필터 입력 후 Run" & vbCrLf & _
           "- Prev/Next로 페이지 이동", vbInformation
End Sub

Public Sub XQL_Audit_Run()
    Dim ws As Worksheet: Set ws = EnsureSheet()
    Dim actor$, action$, table$, since$, until$, limit&, offset&
    actor = CStr(ws.Range("A2").Value2)
    action = CStr(ws.Range("B2").Value2)
    table = CStr(ws.Range("C2").Value2)
    since = CStr(ws.Range("D2").Value2)
    until = CStr(ws.Range("E2").Value2)
    limit = CLng(Nz(ws.Range("F2").Value2, 100))
    Offset = CLng(Nz(ws.Range("G2").Value2, 0)) ' 내부용 오프셋 저장 셀 (숨김 가능)

    Dim q$, vars$, resp$, parsed As Object, arr As Object
    q = "query($a:String,$ac:String,$t:String,$s:String,$u:String,$l:Int,$o:Int){" & _
        "auditLog(actor:$a,action:$ac,table:$t,since:$s,until:$u,limit:$l,offset:$o){id ts actor action table_name detail}}"
    vars = "{""a"":" & JNullOrStr(actor) & ",""ac"":" & JNullOrStr(action) & ",""t"":" & JNullOrStr(table) & _
           ",""s"":" & JNullOrStr(since) & ",""u"":" & JNullOrStr(until) & ",""l"":" & CStr(limit) & ",""o"":" & CStr(offset) & "}"
    resp = XQL_GraphQL(q, vars)
    If XQL_HasErrors(resp) Then
        MsgBox "auditLog 실패: " & resp, vbExclamation: Exit Sub
    End If

    Set parsed = JsonConverter.ParseJson(resp)
    Set arr = parsed("data")("auditLog")

    ClearOut ws
    ws.Range("A6:F6").Value = Array("id", "ts", "actor", "action", "table", "detail")
    ws.rows(6).Font.Bold = True

    If arr Is Nothing Or arr.count = 0 Then
        ws.Range("A8").Value = "(no entries)"
        Exit Sub
    End If

    Dim i&, base&: base = 7
    For i = 1 To arr.count
        ws.Cells(base + i - 1, 1).Value = Nz(arr(i)("id"), "")
        ws.Cells(base + i - 1, 2).Value = Nz(arr(i)("ts"), "")
        ws.Cells(base + i - 1, 3).Value = Nz(arr(i)("actor"), "")
        ws.Cells(base + i - 1, 4).Value = Nz(arr(i)("action"), "")
        ws.Cells(base + i - 1, 5).Value = Nz(arr(i)("table_name"), "")
        ws.Cells(base + i - 1, 6).Value = Nz(arr(i)("detail"), "")
    Next i
    ws.Columns("A:F").AutoFit
End Sub

Public Sub XQL_Audit_Next()
    Dim ws As Worksheet: Set ws = EnsureSheet()
    ws.Range("G2").Value = CLng(Nz(ws.Range("G2").Value2, 0)) + CLng(Nz(ws.Range("F2").Value2, 100))
    XQL_Audit_Run
End Sub

Public Sub XQL_Audit_Prev()
    Dim ws As Worksheet: Set ws = EnsureSheet()
    Dim off&: off = CLng(Nz(ws.Range("G2").Value2, 0)) - CLng(Nz(ws.Range("F2").Value2, 100))
    If off < 0 Then off = 0
    ws.Range("G2").Value = off
    XQL_Audit_Run
End Sub

Public Sub XQL_Audit_Reset()
    Dim ws As Worksheet: Set ws = EnsureSheet()
    ws.Range("A2:F2").ClearContents
    ws.Range("F2").Value = 100
    ws.Range("G2").Value = 0
    ClearOut ws
End Sub

' ── helpers ───────────────────────────────────────────

Private Function EnsureSheet() As Worksheet
    On Error Resume Next: Set EnsureSheet = Sheets(SHEET_AUD): On Error GoTo 0
    If EnsureSheet Is Nothing Then
        XQL_Audit_Setup
        Set EnsureSheet = Sheets(SHEET_AUD)
    End If
End Function

Private Sub ClearOut(ws As Worksheet)
    ws.Range("A6:Z1000000").ClearContents
End Sub

Private Function JNullOrStr(ByVal s As String) As String
    s = Trim$(s)
    If Len(s) = 0 Then JNullOrStr = "null" Else JNullOrStr = """" & Replace(s, """", "\""") & """"
End Function

Private Function Nz(ByVal v As Variant, ByVal fb As Variant) As Variant
    If IsEmpty(v) Or v = "" Or v Is Nothing Then Nz = fb Else Nz = v
End Function

'Module XQL_AutoLock
Option Explicit

' 현재 내가 보유한 락(1개만 관리)
Private gLockedSheet As String
Private gLockedCell As String
Private mRefreshAt As Date

' 선택 변경 시 호출(Workbook 이벤트에서 호출)
Public Sub XQL_AutoLock_OnSelectionChange(ByVal ws As Worksheet, ByVal target As Range)
    On Error GoTo done
    If Not XQL_AutoLockEnabled() Then Exit Sub
    If ws.name = "XQLite" Or ws.name = "XQLite_Conflicts" Or ws.name = "XQLite_Presence" Then Exit Sub
    If target.row < 2 Then Exit Sub ' 헤더/타입 행 제외

    ' 타인 락이면 선택을 차단/이동
    If IsLockedByOther(ws.name, target) Then
        MsgBox "이 영역은 다른 사용자가 잠금 중입니다." & vbCrLf & "(" & ws.name & "!" & target.Address(False, False) & ")", vbExclamation
        SafeMoveSelection ws
        Exit Sub
    End If

    ' 이전 락을 해제(설정에 따라)
    If XQL_AutoReleaseOnMove() Then
        If Len(gLockedSheet) > 0 And (LCase$(gLockedSheet) <> LCase$(ws.name) Or LCase$(gLockedCell) <> LCase$(target.Address(False, False))) Then
            XQL_AutoLock_ReleaseCurrent
        End If
    End If

    ' 현재 셀에 락 획득 시도(같은 셀 재선택일 수도 있음)
    If Acquire(ws.name, target.Address(False, False)) Then
        ShadeMine ws, target
        ScheduleRefresh
    Else
        ' 획득 실패: 이동
        MsgBox "락 획득 실패(다른 사용자가 점유 중).", vbExclamation
        SafeMoveSelection ws
    End If
done:
End Sub

' 워크북 종료/수동 호출용: 현재 락 해제
Public Sub XQL_AutoLock_ReleaseCurrent()
    On Error Resume Next
    If Len(gLockedSheet) = 0 Then Exit Sub
    Dim q$, vars$, resp$
    q = "mutation($s:String!,$c:String!,$n:String!){releaseLock(sheet:$s,cell:$c,nickname:$n)}"
    vars = "{""s"":""" & gLockedSheet & """,""c"":""" & gLockedCell & """,""n"":""" & XQL_GetNickname() & """}"
    resp = XQL_GraphQL(q, vars)
    gLockedSheet = "": gLockedCell = ""
    If mRefreshAt <> 0 Then On Error Resume Next: Application.OnTime mRefreshAt, "XQL_AutoLock_RefreshTick", , False
End Sub

' 내부: 락 갱신 타이머 예약
Private Sub ScheduleRefresh()
    On Error Resume Next
    If mRefreshAt <> 0 Then Application.OnTime mRefreshAt, "XQL_AutoLock_RefreshTick", , False
    mRefreshAt = Now + (XQL_LockRefreshSec() / 86400#)
    Application.OnTime mRefreshAt, "XQL_AutoLock_RefreshTick"
End Sub

' 타이머 진입점: 같은 셀 유지 중이면 재획득으로 TTL 갱신
Public Sub XQL_AutoLock_RefreshTick()
    On Error GoTo done
    If Not XQL_AutoLockEnabled() Then Exit Sub
    If Len(gLockedSheet) = 0 Then Exit Sub
    Dim ws As Worksheet: Set ws = Nothing
    On Error Resume Next: Set ws = Sheets(gLockedSheet): On Error GoTo 0
    If ws Is Nothing Then XQL_AutoLock_ReleaseCurrent: Exit Sub

    ' 아직 같은 셀을 선택 중인지 확인
    If LCase$(ActiveSheet.name) = LCase$(gLockedSheet) And LCase$(Selection.Address(False, False)) = LCase$(gLockedCell) Then
        ' 재획득(갱신)
        Call Acquire(gLockedSheet, gLockedCell)
        ScheduleRefresh
    Else
        ' 선택이 바뀌었으면 설정에 따라 해제
        If XQL_AutoReleaseOnMove() Then XQL_AutoLock_ReleaseCurrent
    End If
done:
End Sub

' 서버 호출: acquireLock
Private Function Acquire(ByVal sheetName As String, ByVal cellAddr As String) As Boolean
    On Error GoTo bad
    Dim q$, vars$, resp$
    q = "mutation($s:String!,$c:String!,$n:String!){acquireLock(sheet:$s,cell:$c,nickname:$n)}"
    vars = "{""s"":""" & sheetName & """,""c"":""" & cellAddr & """,""n"":""" & XQL_GetNickname() & """}"
    resp = XQL_GraphQL(q, vars)
    If InStr(1, resp, "true", vbTextCompare) > 0 Then
        gLockedSheet = sheetName
        gLockedCell = cellAddr
        Acquire = True
    Else
        Acquire = False
    End If
    Exit Function
bad:
    Acquire = False
End Function

' 타인 락 여부(PresenceView에서 유지하는 gLockIndex 재사용)
Private Function IsLockedByOther(ByVal sheetName As String, ByVal rng As Range) As Boolean
    On Error GoTo no
    If XQL_WarnOnLockedSelect() = False Then IsLockedByOther = False: Exit Function
    ' PresenceView 모듈의 gLockIndex를 읽음(없으면 선택 허용)
    Dim idx As Object
    Set idx = Nothing
    ' 메모리 공유: Public 변수가 다른 모듈에 있으므로 직접 참조는 불가
    ' 대신 최신 락을 가져와 확인(짧은 쿼리)
    Dim lcks As Object: Set lcks = XQL_FetchLocks(sheetName)
    Dim i&
    For i = 1 To lcks.count
        If LCase$(CStr(lcks(i)("sheet"))) = LCase$(sheetName) And _
           LCase$(CStr(lcks(i)("cell"))) = LCase$(rng.Address(False, False)) And _
           LCase$(CStr(lcks(i)("nickname"))) <> LCase$(XQL_GetNickname()) Then
            IsLockedByOther = True: Exit Function
        End If
    Next
    IsLockedByOther = False: Exit Function
no:
    IsLockedByOther = False
End Function

' 내 락 셀은 연한 파랑으로 표시
Private Sub ShadeMine(ws As Worksheet, rng As Range)
    On Error Resume Next
    rng.Interior.Color = RGB(220, 240, 255)
End Sub

' 선택 차단 시 안전 이동(오른쪽→아래→A3)
Private Sub SafeMoveSelection(ws As Worksheet)
    On Error Resume Next
    Dim dest As Range: Set dest = Nothing
    Set dest = Intersect(ws.UsedRange, ws.Cells(Selection.row, Selection.Column + 1))
    If dest Is Nothing Then
        Set dest = Intersect(ws.UsedRange, ws.Cells(Selection.row + 1, Selection.Column))
    End If
    If dest Is Nothing Then Set dest = ws.Range("A3")
    dest.Select
End Sub

'Module XQL_Config
Option Explicit

Public Function XQL_GetApiUrl() As String
    XQL_GetApiUrl = CStr(Sheets("XQLite").Range("A1").Offset(0, 1).Value2)
End Function
Public Function XQL_GetApiKey() As String
    XQL_GetApiKey = CStr(Sheets("XQLite").Range("A2").Offset(0, 1).Value2)
End Function
Public Function XQL_GetNickname() As String
    XQL_GetNickname = CStr(Sheets("XQLite").Range("A3").Offset(0, 1).Value2)
End Function
Public Function XQL_GetDebounceSec() As Double
    XQL_GetDebounceSec = CDbl(Sheets("XQLite").Range("A5").Offset(0, 1).Value2) / 1000#
    If XQL_GetDebounceSec <= 0 Then XQL_GetDebounceSec = 2
End Function

Public Sub XQL_InitConfig()
    ' 아무 것도 안 해도 됨. 향후 초기화 훅.
End Sub

' 시트의 데이터 영역 메타(헤더/마지막열/보조열)
Public Function XQL_LastCol(ws As Worksheet) As Long
    XQL_LastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
End Function

Public Function XQL_VerCol(ws As Worksheet) As Long
    XQL_VerCol = XQL_LastCol(ws) + 1 ' 숨김: 버전
End Function
Public Function XQL_DirtyCol(ws As Worksheet) As Long
    XQL_DirtyCol = XQL_LastCol(ws) + 2 ' 숨김: dirty flag
End Function

Public Sub XQL_EnsureAuxCols(ws As Worksheet)
    Dim lc&, vcol&, dcol&
    lc = XQL_LastCol(ws)
    vcol = lc + 1
    dcol = lc + 2
    ws.Cells(1, vcol).Value = "_ver"
    ws.Cells(1, dcol).Value = "_dirty"
    ws.Columns(vcol).Hidden = True
    ws.Columns(dcol).Hidden = True
End Sub

Public Sub XQL_MarkDirty(ws As Worksheet, ByVal row As Long)
    XQL_EnsureAuxCols ws
    ws.Cells(row, XQL_DirtyCol(ws)).Value = 1
End Sub
Public Sub XQL_ClearDirty(ws As Worksheet, ByVal row As Long)
    ws.Cells(row, XQL_DirtyCol(ws)).Value = Empty
End Sub

' 안전한 문자열 이스케이프
Public Function XQL_EscapeJson(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    XQL_EscapeJson = s
End Function

Public Function XQL_GetPresenceRefreshSec() As Double
    On Error Resume Next
    XQL_GetPresenceRefreshSec = CDbl(Sheets("XQLite").Range("A8").Offset(0, 1).Value2)
    If XQL_GetPresenceRefreshSec <= 0 Then XQL_GetPresenceRefreshSec = 6
End Function

Public Function XQL_ShowLockShading() As Boolean
    On Error Resume Next
    XQL_ShowLockShading = CBool(Sheets("XQLite").Range("A9").Offset(0, 1).Value)
End Function

Public Function XQL_WarnOnLockedSelect() As Boolean
    On Error Resume Next
    XQL_WarnOnLockedSelect = CBool(Sheets("XQLite").Range("A10").Offset(0, 1).Value)
End Function

Public Function XQL_AutoLockEnabled() As Boolean
    On Error Resume Next
    XQL_AutoLockEnabled = CBool(Sheets("XQLite").Range("A11").Offset(0, 1).Value)
End Function

Public Function XQL_AutoReleaseOnMove() As Boolean
    On Error Resume Next
    XQL_AutoReleaseOnMove = CBool(Sheets("XQLite").Range("A12").Offset(0, 1).Value)
End Function

Public Function XQL_LockRefreshSec() As Double
    On Error Resume Next
    XQL_LockRefreshSec = CDbl(Sheets("XQLite").Range("A13").Offset(0, 1).Value2)
    If XQL_LockRefreshSec <= 0 Then XQL_LockRefreshSec = 5
End Function

Public Function XQL_OutboxRetrySec() As Double
    On Error Resume Next
    XQL_OutboxRetrySec = CDbl(Sheets("XQLite").Range("A14").Offset(0, 1).Value2)
    If XQL_OutboxRetrySec <= 0 Then XQL_OutboxRetrySec = 15
End Function

Public Function XQL_OutboxMaxRetry() As Long
    On Error Resume Next
    XQL_OutboxMaxRetry = CLng(Sheets("XQLite").Range("A15").Offset(0, 1).Value2)
    If XQL_OutboxMaxRetry <= 0 Then XQL_OutboxMaxRetry = 10
End Function

Public Function XQL_PullSec() As Double
    On Error Resume Next
    XQL_PullSec = CDbl(Sheets("XQLite").Range("A16").Offset(0, 1).Value2)
    If XQL_PullSec <= 0 Then XQL_PullSec = 10
End Function

Public Function XQL_PullBatch() As Long
    On Error Resume Next
    XQL_PullBatch = CLng(Sheets("XQLite").Range("A17").Offset(0, 1).Value2)
    If XQL_PullBatch <= 0 Then XQL_PullBatch = 500
End Function

Public Function XQL_PullEnabled() As Boolean
    On Error Resume Next
    XQL_PullEnabled = CBool(Sheets("XQLite").Range("A18").Offset(0, 1).Value)
End Function

Public Function XQL_ProtectPassword() As String
    On Error Resume Next
    Dim v: v = Sheets("XQLite").Range("A22").Offset(0, 1).Value
    If Len(CStr(v)) = 0 Then XQL_ProtectPassword = "xql" Else XQL_ProtectPassword = CStr(v)
End Function

Public Function XQL_PermsAutoApply() As Boolean
    On Error Resume Next
    XQL_PermsAutoApply = CBool(Sheets("XQLite").Range("A23").Offset(0, 1).Value)
End Function

'Module XQL_ConflictResolver
' === Module: XQL_ConflictResolver ===
Option Explicit

Private Const SHEET_RES As String = "XQLite_Resolve"
Private Const CONFLICT_RGB As Long = &HB4FFFF ' = RGB(180,255,255)? Nope, Excel uses BGR. 안전하게 함수로 비교.
' 실제 비교는 RGB(255,255,180) 직접 사용.

' ─────────────────────────────────────────────────────────────
' 엔트리

Public Sub XQL_Resolve_ScanConflicts()
    Dim ws As Worksheet: Set ws = ActiveSheet
    If Not IsDataSheet(ws) Then
        MsgBox "데이터 시트에서 실행하세요.", vbExclamation
        Exit Sub
    End If

    Dim ids As Object, cols As Object
    Set ids = CreateObject("Scripting.Dictionary")   ' id -> 1
    Set cols = CreateObject("Scripting.Dictionary")  ' "id|col" -> 1

    Dim lastRow&, lastCol&, r&, c&, idv&
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    lastCol = XQL_LastCol(ws)

    For r = 3 To lastRow
        idv = CLng(val(ws.Cells(r, 1).Value2))
        If idv <= 0 Then GoTo nextr
        For c = 1 To lastCol
            If ws.Cells(r, c).Interior.Color = RGB(255, 255, 180) Then
                ids(CStr(idv)) = 1
                cols(CStr(idv) & "|" & CStr(ws.Cells(1, c).Value2)) = 1
            End If
        Next c
nextr:
    Next r

    If ids.count = 0 Then
        MsgBox "충돌(노란색) 셀이 없습니다.", vbInformation
        Exit Sub
    End If

    ' 서버에서 해당 id들의 최신 행을 가져옴
    Dim listIds As String: listIds = JoinDictKeysAsCsv(ids)
    Dim serverRows As Object: Set serverRows = FetchServerRowsByIds(ws.name, listIds)
    If serverRows Is Nothing Or serverRows.count = 0 Then
        MsgBox "서버에서 해당 행을 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If

    ' 비교표 렌더
    RenderResolveSheet ws, cols, serverRows
    MsgBox "충돌 목록을 " & SHEET_RES & " 시트에 생성했습니다.", vbInformation
End Sub

Public Sub XQL_Resolve_AcceptServerAll()
    Dim rs As Worksheet: Set rs = EnsureResolveSheet()
    Dim last&, r&
    last = rs.Cells(rs.rows.count, 1).End(xlUp).row
    For r = 2 To last
        rs.Cells(r, 5).Value = "SERVER"
    Next
End Sub

Public Sub XQL_Resolve_AcceptExcelAll()
    Dim rs As Worksheet: Set rs = EnsureResolveSheet()
    Dim last&, r&
    last = rs.Cells(rs.rows.count, 1).End(xlUp).row
    For r = 2 To last
        rs.Cells(r, 5).Value = "EXCEL"
    Next
End Sub

Public Sub XQL_Resolve_ApplyDecisions()
    Dim rs As Worksheet: Set rs = EnsureResolveSheet()
    Dim srcSheet As String: srcSheet = CStr(rs.Range("H1").Value2)
    If Len(srcSheet) = 0 Then
        MsgBox "원본 시트 정보가 없습니다. Scan Conflicts 를 먼저 실행하세요.", vbExclamation
        Exit Sub
    End If
    Dim ws As Worksheet
    On Error Resume Next: Set ws = Sheets(srcSheet): On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "원본 시트를 찾을 수 없습니다: " & srcSheet, vbExclamation
        Exit Sub
    End If

    Dim last&, r&, id&, col$, decision$, sver&, appliedCnt&, excelCnt&, serverCnt&
    last = rs.Cells(rs.rows.count, 1).End(xlUp).row

    ' 행별 서버 row_version 최신화 맵
    Dim verMap As Object: Set verMap = CreateObject("Scripting.Dictionary")

    For r = 2 To last
        id = CLng(val(rs.Cells(r, 1).Value2))
        If id <= 0 Then GoTo nextr
        col = CStr(rs.Cells(r, 2).Value2)
        decision = UCase$(Trim$(CStr(rs.Cells(r, 5).Value2)))
        sver = CLng(val(Nz(rs.Cells(r, 6).Value2, 0)))
        If sver > 0 Then verMap(CStr(id)) = sver

        If decision = "SERVER" Then
            ' 서버 값으로 엑셀 갱신 + _ver 최신화 + _dirty 클리어
            ApplyServerToExcel ws, CLng(id), col, rs.Cells(r, 4).Value
            serverCnt = serverCnt + 1
        ElseIf decision = "EXCEL" Then
            ' 엑셀 값을 그대로 유지하고 나중에 Outbox로 덮어쓰기(업서트)
            excelCnt = excelCnt + 1
        End If
nextr:
    Next r

    ' _ver 최신화 후 Outbox 큐잉(Excel로 덮어씌우기 대상만)
    Dim k As Variant, rr&, vcol&, dcol&
    vcol = XQL_VerCol(ws): dcol = XQL_DirtyCol(ws)
    For Each k In verMap.keys
        rr = FindRowById(ws, CLng(k))
        If rr > 0 And vcol > 0 Then ws.Cells(rr, vcol).Value = CLng(verMap(k))
    Next k

    ' EXCEL 결정을 Outbox로
    Dim enq As Object: Set enq = BuildExcelWinningRows(rs)
    Dim i&
    For i = 0 To enq.count - 1
        EnqueueUpsertRow ws, CLng(enq(i))
        appliedCnt = appliedCnt + 1
    Next i

    ' 성공 알림
    MsgBox "적용 완료:" & vbCrLf & _
           "- 서버값 채택: " & serverCnt & "개" & vbCrLf & _
           "- 엑셀값 덮어쓰기(Outbox 큐): " & appliedCnt & "행" & vbCrLf & _
           "※ Outbox가 자동으로 재시도/전송합니다.", vbInformation
End Sub

' ─────────────────────────────────────────────────────────────
' 내부 구현

Private Function IsDataSheet(ws As Worksheet) As Boolean
    Dim n$: n = ws.name
    IsDataSheet = Not (n = "XQLite" Or n = "XQLite_Conflicts" Or n = "XQLite_Presence" Or n = "XQLite_EnumCache" Or n = "XQLite_Enums" Or n = "XQLite_Check" Or n = "XQLite_Query" Or n = "XQLite_Audit" Or n = "XQLite_Outbox" Or n = SHEET_RES)
End Function

Private Function EnsureResolveSheet() As Worksheet
    On Error Resume Next
    Set EnsureResolveSheet = Sheets(SHEET_RES)
    On Error GoTo 0
    If EnsureResolveSheet Is Nothing Then
        Set EnsureResolveSheet = Sheets.add(After:=Sheets(Sheets.count))
        EnsureResolveSheet.name = SHEET_RES
    End If
End Function

Private Sub RenderResolveSheet(srcWs As Worksheet, cols As Object, serverRows As Object)
    Dim rs As Worksheet: Set rs = EnsureResolveSheet()
    rs.Cells.Clear

    rs.Range("A1:F1").Value = Array("id", "column", "excel_value", "server_value", "decision(EXCEL|SERVER)", "server_row_version")
    rs.rows(1).Font.Bold = True
    rs.Range("H1").Value = srcWs.name ' 원본 시트명 메타 저장

    ' 데이터 유효성: E열에 드롭다운
    With rs.Range("E2:E1000000")
        .Validation.Delete
        .Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="EXCEL,SERVER"
    End With

    Dim r As Long: r = 2
    Dim key As Variant, id$, col$, excelV, serverV, sver&

    For Each key In cols.keys
        id = Split(CStr(key), "|")(0)
        col = Split(CStr(key), "|")(1)

        ' 엑셀 값
        excelV = GetExcelCellValueByIdCol(srcWs, CLng(id), col)
        ' 서버 값 + row_version
        serverV = GetServerCellValueByIdCol(serverRows, CLng(id), col)
        sver = GetServerRowVersion(serverRows, CLng(id))

        rs.Cells(r, 1).Value = CLng(id)
        rs.Cells(r, 2).Value = col
        rs.Cells(r, 3).Value = excelV
        rs.Cells(r, 4).Value = serverV
        rs.Cells(r, 5).Value = ""          ' 사용자가 선택
        rs.Cells(r, 6).Value = sver

        r = r + 1
    Next key

    rs.Columns("A:F").AutoFit
    rs.Activate
End Sub

' 원본 시트에서 (id, col) 값 가져오기
Private Function GetExcelCellValueByIdCol(ws As Worksheet, ByVal id As Long, ByVal colName As String) As Variant
    Dim rr&, cc&
    rr = FindRowById(ws, id)
    If rr <= 0 Then Exit Function
    cc = FindColumnByHeader(ws, colName)
    If cc <= 0 Then Exit Function
    GetExcelCellValueByIdCol = ws.Cells(rr, cc).Value
End Function

' 서버 결과에서 (id, col) 값
Private Function GetServerCellValueByIdCol(rows As Object, ByVal id As Long, ByVal colName As String) As Variant
    On Error GoTo done
    Dim i&, row As Object
    For i = 1 To rows.count
        Set row = rows(i)
        If CLng(Nz(row("id"), 0)) = id Then
            If row.Exists(colName) Then GetServerCellValueByIdCol = row(colName)
            Exit Function
        End If
    Next i
done:
End Function

Private Function GetServerRowVersion(rows As Object, ByVal id As Long) As Long
    On Error GoTo done
    Dim i&, row As Object
    For i = 1 To rows.count
        Set row = rows(i)
        If CLng(Nz(row("id"), 0)) = id Then
            GetServerRowVersion = CLng(Nz(row("row_version"), 0))
            Exit Function
        End If
    Next i
done:
    GetServerRowVersion = 0
End Function

' 서버에서 id IN (...) 조회
Private Function FetchServerRowsByIds(ByVal table As String, ByVal idsCsv As String) As Object
    On Error GoTo bad
    Dim wr$, q$, vars$, resp$, parsed As Object
    wr = "id IN (" & idsCsv & ")"
    q = "query($t:String!,$w:String){rows(table:$t,whereRaw:$w,limit:100000){rows}}"
    vars = "{""t"":""" & table & """,""w"":""" & EscapeSqlStr(wr) & """}"
    resp = XQL_GraphQL(q, vars)
    If XQL_HasErrors(resp) Then GoTo bad
    Set parsed = JsonConverter.ParseJson(resp)
    Set FetchServerRowsByIds = parsed("data")("rows")("rows")
    Exit Function
bad:
    Set FetchServerRowsByIds = Nothing
End Function

Private Function JoinDictKeysAsCsv(d As Object) As String
    Dim k As Variant, s$: s = ""
    For Each k In d.keys
        If Len(s) > 0 Then s = s & ","
        s = s & CStr(k)
    Next
    JoinDictKeysAsCsv = s
End Function

' 서버값 선택 반영
Private Sub ApplyServerToExcel(ws As Worksheet, ByVal id As Long, ByVal colName As String, ByVal serverVal As Variant)
    Dim rr&, cc&, vcol&, dcol&
    rr = FindRowById(ws, id)
    If rr <= 0 Then Exit Sub
    cc = FindColumnByHeader(ws, colName)
    If cc <= 0 Then Exit Sub

    ws.Cells(rr, cc).Value = serverVal
    ' 충돌 색 제거
    ws.Cells(rr, cc).Interior.ColorIndex = xlNone

    ' 더티 마크 제거(행 단위)
    dcol = XQL_DirtyCol(ws)
    If dcol > 0 Then ws.Cells(rr, dcol).Value = Empty
End Sub

' "EXCEL" 결정들을 행 인덱스로 모음(중복 제거)
Private Function BuildExcelWinningRows(rs As Worksheet) As Object
    Dim setR As Object: Set setR = CreateObject("Scripting.Dictionary")
    Dim last&, r&, id&, col$, srcSheet$, ws As Worksheet, rr&
    last = rs.Cells(rs.rows.count, 1).End(xlUp).row
    srcSheet = CStr(rs.Range("H1").Value2)
    On Error Resume Next: Set ws = Sheets(srcSheet): On Error GoTo 0
    If ws Is Nothing Then Set BuildExcelWinningRows = CreateObject("System.Collections.ArrayList"): Exit Function

    For r = 2 To last
        If UCase$(Trim$(CStr(rs.Cells(r, 5).Value2))) = "EXCEL" Then
            id = CLng(val(rs.Cells(r, 1).Value2))
            rr = FindRowById(ws, id)
            If rr > 0 Then setR(CStr(rr)) = 1
            ' 엑셀 주도이므로 충돌 색만 먼저 제거
            Dim cc&: cc = FindColumnByHeader(ws, CStr(rs.Cells(r, 2).Value2))
            If cc > 0 Then ws.Cells(rr, cc).Interior.ColorIndex = xlNone
        End If
    Next r

    Dim list As Object: Set list = CreateObject("System.Collections.ArrayList")
    Dim k As Variant
    For Each k In setR.keys: list.add CLng(k): Next
    Set BuildExcelWinningRows = list
End Function

' 공용 유틸

Public Function FindRowById(ws As Worksheet, ByVal id As Long) As Long
    Dim lastRow&, r&
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = 3 To lastRow
        If CLng(val(ws.Cells(r, 1).Value2)) = id Then FindRowById = r: Exit Function
    Next
    FindRowById = 0
End Function

Public Function FindColumnByHeader(ws As Worksheet, ByVal header As String) As Long
    Dim lastCol&, c&
    lastCol = XQL_LastCol(ws)
    For c = 1 To lastCol
        If LCase$(CStr(ws.Cells(1, c).Value2)) = LCase$(header) Then
            FindColumnByHeader = c: Exit Function
        End If
    Next
    FindColumnByHeader = 0
End Function

Private Function EscapeSqlStr(ByVal s As String) As String
    EscapeSqlStr = Replace(s, """", "\""") ' for JSON-bound whereRaw
End Function

Private Function Nz(ByVal v As Variant, ByVal fb As Variant) As Variant
    If IsEmpty(v) Or v = "" Or v Is Nothing Then Nz = fb Else Nz = v
End Function

'Module XQL_Conflicts
Option Explicit

Private Const SHEET_CF As String = "XQLite_Conflicts"

' 충돌 스캔 → 시트에 목록화
Public Sub XQL_Conflicts_Scan()
    Dim ws As Worksheet, cf As Worksheet, r&, c&, lastRow&, lastCol&, hdr$, localV, serverV$, hasAny As Boolean

    Application.ScreenUpdating = False

    ' 준비: 시트 생성/초기화
    On Error Resume Next
    Set cf = Sheets(SHEET_CF)
    On Error GoTo 0
    If cf Is Nothing Then
        Set cf = Sheets.add(After:=Sheets(Sheets.count))
        cf.name = SHEET_CF
    End If
    cf.Cells.Clear

    ' 헤더
    cf.Range("A1:F1").Value = Array("Sheet", "Row", "Col", "Header", "Local", "Server")
    cf.Range("G1").Value = "Choice(server/local)"
    cf.rows(1).Font.Bold = True

    Dim outR&: outR = 2

    ' 모든 데이터 시트 순회
    For Each ws In ThisWorkbook.Worksheets
        If ws.name <> "XQLite" And ws.name <> SHEET_CF Then
            lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
            lastCol = XQL_LastCol(ws)

            For r = 2 To lastRow
                For c = 1 To lastCol
                    If ws.Cells(r, c).Interior.Color = RGB(255, 255, 180) Then ' 노란색 = 충돌
                        hdr = CStr(ws.Cells(1, c).Value2)
                        localV = ws.Cells(r, c).Value
                        serverV = ExtractServerFromComment(ws.Cells(r, c))

                        cf.Cells(outR, 1).Value = ws.name
                        cf.Cells(outR, 2).Value = r
                        cf.Cells(outR, 3).Value = c
                        cf.Cells(outR, 4).Value = hdr
                        cf.Cells(outR, 5).Value = localV
                        cf.Cells(outR, 6).Value = serverV
                        cf.Cells(outR, 7).Value = "" ' 선택 칸

                        outR = outR + 1
                        hasAny = True
                    End If
                Next c
            Next r
        End If
    Next ws

    If hasAny Then
        cf.Columns("A:G").AutoFit
        cf.Activate
        MsgBox "충돌 항목을 나열했습니다. G열에 'server' 또는 'local'을 입력 후 '적용'을 실행하세요.", vbInformation
    Else
        cf.Activate
        MsgBox "충돌 항목이 없습니다.", vbInformation
    End If

    Application.ScreenUpdating = True
End Sub

' 선택 결과 적용: server → 서버값으로 덮어쓰기 / local → 현 로컬 유지+강제 푸시
Public Sub XQL_Conflicts_Apply()
    Dim cf As Worksheet, lastRow&, r&, sheetName$, rowN&, colN&, choice$, ws As Worksheet
    On Error Resume Next
    Set cf = Sheets(SHEET_CF)
    On Error GoTo 0
    If cf Is Nothing Then
        MsgBox SHEET_CF & " 시트가 없습니다. 먼저 '스캔'을 실행하세요.", vbExclamation
        Exit Sub
    End If

    lastRow = cf.Cells(cf.rows.count, 1).End(xlUp).row
    If lastRow < 2 Then
        MsgBox "적용할 항목이 없습니다.", vbInformation
        Exit Sub
    End If

    ' 로컬 푸시 대상 행을 시트별로 모아서 한 번에 업서트
    Dim groups As Object: Set groups = CreateObject("Scripting.Dictionary")

    Application.ScreenUpdating = False

    For r = 2 To lastRow
        sheetName = CStr(cf.Cells(r, 1).Value2)
        rowN = CLng(cf.Cells(r, 2).Value2)
        colN = CLng(cf.Cells(r, 3).Value2)
        choice = LCase$(Trim$(CStr(cf.Cells(r, 7).Value2)))

        If choice <> "server" And choice <> "local" Then
            ' 미선택은 건너뜀
        Else
            Set ws = Sheets(sheetName)
            If choice = "server" Then
                ' 서버값 채택: 셀에 Server값 반영 + 표시 제거
                ws.Cells(rowN, colN).Value = cf.Cells(r, 6).Value
                ClearConflictMark ws.Cells(rowN, colN)
            Else
                ' 로컬 유지: 표시만 제거하고 업서트 대기열에 추가
                ClearConflictMark ws.Cells(rowN, colN)
                Dim list As Object
                If Not groups.Exists(sheetName) Then
                    Set list = CreateObject("System.Collections.ArrayList")
                    groups.add sheetName, list
                Else
                    Set list = groups(sheetName)
                End If
                If Not list.Contains(rowN) Then list.add rowN
            End If
        End If
    Next r

    ' 로컬 유지 선택건 강제 업서트
    Dim k As Variant, arr
    For Each k In groups.keys
        arr = groups(k).ToArray()
        XQL_UpsertRows Sheets(CStr(k)), arr
    Next

    Application.ScreenUpdating = True
    MsgBox "적용 완료", vbInformation
End Sub

' 보조: 모든 충돌을 Server로/Local로 일괄 설정
Public Sub XQL_Conflicts_MarkAllServer()
    MarkAllChoice "server"
End Sub
Public Sub XQL_Conflicts_MarkAllLocal()
    MarkAllChoice "local"
End Sub

' ── internal helpers ───────────────────────────────────

Private Sub MarkAllChoice(ByVal v As String)
    Dim cf As Worksheet, lastRow&, r&
    On Error Resume Next
    Set cf = Sheets(SHEET_CF)
    On Error GoTo 0
    If cf Is Nothing Then Exit Sub
    lastRow = cf.Cells(cf.rows.count, 1).End(xlUp).row
    If lastRow < 2 Then Exit Sub
    For r = 2 To lastRow
        cf.Cells(r, 7).Value = v
    Next
End Sub

Private Function ExtractServerFromComment(ByVal cell As Range) As String
    On Error Resume Next
    Dim t$: t = ""
    If Not cell.Comment Is Nothing Then t = cell.Comment.Text
    If Len(t) = 0 Then
        ExtractServerFromComment = ""
        Exit Function
    End If
    ' "Server: " 접두부 제거
    If Left$(t, 8) = "Server: " Then t = Mid$(t, 9)
    ExtractServerFromComment = t
End Function

Private Sub ClearConflictMark(ByVal cell As Range)
    On Error Resume Next
    cell.Interior.ColorIndex = xlNone
    If Not cell.Comment Is Nothing Then cell.Comment.Delete
End Sub

'Module XQL_Dashboard
' === Module: XQL_Dashboard ===
Option Explicit

Public Sub XQL_ShowDashboard()
    On Error Resume Next
    If frmXQLite Is Nothing Then
        Dim f As New frmXQLite
        f.Show       ' 모델리스 (ShowModal=False)
    Else
        frmXQLite.Show
    End If
End Sub

'Module XQL_DeleteRestore
' === Module: XQL_DeleteRestore ===
Option Explicit

' 선택한 행의 id들을 모아 서버에 soft delete 요청
Public Sub XQL_Delete_SelectedRows()
    On Error GoTo eh
    Dim ws As Worksheet: Set ws = ActiveSheet
    If UCase$(CStr(ws.Cells(1, 1).Value2)) <> "ID" Then
        MsgBox "A1 헤더가 'id'가 아닙니다.", vbExclamation: Exit Sub
    End If

    Dim ids As Object: Set ids = CollectSelectedIds(ws)
    If ids Is Nothing Or ids.count = 0 Then
        MsgBox "선택된 행에 유효한 id가 없습니다.", vbInformation: Exit Sub
    End If

    If MsgBox(ids.count & "개 행을 삭제 표시 하시겠습니까? (soft delete)", vbQuestion + vbOKCancel) <> vbOK Then Exit Sub

    Dim arr$, i&
    arr = "["
    For i = 0 To ids.count - 1
        arr = arr & CStr(ids(i))
        If i < ids.count - 1 Then arr = arr & ","
    Next i
    arr = arr & "]"

    Dim q$, vars$, resp$, mx&
    q = "mutation($t:String!,$ids:[Int!]!,$a:String!){deleteRows(table:$t,ids:$ids,actor:$a){max_row_version affected errors}}"
    vars = "{""t"":""" & ws.name & """,""ids"":" & arr & ",""a"":""" & XQL_GetNickname() & """}"
    resp = XQL_GraphQL(q, vars)
    If XQL_HasErrors(resp) Then
        MsgBox "deleteRows 실패: " & resp, vbExclamation: Exit Sub
    End If

    mx = XQL_ExtractMaxRowVersion(resp)
    If mx > 0 Then Sheets("XQLite").Range("A7").Offset(0, 1).Value = mx

    ' UI: 줄취소선+그레이, _ver 갱신
    Dim vcol&: vcol = XQL_VerCol(ws)
    For i = 0 To ids.count - 1
        Dim r&: r = FindRowById(ws, CLng(ids(i)))
        If r > 0 Then
            ws.rows(r).Font.Strikethrough = True
            ws.rows(r).Interior.Color = RGB(240, 240, 240)
            If mx > 0 Then ws.Cells(r, vcol).Value = mx
        End If
    Next i

    MsgBox "삭제(soft) 처리 완료", vbInformation
    Exit Sub
eh:
    MsgBox "XQL_Delete_SelectedRows 오류: " & Err.Description, vbExclamation
End Sub

' 선택한 행의 id들을 모아 deleted=0 으로 복구 시도
' 주의: 서버가 meta 컬럼(deleted) 업서트를 허용하지 않으면 실패할 수 있음.
' 그 경우 서버에 'undeleteRows' 뮤테이션을 추가하는 것이 가장 깔끔합니다.
Public Sub XQL_Restore_SelectedRows()
    On Error GoTo eh
    Dim ws As Worksheet: Set ws = ActiveSheet
    If UCase$(CStr(ws.Cells(1, 1).Value2)) <> "ID" Then
        MsgBox "A1 헤더가 'id'가 아닙니다.", vbExclamation: Exit Sub
    End If

    Dim ids As Object: Set ids = CollectSelectedIds(ws)
    If ids Is Nothing Or ids.count = 0 Then
        MsgBox "선택된 행에 유효한 id가 없습니다.", vbInformation: Exit Sub
    End If

    If MsgBox(ids.count & "개 행을 복구(deleted=0) 하시겠습니까?", vbQuestion + vbOKCancel) <> vbOK Then Exit Sub

    ' UpsertRowInput[] 직조: {"id":X, "base_row_version": _ver, "data":{"deleted":0}}
    Dim vcol&, i&, buf$, r&, basev&
    vcol = XQL_VerCol(ws)
    buf = "["

    For i = 0 To ids.count - 1
        r = FindRowById(ws, CLng(ids(i)))
        If r <= 0 Then GoTo cont
        basev = 0
        On Error Resume Next: basev = CLng(ws.Cells(r, vcol).Value2): On Error GoTo 0
        buf = buf & "{""id"":" & CLng(ids(i)) & ",""base_row_version"":" & basev & ",""data"":{""deleted"":0}}"
        If i < ids.count - 1 Then buf = buf & ","
cont:
    Next i
    buf = buf & "]"

    Dim q$, vars$, resp$, mx&
    q = "mutation($t:String!,$rows:[UpsertRowInput!]!,$a:String!){upsertRows(table:$t,rows:$rows,actor:$a){max_row_version affected errors conflicts}}"
    vars = "{""t"":""" & ws.name & """,""rows"":" & buf & ",""a"":""" & XQL_GetNickname() & """}"
    resp = XQL_GraphQL(q, vars)

    If XQL_HasErrors(resp) Then
        MsgBox "복구(upsert deleted=0) 실패: " & resp & vbCrLf & _
               "서버가 meta 컬럼 업서트를 막는 경우 'undeleteRows' 뮤테이션을 추가해야 합니다.", vbExclamation
        Exit Sub
    End If

    mx = XQL_ExtractMaxRowVersion(resp)
    If mx > 0 Then Sheets("XQLite").Range("A7").Offset(0, 1).Value = mx

    ' UI: 줄취소선/그레이 해제
    For i = 0 To ids.count - 1
        r = FindRowById(ws, CLng(ids(i)))
        If r > 0 Then
            ws.rows(r).Font.Strikethrough = False
            ws.rows(r).Interior.ColorIndex = xlNone
            If mx > 0 Then ws.Cells(r, vcol).Value = mx
        End If
    Next i

    MsgBox "복구 처리 완료", vbInformation
    Exit Sub
eh:
    MsgBox "XQL_Restore_SelectedRows 오류: " & Err.Description, vbExclamation
End Sub

' 선택 영역에서 id 수집 (중복 제거)
Private Function CollectSelectedIds(ws As Worksheet) As Object
    Dim setIds As Object: Set setIds = CreateObject("Scripting.Dictionary")
    Dim area As Range, r As Range, idv
    For Each area In Selection.Areas
        For Each r In area.rows
            If r.row >= 2 Then
                idv = ws.Cells(r.row, 1).Value2
                If Len(idv) > 0 Then setIds(CStr(CLng(val(idv)))) = 1
            End If
        Next r
    Next area
    Dim list As Object: Set list = CreateObject("System.Collections.ArrayList")
    Dim k As Variant
    For Each k In setIds.keys
        list.add CLng(k)
    Next
    Set CollectSelectedIds = list
End Function

' id 행 찾기 (없으면 0)
Private Function FindRowById(ws As Worksheet, ByVal id As Long) As Long
    Dim lastRow&, r&
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        If CLng(val(ws.Cells(r, 1).Value2)) = id Then FindRowById = r: Exit Function
    Next
    FindRowById = 0
End Function

'Module XQL_Enums
' === Module: XQL_Enums ===
Option Explicit

' 매핑/캐시 시트
Private Const SHEET_MAP As String = "XQLite_Enums"
Private Const SHEET_CACHE As String = "XQLite_EnumCache"

' 메모리 매핑: "sheet|col" -> Dictionary( mode, cacheName, valueCol, labelCol )
Private gEnumMap As Scripting.Dictionary

' ─────────────────────────────────────────────────────────────────
' 1) 설정 시트 생성

Public Sub XQL_Enums_Setup()
    Dim ws As Worksheet
    On Error Resume Next: Set ws = Sheets(SHEET_MAP): On Error GoTo 0
    If ws Is Nothing Then
        Set ws = Sheets.add(After:=Sheets(Sheets.count))
        ws.name = SHEET_MAP
    Else
        ws.Cells.Clear
    End If

    With ws
        .Range("A1:I1").Value = Array( _
            "target_sheet", _          ' 필수: 적용할 시트명
            "target_column", _         ' 필수: 적용할 컬럼(헤더명)
            "mode(LABEL|ID_LABEL)", _  ' LABEL=라벨 저장, ID_LABEL=라벨 표시+<col>_id에 ID자동기입
            "source_table", _          ' 필수: 참조 테이블명
            "source_value_col", _      ' 기본 id
            "source_label_col", _      ' 예: name
            "whereRaw(optional)", _    ' 예: deleted=0
            "orderBy(optional)", _     ' 예: name ASC
            "limit(optional)" _        ' 기본 5000
        )
        .rows(1).Font.Bold = True
        .Columns("A:I").ColumnWidth = 22

        ' 예시 한 줄
        .Range("A2:I2").Value = Array( _
            "items", _
            "rarity", _
            "LABEL", _
            "rarity_ref", _
            "id", _
            "name", _
            "deleted=0", _
            "sort_order ASC", _
            5000 _
        )
        .Range("A3:I3").Value = Array( _
            "items", _
            "weapon_type", _
            "ID_LABEL", _
            "weapon_type_ref", _
            "id", _
            "name", _
            "deleted=0", _
            "name ASC", _
            500 _
        )
    End With

    ' 캐시 시트 생성/숨김
    Dim c As Worksheet
    On Error Resume Next: Set c = Sheets(SHEET_CACHE): On Error GoTo 0
    If c Is Nothing Then
        Set c = Sheets.add(After:=Sheets(Sheets.count))
        c.name = SHEET_CACHE
        c.Visible = xlSheetVeryHidden
    End If

    MsgBox SHEET_MAP & " 시트를 만들었습니다." & vbCrLf & _
           "행을 추가/편집 후, XQL_Enums_RefreshAll 을 실행하세요.", vbInformation
End Sub

' ─────────────────────────────────────────────────────────────────
' 2) 새로고침 엔트리

Public Sub XQL_Enums_RefreshAll()
    BuildEnums Nothing
End Sub

Public Sub XQL_Enums_RefreshActiveSheet()
    BuildEnums ActiveSheet
End Sub

' ─────────────────────────────────────────────────────────────────
' 3) SheetChange 훅: ID_LABEL 모드 처리

Public Sub XQL_Enums_HandleChange(ByVal ws As Worksheet, ByVal target As Range)
    On Error GoTo done
    If gEnumMap Is Nothing Then Exit Sub

    Dim key$, map As Object
    Dim c As Range
    For Each c In target.Cells
        If c.row >= 3 Then
            key = LCase$(ws.name) & "|" & LCase$(HeaderAt(ws, c.Column))
            If gEnumMap.Exists(key) Then
                Set map = gEnumMap(key)
                If UCase$(CStr(map("mode"))) = "ID_LABEL" Then
                    ' 라벨 -> ID 매핑 후 <col>_id 셀에 기록
                    Dim label$, idVal, idCol&
                    label = CStr(c.Value)
                    If Len(label) > 0 Then
                        idVal = FindIdByLabel(CStr(map("cacheName")), label)
                        If Not IsEmpty(idVal) Then
                            idCol = FindColumnByHeader(ws, HeaderAt(ws, c.Column) & "_id")
                            If idCol > 0 Then
                                ws.Cells(c.row, idCol).Value = idVal
                                ' 더티 마킹
                                XQL_MarkDirty ws, c.row
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next c
done:
End Sub

' ─────────────────────────────────────────────────────────────────
' 내부 구현

Private Sub BuildEnums(ByVal onlyWs As Worksheet)
    Dim mapWs As Worksheet
    On Error Resume Next: Set mapWs = Sheets(SHEET_MAP): On Error GoTo 0
    If mapWs Is Nothing Then
        MsgBox SHEET_MAP & " 시트가 없습니다. 먼저 XQL_Enums_Setup 을 실행하세요.", vbExclamation
        Exit Sub
    End If

    If gEnumMap Is Nothing Then Set gEnumMap = New Scripting.Dictionary Else gEnumMap.RemoveAll

    Dim lastRow&, r&, tSheet$, tCol$, mode$, sTable$, vcol$, lCol$, where$, order$, lim&
    lastRow = mapWs.Cells(mapWs.rows.count, 1).End(xlUp).row

    For r = 2 To lastRow
        tSheet = Trim$(CStr(mapWs.Cells(r, 1).Value2))
        tCol = Trim$(CStr(mapWs.Cells(r, 2).Value2))
        mode = UCase$(Trim$(CStr(mapWs.Cells(r, 3).Value2)))
        sTable = Trim$(CStr(mapWs.Cells(r, 4).Value2))
        vcol = Trim$(CStr(mapWs.Cells(r, 5).Value2)): If Len(vcol) = 0 Then vcol = "id"
        lCol = Trim$(CStr(mapWs.Cells(r, 6).Value2)): If Len(lCol) = 0 Then lCol = "name"
        where = Trim$(CStr(mapWs.Cells(r, 7).Value2))
        order = Trim$(CStr(mapWs.Cells(r, 8).Value2))
        lim = CLng(IIf(Len(Trim$(CStr(mapWs.Cells(r, 9).Value2))) = 0, 5000, mapWs.Cells(r, 9).Value2))

        If Len(tSheet) = 0 Or Len(tCol) = 0 Or Len(sTable) = 0 Then GoTo cont

        If Not onlyWs Is Nothing Then
            If LCase$(onlyWs.name) <> LCase$(tSheet) Then GoTo cont
        End If

        ' 1) 서버에서 (value,label) 목록 가져오기
        Dim pairs As Object: Set pairs = FetchPairs(sTable, vcol, lCol, where, order, lim)

        ' 2) 캐시 테이블/이름 생성
        Dim cacheName$: cacheName = "_enum_" & SanitizeName(tSheet) & "_" & SanitizeName(tCol)
        WriteCache cacheName, pairs

        ' 3) 대상 시트/컬럼에 유효성 리스트 적용
        ApplyValidation tSheet, tCol, cacheName

        ' 4) 매핑 등록
        Dim key$: key = LCase$(tSheet) & "|" & LCase$(tCol)
        Dim m As Object: Set m = CreateObject("Scripting.Dictionary")
        m("mode") = mode: m("cacheName") = cacheName: m("valueCol") = vcol: m("labelCol") = lCol
        gEnumMap(key) = m
cont:
    Next r

    If onlyWs Is Nothing Then
        MsgBox "ENUM 유효성 목록을 갱신했습니다.", vbInformation
    End If
End Sub

' 서버에서 value/label 쌍 조회
Private Function FetchPairs(ByVal table As String, ByVal vcol As String, ByVal lCol As String, ByVal whereRaw As String, ByVal orderBy As String, ByVal limitN As Long) As Object
    Dim q$, vars$, resp$, parsed As Object, rows As Object
    q = "query($t:String!,$w:String,$o:String,$l:Int){rows(table:$t,whereRaw:$w,orderBy:$o,limit:$l){rows}}"
    vars = "{""t"":""" & table & """,""w"":" & JStrOrNull(whereRaw) & ",""o"":" & JStrOrNull(orderBy) & ",""l"":" & CStr(limitN) & "}"
    resp = XQL_GraphQL(q, vars)
    Set parsed = JsonConverter.ParseJson(resp)
    Set rows = parsed("data")("rows")("rows")

    Dim out As Object: Set out = CreateObject("System.Collections.ArrayList")
    Dim i&, v, lbl
    If Not rows Is Nothing Then
        For i = 1 To rows.count
            v = Nz(rows(i)(vcol), "")
            lbl = Nz(rows(i)(lCol), "")
            If Len(CStr(lbl)) > 0 Then
                Dim pair As Object: Set pair = CreateObject("Scripting.Dictionary")
                pair("v") = v: pair("l") = lbl
                out.add pair
            End If
        Next i
    End If
    Set FetchPairs = out
End Function

' 캐시에 [value,label] 작성 + 이름 정의
Private Sub WriteCache(ByVal cacheName As String, ByVal pairs As Object)
    Dim c As Worksheet
    On Error Resume Next: Set c = Sheets(SHEET_CACHE): On Error GoTo 0
    If c Is Nothing Then Exit Sub

    c.Visible = xlSheetVisible
    c.Cells.Clear

    ' 캐시는 매핑별로 "테이블"처럼 분리하지 않고, 이름정의로 범위를 분리합니다.
    ' → 기존 이름정의를 제거 후 다시 만듭니다.
    On Error Resume Next: ThisWorkbook.Names(cacheName).Delete: On Error GoTo 0

    ' 값을 A:B에 씁니다.
    Dim i&, base&: base = 1
    c.Cells(base, 1).Value = "value": c.Cells(base, 2).Value = "label"
    For i = 0 To pairs.count - 1
        c.Cells(base + i + 1, 1).Value = pairs(i)("v")
        c.Cells(base + i + 1, 2).Value = pairs(i)("l")
    Next i

    ' label 열에 이름정의 생성
    Dim lastRow&: lastRow = base + pairs.count
    Dim ref$: ref = "'" & SHEET_CACHE & "'!" & c.Range(c.Cells(2, 2), c.Cells(lastRow, 2)).Address
    ThisWorkbook.Names.add name:=cacheName, RefersTo:=("=" & ref), Visible:=True

    c.Visible = xlSheetVeryHidden
End Sub

' 대상 시트/컬럼에 데이터 유효성(리스트=캐시 이름정의) 적용
Private Sub ApplyValidation(ByVal sheetName As String, ByVal colHeader As String, ByVal cacheName As String)
    Dim ws As Worksheet: Set ws = Sheets(sheetName)
    Dim col&: col = FindColumnByHeader(ws, colHeader)
    If col <= 0 Then Exit Sub

    ' 3행~끝까지 적용
    With ws.Range(ws.Cells(3, col), ws.Cells(ws.rows.count, col))
        .Validation.Delete
        .Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Formula1:="=" & cacheName
    End With
End Sub

' 라벨에서 ID 찾기(캐시 시트)
Private Function FindIdByLabel(ByVal cacheName As String, ByVal label As String) As Variant
    On Error GoTo done
    Dim nm As name: Set nm = ThisWorkbook.Names(cacheName)
    Dim rng As Range: Set rng = nm.RefersToRange ' label 목록(B열)
    Dim c As Range
    For Each c In rng.Cells
        If CStr(c.Value) = label Then
            ' 같은 행의 A열이 value
            FindIdByLabel = c.Offset(0, -1).Value
            Exit Function
        End If
    Next c
done:
End Function

' 유틸: 헤더명/컬럼찾기/이름정리/JStr/Nz

Private Function HeaderAt(ws As Worksheet, ByVal col As Long) As String
    HeaderAt = CStr(ws.Cells(1, col).Value2)
End Function

Public Function FindColumnByHeader(ws As Worksheet, ByVal header As String) As Long
    Dim lastCol&, c&
    lastCol = XQL_LastCol(ws)
    For c = 1 To lastCol
        If LCase$(CStr(ws.Cells(1, c).Value2)) = LCase$(header) Then
            FindColumnByHeader = c: Exit Function
        End If
    Next
    FindColumnByHeader = 0
End Function

Private Function SanitizeName(ByVal s As String) As String
    Dim i&, ch$, out$: out = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z0-9_]" Then out = out & ch Else out = out & "_"
    Next
    If Len(out) = 0 Then out = "x"
    SanitizeName = out
End Function

Private Function JStrOrNull(ByVal s As String) As String
    s = Trim$(s)
    If Len(s) = 0 Then JStrOrNull = "null" Else JStrOrNull = """" & Replace(s, """", "\""") & """"
End Function

Private Function Nz(ByVal v As Variant, ByVal fb As Variant) As Variant
    If IsEmpty(v) Or v = "" Or v Is Nothing Then Nz = fb Else Nz = v
End Function

'Module XQL_IdUtils
' === Module: XQL_IdUtils ===
Option Explicit

' 빈 id 셀을 다음 증가 값으로 채움(중복 방지)
Public Sub XQL_AssignMissingIds()
    Dim ws As Worksheet: Set ws = ActiveSheet
    If UCase$(CStr(ws.Cells(1, 1).Value2)) <> "ID" Then
        MsgBox "A1 헤더가 'id'가 아닙니다.", vbExclamation: Exit Sub
    End If
    Dim lastRow&, r&, NextId&
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    NextId = NextId(ws)
    For r = 2 To lastRow
        If Len(ws.Cells(r, 1).Value2) = 0 Then
            ws.Cells(r, 1).Value = NextId
            NextId = NextId + 1
        End If
    Next
    MsgBox "빈 id가 자동 할당되었습니다.", vbInformation
End Sub

' 현재 시트에서 id의 최대값+1을 반환
Public Function NextId(ws As Worksheet) As Long
    Dim lastRow&, r&, m&
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    m = 0
    For r = 2 To lastRow
        If Len(ws.Cells(r, 1).Value2) > 0 Then
            If CLng(val(ws.Cells(r, 1).Value2)) > m Then m = CLng(val(ws.Cells(r, 1).Value2))
        End If
    Next
    NextId = m + 1
End Function

'Module XQL_Integrity
' === Module: XQL_IdUtils ===
Option Explicit

' 빈 id 셀을 다음 증가 값으로 채움(중복 방지)
Public Sub XQL_AssignMissingIds()
    Dim ws As Worksheet: Set ws = ActiveSheet
    If UCase$(CStr(ws.Cells(1, 1).Value2)) <> "ID" Then
        MsgBox "A1 헤더가 'id'가 아닙니다.", vbExclamation: Exit Sub
    End If
    Dim lastRow&, r&, NextId&
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    NextId = NextId(ws)
    For r = 2 To lastRow
        If Len(ws.Cells(r, 1).Value2) = 0 Then
            ws.Cells(r, 1).Value = NextId
            NextId = NextId + 1
        End If
    Next
    MsgBox "빈 id가 자동 할당되었습니다.", vbInformation
End Sub

' 현재 시트에서 id의 최대값+1을 반환
Public Function NextId(ws As Worksheet) As Long
    Dim lastRow&, r&, m&
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    m = 0
    For r = 2 To lastRow
        If Len(ws.Cells(r, 1).Value2) > 0 Then
            If CLng(val(ws.Cells(r, 1).Value2)) > m Then m = CLng(val(ws.Cells(r, 1).Value2))
        End If
    Next
    NextId = m + 1
End Function

'Module XQL_Integrity
' === Module: XQL_Integrity ===
Option Explicit

Private Const SHEET_CHECK As String = "XQLite_Check"
Private Const PULL_PAGE_SIZE As Long = 1000

' ───────────────── PUBLIC ENTRIES ─────────────────

' 1) 활성 시트 무결성 검사
Public Sub XQL_Check_RunActive()
    Dim ws As Worksheet: Set ws = ActiveSheet
    If Not EnsureDataSheet(ws) Then Exit Sub

    Dim table$: table = ws.name
    Dim serverDef As Object: Set serverDef = ServerFetchSchema(table)
    If serverDef Is Nothing Then
        MsgBox "서버에 스키마 정보(schema:" & table & ")가 없습니다.", vbExclamation
        Exit Sub
    End If

    Dim rows As Object: Set rows = ServerFetchAllRows(table, True) ' include_deleted=true
    RenderCheckReport ws, serverDef, rows
End Sub

' 2) Full Pull (서버 → 엑셀 덮어쓰기; 헤더/타입/데이터/보조열 세팅)
Public Sub XQL_Check_FullPullActive()
    On Error GoTo eh
    Dim ws As Worksheet: Set ws = ActiveSheet
    If ws.name = "XQLite" Then
        MsgBox "데이터 시트에서 실행하세요.", vbExclamation: Exit Sub
    End If

    Dim table$: table = ws.name
    Dim serverDef As Object: Set serverDef = ServerFetchSchema(table)
    If serverDef Is Nothing Then
        MsgBox "서버 스키마를 찾을 수 없습니다(schema:" & table & ").", vbExclamation: Exit Sub
    End If

    Application.ScreenUpdating = False

    ' 1) 헤더/타입 재작성
    WriteHeadersFromServer ws, serverDef

    ' 2) 데이터 클리어 후 풀
    ClearDataRows ws
    Dim rows As Object: Set rows = ServerFetchAllRows(table, True)
    WriteRowsToSheet ws, serverDef, rows

    ' 3) 보조열 세팅 + 최신 max_row_version 기록
    XQL_EnsureAuxCols ws
    Dim mx&: mx = FetchServerMaxRowVersion()
    Sheets("XQLite").Range("A7").Offset(0, 1).Value = mx

    Application.ScreenUpdating = True
    MsgBox "서버 → 엑셀 동기화(Full Pull) 완료", vbInformation
    Exit Sub
eh:
    Application.ScreenUpdating = True
    MsgBox "Full Pull 실패: " & Err.Description, vbExclamation
End Sub

' 3) Full Push (엑셀 → 서버 복구) - 기존 Recover 래핑
Public Sub XQL_Check_FullPushActive()
    XQL_RecoverServer
End Sub

' ──────────────── REPORT RENDERING ────────────────

Private Sub RenderCheckReport(ws As Worksheet, serverDef As Object, rows As Object)
    Dim rep As Worksheet: Set rep = EnsureReportSheet()

    rep.Cells.Clear
    rep.Range("A1").Value = "Integrity Report"
    rep.Range("A1").Font.Bold = True

    Dim r&: r = 3

    ' 0) 버전 비교
    Dim localMax&: localMax = CLng(Nz(Sheets("XQLite").Range("A7").Offset(0, 1).Value2, 0))
    Dim serverMax&: serverMax = FetchServerMaxRowVersion()
    rep.Cells(r, 1).Value = "Local LAST_MAX_ROW_VERSION": rep.Cells(r, 2).Value = localMax: r = r + 1
    rep.Cells(r, 1).Value = "Server max_row_version": rep.Cells(r, 2).Value = serverMax: r = r + 2

    ' 1) 스키마 비교
    rep.Cells(r, 1).Value = "Schema"
    rep.rows(r).Font.Bold = True: r = r + 1
    r = RenderSchemaDiff(ws, serverDef, rep, r)
    r = r + 1

    ' 2) 행 비교(개수/ID 차이/샘플 값)
    rep.Cells(r, 1).Value = "Rows"
    rep.rows(r).Font.Bold = True: r = r + 1
    r = RenderRowsDiff(ws, rows, rep, r)

    rep.Columns("A:G").AutoFit
    rep.Activate
    MsgBox "무결성 검사 완료. " & SHEET_CHECK & " 시트를 확인하세요.", vbInformation
End Sub

Private Function RenderSchemaDiff(ws As Worksheet, serverDef As Object, rep As Worksheet, ByVal r As Long) As Long
    Dim localCols As Object: Set localCols = CollectLocalCols(ws)
    Dim serverCols As Object: Set serverCols = CollectServerCols(serverDef)

    ' 누락/추가/타입불일치
    rep.Cells(r, 1).Value = "Missing in Excel": rep.rows(r).Font.Bold = True: r = r + 1
    Dim k As Variant, any As Boolean: any = False
    For Each k In serverCols.keys
        If Not localCols.Exists(k) Then
            rep.Cells(r, 1).Value = CStr(k) & " : " & serverCols(k)
            r = r + 1: any = True
        End If
    Next k
    If Not any Then rep.Cells(r, 1).Value = "(none)": r = r + 1

    rep.Cells(r, 1).Value = "Extra in Excel": rep.rows(r).Font.Bold = True: r = r + 1
    any = False
    For Each k In localCols.keys
        If Not serverCols.Exists(k) And k <> "id" And k <> "_ver" And k <> "_dirty" Then
            rep.Cells(r, 1).Value = CStr(k) & " : " & localCols(k)
            r = r + 1: any = True
        End If
    Next k
    If Not any Then rep.Cells(r, 1).Value = "(none)": r = r + 1

    rep.Cells(r, 1).Value = "Type mismatch": rep.rows(r).Font.Bold = True: r = r + 1
    any = False
    For Each k In serverCols.keys
        If localCols.Exists(k) Then
            If UCase$(CStr(localCols(k))) <> UCase$(CStr(serverCols(k))) Then
                rep.Cells(r, 1).Value = CStr(k) & " : Excel=" & CStr(localCols(k)) & " / Server=" & CStr(serverCols(k))
                r = r + 1: any = True
            End If
        End If
    Next k
    If Not any Then rep.Cells(r, 1).Value = "(none)": r = r + 1

    RenderSchemaDiff = r
End Function

Private Function RenderRowsDiff(ws As Worksheet, rows As Object, rep As Worksheet, ByVal r As Long) As Long
    Dim serverCount&, excelCount&, i&, lastRow&, idsExcel As Object, idsServer As Object
    serverCount = IIf(rows Is Nothing, 0, rows.count)
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    excelCount = IIf(lastRow >= 3, lastRow - 2, 0)

    rep.Cells(r, 1).Value = "Excel rows": rep.Cells(r, 2).Value = excelCount: r = r + 1
    rep.Cells(r, 1).Value = "Server rows (incl. deleted)": rep.Cells(r, 2).Value = serverCount: r = r + 2

    Set idsExcel = CollectExcelIds(ws)
    Set idsServer = CollectServerIds(rows)

    ' 누락/여분 ID (상위 100개만)
    rep.Cells(r, 1).Value = "IDs missing in Excel (sample)": rep.rows(r).Font.Bold = True: r = r + 1
    r = RenderIdDiff(rep, r, idsServer, idsExcel)

    rep.Cells(r, 1).Value = "IDs extra in Excel (sample)": rep.rows(r).Font.Bold = True: r = r + 1
    r = RenderIdDiff(rep, r, idsExcel, idsServer)

    ' 샘플 값 비교(최대 50행, 10컬럼)
    rep.Cells(r, 1).Value = "Cell mismatch (sample)": rep.rows(r).Font.Bold = True: r = r + 1
    r = RenderCellMismatchesSample(ws, rows, rep, r)

    RenderRowsDiff = r
End Function

Private Function RenderIdDiff(rep As Worksheet, ByVal r As Long, superset As Object, subset As Object) As Long
    Dim shown&, k As Variant: shown = 0
    For Each k In superset.keys
        If Not subset.Exists(k) Then
            rep.Cells(r, 1).Value = CLng(k): r = r + 1: shown = shown + 1
            If shown >= 100 Then Exit For
        End If
    Next k
    If shown = 0 Then rep.Cells(r, 1).Value = "(none)": r = r + 1
    RenderIdDiff = r
End Function

Private Function RenderCellMismatchesSample(ws As Worksheet, rows As Object, rep As Worksheet, ByVal r As Long) As Long
    Dim localCols As Object: Set localCols = CollectLocalCols(ws)
    Dim count&, i&, shown&: shown = 0
    If rows Is Nothing Then rep.Cells(r, 1).Value = "(none)": r = r + 1: RenderCellMismatchesSample = r: Exit Function

    For i = 1 To rows.count
        Dim row As Object: Set row = rows(i)
        Dim id&: id = CLng(Nz(row("id"), 0))
        If id = 0 Then GoTo conti

        Dim rr&: rr = FindRowById(ws, id)
        If rr = 0 Then GoTo conti

        Dim c&, mism As Boolean: mism = False
        For c = 1 To Application.WorksheetFunction.Min(localCols.count, 10)
            Dim k$: k = GetKeyByIndex(localCols, c)
            If k = "id" Or k = "_ver" Or k = "_dirty" Then GoTo contc

            Dim sv: If row.Exists(k) Then sv = row(k) Else sv = Empty
            Dim lv: lv = ws.Cells(rr, FindColumnByHeader(ws, k)).Value

            If Not EqLoose(lv, sv) Then
                rep.Cells(r, 1).Value = "id=" & id & ", col=" & k
                rep.Cells(r, 2).Value = "Excel=" & CStr(lv)
                rep.Cells(r, 3).Value = "Server=" & CStr(sv)
                r = r + 1: shown = shown + 1
                If shown >= 50 Then RenderCellMismatchesSample = r: Exit Function
            End If
contc:
        Next c
conti:
    Next i
    If shown = 0 Then rep.Cells(r, 1).Value = "(none)": r = r + 1
    RenderCellMismatchesSample = r
End Function

' ──────────────── SERVER HELPERS ────────────────

Private Function ServerFetchSchema(ByVal table As String) As Object
    On Error GoTo bad
    Dim wr$, q$, resp$, parsed As Object, rows As Object
    wr = "key='schema:" & EscapeSqlStr(table) & "'"
    q = "query{ rows(table:""meta"", whereRaw:""" & wr & """, limit:1){ rows } }"
    resp = XQL_GraphQL(q, "{}")
    Set parsed = JsonConverter.ParseJson(resp)
    Set rows = parsed("data")("rows")("rows")
    If rows Is Nothing Or rows.count = 0 Then GoTo bad

    Dim js$: js = CStr(rows(1)("value"))
    Set ServerFetchSchema = JsonConverter.ParseJson(js)
    Exit Function
bad:
    Set ServerFetchSchema = Nothing
End Function

Private Function ServerFetchAllRows(ByVal table As String, ByVal includeDeleted As Boolean) As Object
    Dim out As Object: Set out = CreateObject("Scripting.Dictionary") ' id -> row
    Dim off&, got&, q$, vars$, resp$, parsed As Object, rs As Object
    Do
        q = "query($t:String!,$o:Int,$l:Int,$inc:Boolean){rows(table:$t,offset:$o,limit:$l,include_deleted:$inc){rows}}"
        vars = "{""t"":""" & table & """,""o"":" & CStr(off) & ",""l"":" & CStr(PULL_PAGE_SIZE) & ",""inc"":" & LCase$(CStr(includeDeleted)) & "}"
        resp = XQL_GraphQL(q, vars)
        Set parsed = JsonConverter.ParseJson(resp)
        Set rs = parsed("data")("rows")("rows")
        If rs Is Nothing Then Exit Do
        Dim i&
        For i = 1 To rs.count
            out(CLng(rs(i)("id"))) = rs(i)
        Next i
        got = rs.count
        off = off + got
    Loop While got = PULL_PAGE_SIZE
    Set ServerFetchAllRows = outToArrayList(out)
End Function

Private Function FetchServerMaxRowVersion() As Long
    On Error GoTo bad
    Dim q$, resp$, parsed As Object
    q = "query{ meta{ max_row_version } }"
    resp = XQL_GraphQL(q, "{}")
    Set parsed = JsonConverter.ParseJson(resp)
    FetchServerMaxRowVersion = CLng(Nz(parsed("data")("meta")("max_row_version"), 0))
    Exit Function
bad:
    FetchServerMaxRowVersion = 0
End Function

' ──────────────── SHEET I/O ────────────────

Private Sub WriteHeadersFromServer(ws As Worksheet, serverDef As Object)
    Dim i&, c&
    ws.Cells.Clear
    ' A1 id, A2 INTEGER
    ws.Cells(1, 1).Value = "id"
    ws.Cells(2, 1).Value = "INTEGER"
    c = 2
    For i = 0 To serverDef("columns").count - 1
        Dim nm$, ty$
        nm = CStr(serverDef("columns")(i)("name"))
        If LCase$(nm) = "id" Then GoTo cont
        ty = MapSqliteTypeToToken(CStr(serverDef("columns")(i)("type")))
        ws.Cells(1, c).Value = nm
        ws.Cells(2, c).Value = ty
        c = c + 1
cont:
    Next i
    ' 디자인/고정
    ws.rows(1).Font.Bold = True
    ws.rows(2).Font.Italic = True
    ws.rows(2).Interior.Color = RGB(245, 245, 245)
    ws.Columns.AutoFit
    ws.Activate
    ActiveWindow.SplitRow = 2
    ActiveWindow.FreezePanes = True
End Sub

Private Sub WriteRowsToSheet(ws As Worksheet, serverDef As Object, rows As Object)
    If rows Is Nothing Then Exit Sub
    Dim i&, r&, lc&, vcol&, dcol&
    lc = XQL_LastCol(ws)
    vcol = lc + 1: dcol = lc + 2
    XQL_EnsureAuxCols ws

    r = 2
    Dim keys As Object: Set keys = CollectServerCols(serverDef)
    Dim k As Variant

    For i = 0 To rows.count - 1
        Dim row As Object: Set row = rows(i)
        r = r + 1
        ws.Cells(r, 1).Value = Nz(row("id"), "")
        For Each k In keys.keys
            If k <> "id" Then
                Dim col&: col = FindColumnByHeader(ws, CStr(k))
                If col > 0 Then
                    SetCellValue ws.Cells(r, col), IIf(row.Exists(k), row(k), Empty)
                End If
            End If
        Next k
        ' 메타
        On Error Resume Next
        ws.Cells(r, vcol).Value = CLng(Nz(row("row_version"), 0))
        If CBool(Nz(row("deleted"), False)) Then
            ws.rows(r).Font.Strikethrough = True
            ws.rows(r).Interior.Color = RGB(240, 240, 240)
        End If
        On Error GoTo 0
    Next i
    ws.Columns.AutoFit
End Sub

Private Sub ClearDataRows(ws As Worksheet)
    Dim lastRow&: lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If lastRow >= 3 Then ws.Range("A3:" & ws.Cells(lastRow, ws.Columns.count).Address).Clear
End Sub

' ──────────────── COLLECT/COMPARE HELPERS ────────────────

Private Function CollectLocalCols(ws As Worksheet) As Object
    Dim lastCol&, c&, dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    lastCol = XQL_LastCol(ws)
    For c = 1 To lastCol
        Dim name$, ty$: name = LCase$(CStr(ws.Cells(1, c).Value2))
        If Len(name) = 0 Then GoTo cont
        ty = UCase$(CStr(ws.Cells(2, c).Value2))
        If Len(ty) = 0 Then ty = "TEXT"
        dict(name) = ty
cont:
    Next c
    Set CollectLocalCols = dict
End Function

Private Function CollectServerCols(serverDef As Object) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim i&
    For i = 0 To serverDef("columns").count - 1
        Dim nm$, ty$
        nm = LCase$(CStr(serverDef("columns")(i)("name")))
        ty = MapSqliteTypeToToken(CStr(serverDef("columns")(i)("type")))
        dict(nm) = ty
    Next i
    Set CollectServerCols = dict
End Function

Private Function CollectExcelIds(ws As Worksheet) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim lastRow&, r&
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = 3 To lastRow
        Dim v&: v = CLng(val(ws.Cells(r, 1).Value2))
        If v > 0 Then dict(CStr(v)) = 1
    Next
    Set CollectExcelIds = dict
End Function

Private Function CollectServerIds(rows As Object) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim i&
    If rows Is Nothing Then Set CollectServerIds = dict: Exit Function
    For i = 0 To rows.count - 1
        dict(CStr(CLng(Nz(rows(i)("id"), 0)))) = 1
    Next i
    Set CollectServerIds = dict
End Function

Private Function GetKeyByIndex(dict As Object, ByVal idx As Long) As String
    Dim i&, k As Variant: i = 0
    For Each k In dict.keys
        i = i + 1
        If i = idx Then GetKeyByIndex = CStr(k): Exit Function
    Next k
    GetKeyByIndex = ""
End Function

' ──────────────── SMALL UTILS ────────────────

Private Function EnsureReportSheet() As Worksheet
    On Error Resume Next
    Set EnsureReportSheet = Sheets(SHEET_CHECK)
    On Error GoTo 0
    If EnsureReportSheet Is Nothing Then
        Set EnsureReportSheet = Sheets.add(After:=Sheets(Sheets.count))
        EnsureReportSheet.name = SHEET_CHECK
    End If
End Function

Private Function EnsureDataSheet(ws As Worksheet) As Boolean
    EnsureDataSheet = Not (ws.name = "XQLite" Or ws.name = "XQLite_Conflicts" Or ws.name = "XQLite_Presence" Or ws.name = "XQLite_Query" Or ws.name = "XQLite_Audit" Or ws.name = "XQLite_EnumCache" Or ws.name = "XQLite_Enums")
    If Not EnsureDataSheet Then MsgBox "데이터 시트에서 실행하세요.", vbExclamation
End Function

Private Function MapSqliteTypeToToken(ByVal t As String) As String
    Dim u$: u = UCase$(Trim$(t))
    If InStr(u, "INT") > 0 Then MapSqliteTypeToToken = "INTEGER": Exit Function
    If u = "REAL" Or u = "FLOAT" Or u = "DOUBLE" Or InStr(u, "NUMERIC") > 0 Then MapSqliteTypeToToken = "REAL": Exit Function
    If u = "BOOLEAN" Or InStr(u, "BOOL") > 0 Then MapSqliteTypeToToken = "BOOLEAN": Exit Function
    If u = "BLOB" Then MapSqliteTypeToToken = "BLOB": Exit Function
    MapSqliteTypeToToken = "TEXT"
End Function

Private Function EqLoose(a As Variant, b As Variant) As Boolean
    If IsEmpty(a) And (IsEmpty(b) Or b = "" Or b Is Nothing) Then EqLoose = True: Exit Function
    If IsEmpty(b) And (IsEmpty(a) Or a = "" Or a Is Nothing) Then EqLoose = True: Exit Function
    On Error Resume Next
    EqLoose = (CStr(a) = CStr(b))
End Function

Private Sub SetCellValue(ByVal cell As Range, ByVal v As Variant)
    If IsObject(v) Then
        cell.Value = CStr(v)
    Else
        cell.Value = v
    End If
End Sub

Private Function outToArrayList(dict As Object) As Object
    Dim arr As Object: Set arr = CreateObject("System.Collections.ArrayList")
    Dim k As Variant
    For Each k In dict.keys
        arr.add dict(k)
    Next k
    Set outToArrayList = arr
End Function

Private Function EscapeSqlStr(ByVal s As String) As String
    EscapeSqlStr = Replace(s, "'", "''")
End Function

Private Function Nz(ByVal v As Variant, ByVal fb As Variant) As Variant
    If IsEmpty(v) Or v = "" Or v Is Nothing Then Nz = fb Else Nz = v
End Function

Private Function FindRowById(ws As Worksheet, ByVal id As Long) As Long
    Dim lastRow&, r&
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = 3 To lastRow
        If CLng(val(ws.Cells(r, 1).Value2)) = id Then FindRowById = r: Exit Function
    Next
    FindRowById = 0
End Function

'Module XQL_Navigator
' === Module: XQL_Navigator ===
Option Explicit

Public Sub XQL_GoNextDirty()
    GoNextByColor OrByDirtyCol:=True, rgbCrit:=0
End Sub

Public Sub XQL_GoNextConflict()
    GoNextByColor OrByDirtyCol:=False, rgbCrit:=RGB(255, 255, 180) ' 노랑(충돌)
End Sub

Public Sub XQL_GoNextLockedByOthers()
    GoNextByColor OrByDirtyCol:=False, rgbCrit:=RGB(255, 230, 230) ' 연분홍(타인락)
End Sub

Private Sub GoNextByColor(ByVal OrByDirtyCol As Boolean, ByVal rgbCrit As Long)
    Dim ws As Worksheet: Set ws = ActiveSheet
    If ws.name = "XQLite" Then Exit Sub

    Dim startR&: startR = Selection.row
    Dim lastRow&, lastCol&: lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row: lastCol = XQL_LastCol(ws)
    Dim r&, c&, vcol&, dcol&: vcol = XQL_VerCol(ws): dcol = XQL_DirtyCol(ws)

    For r = IIf(startR < 3, 3, startR) To lastRow
        If OrByDirtyCol Then
            If dcol > 0 And Len(ws.Cells(r, dcol).Value2) > 0 Then
                ws.Cells(r, 1).Select: Exit Sub
            End If
        Else
            For c = 1 To lastCol
                If ws.Cells(r, c).Interior.Color = rgbCrit Then
                    ws.Cells(r, c).Select: Exit Sub
                End If
            Next c
        End If
    Next r
    MsgBox "다음 항목이 없습니다.", vbInformation
End Sub

'Module XQL_Outbox
' === Module: XQL_Outbox ===
Option Explicit

Private Const SHEET_OB As String = "XQLite_Outbox"
' 컬럼: A=id(자동), B=ts, C=table, D=sheet, E=row_idx, F=payload_json, G=tries, H=next_at, I=last_error

Private mTickAt As Date
Private Const BATCH_MAX As Long = 100        ' 한번에 보낼 최대 항목
Private Const JITTER_MS As Long = 250        ' 지터 최대치

' ────────────── 엔트리 ──────────────

Public Sub XQL_Outbox_Setup()
    EnsureOutboxSheet
    MsgBox "Outbox 시트를 준비했습니다. (숨김)", vbInformation
End Sub

Public Sub XQL_Outbox_Start()
    EnsureOutboxSheet
    ScheduleTick True
End Sub

Public Sub XQL_Outbox_Stop()
    On Error Resume Next
    If mTickAt <> 0 Then Application.OnTime mTickAt, "XQL_Outbox_Tick", , False
    mTickAt = 0
End Sub

Public Sub XQL_Outbox_Open()
    Dim ws As Worksheet: Set ws = EnsureOutboxSheet
    ws.Visible = xlSheetVisible
    ws.Activate
End Sub

Public Sub XQL_Outbox_Hide()
    Dim ws As Worksheet: Set ws = EnsureOutboxSheet
    ws.Visible = xlSheetVeryHidden
End Sub

Public Sub XQL_Outbox_RetryNow()
    Dim ws As Worksheet: Set ws = EnsureOutboxSheet
    Dim last&, r&
    last = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If last < 2 Then Exit Sub
    For r = 2 To last
        If Len(ws.Cells(r, 1).Value2) > 0 Then
            ws.Cells(r, 8).Value = Now   ' next_at = now
        End If
    Next
    ScheduleTick True
    MsgBox "모든 보류 항목을 즉시 재시도합니다.", vbInformation
End Sub

Public Sub XQL_Outbox_ClearFailed()
    Dim ws As Worksheet: Set ws = EnsureOutboxSheet
    Dim maxTry&: maxTry = XQL_OutboxMaxRetry()
    Dim last&, r&
    last = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = last To 2 Step -1
        If CLng(Nz(ws.Cells(r, 7).Value2, 0)) >= maxTry Then
            ws.rows(r).Delete
        End If
    Next
    MsgBox "최대 재시도 초과 항목을 정리했습니다.", vbInformation
End Sub

' **선택된 행을 즉시 큐잉** (Dirty Navigator나 실패시 수동 보강용)
Public Sub XQL_Outbox_EnqueueSelection()
    Dim ws As Worksheet: Set ws = ActiveSheet
    If Not IsDataSheet(ws) Then MsgBox "데이터 시트에서 실행하세요.", vbExclamation: Exit Sub

    Dim rows As Object: Set rows = CreateObject("System.Collections.ArrayList")
    Dim a As Range, r As Range
    For Each a In Selection.Areas
        For Each r In a.rows
            If r.row >= 3 Then rows.add r.row
        Next r
    Next a
    If rows.count = 0 Then Exit Sub

    Dim i&
    For i = 0 To rows.count - 1
        EnqueueUpsertRow ws, CLng(rows(i))
    Next i

    MsgBox rows.count & "건을 Outbox에 등록했습니다.", vbInformation
End Sub

' ────────────── 내부: 주기 처리 ──────────────

Public Sub XQL_Outbox_Tick()
    On Error GoTo done
    ProcessDueBatch
done:
    ScheduleTick False
End Sub

Private Sub ProcessDueBatch()
    Dim ob As Worksheet: Set ob = EnsureOutboxSheet

    ' Due rows 뽑기(상위 BATCH_MAX)
    Dim last&, r&, due As Object: Set due = CreateObject("System.Collections.ArrayList")
    last = ob.Cells(ob.rows.count, 1).End(xlUp).row
    If last < 2 Then Exit Sub

    For r = 2 To last
        If Len(ob.Cells(r, 1).Value2) > 0 Then
            If Nz(ob.Cells(r, 8).Value2, 0) = 0 Or ob.Cells(r, 8).Value <= Now Then
                due.add r
                If due.count >= BATCH_MAX Then Exit For
            End If
        End If
    Next r
    If due.count = 0 Then Exit Sub

    ' 테이블 단위로 묶어 보내기
    Dim groups As Object: Set groups = CreateObject("Scripting.Dictionary") ' table -> ArrayList(rows)
    For r = 0 To due.count - 1
        Dim row&: row = CLng(due(r))
        Dim t$: t = CStr(ob.Cells(row, 3).Value2)
        Dim list As Object
        If Not groups.Exists(t) Then
            Set list = CreateObject("System.Collections.ArrayList")
            groups.add t, list
        Else
            Set list = groups(t)
        End If
        list.add row
    Next r

    Dim k As Variant
    For Each k In groups.keys
        SendGroup ob, CStr(k), groups(k)
    Next k
End Sub

Private Sub SendGroup(ob As Worksheet, ByVal table As String, rows As Object)
    If rows.count = 0 Then Exit Sub

    ' rows -> payload 묶기
    Dim actor$: actor = XQL_GetNickname()
    Dim arr$, i&, row&, pl$, joined$
    joined = "["
    For i = 0 To rows.count - 1
        row = CLng(rows(i))
        pl = CStr(ob.Cells(row, 6).Value2)   ' payload_json
        If Len(pl) = 0 Then GoTo cont
        If Right$(joined, 1) <> "[" Then joined = joined & ","
        joined = joined & pl
cont:
    Next i
    joined = joined & "]"

    Dim q$, vars$, resp$
    q = "mutation($t:String!,$rows:[UpsertRowInput!]!,$a:String!){upsertRows(table:$t,rows:$rows,actor:$a){max_row_version affected errors conflicts}}"
    vars = "{""t"":""" & table & """,""rows"":" & joined & ",""a"":""" & actor & """}"

    resp = XQL_GraphQL(q, vars)

    If XQL_HasErrors(resp) Then
        MarkFailed ob, rows, resp
        Exit Sub
    End If

    ' 에러 구조가 없으면 성공 처리
    Dim mx&: mx = XQL_ExtractMaxRowVersion(resp)
    Dim i&
    For i = rows.count - 1 To 0 Step -1
        Dim ro&: ro = CLng(rows(i))
        ' 성공: 행 삭제 + 시트 반영
        Dim sheetName$: sheetName = CStr(ob.Cells(ro, 4).Value2)
        Dim rr&: rr = CLng(Nz(ob.Cells(ro, 5).Value2, 0))
        If mx > 0 Then
            On Error Resume Next
            Dim ws As Worksheet: Set ws = Sheets(sheetName)
            If Not ws Is Nothing Then
                Dim vcol&: vcol = XQL_VerCol(ws)
                If rr > 0 And vcol > 0 Then
                    ws.Cells(rr, vcol).Value = mx
                    ws.Cells(rr, XQL_DirtyCol(ws)).Value = Empty
                End If
            End If
            On Error GoTo 0
        End If
        ob.rows(ro).Delete
    Next i
End Sub

Private Sub MarkFailed(ob As Worksheet, rows As Object, ByVal errMsg As String)
    Dim i&, r&
    For i = 0 To rows.count - 1
        r = CLng(rows(i))
        Dim tries&: tries = CLng(Nz(ob.Cells(r, 7).Value2, 0)) + 1
        ob.Cells(r, 7).Value = tries
        ob.Cells(r, 9).Value = Left$(errMsg, 250)
        Dim BackoffSec As Double: BackoffSec = BackoffSec(tries)
        ob.Cells(r, 8).Value = Now + (BackoffSec / 86400#) ' next_at
        ' 시각적 표시(선택): 아무 것도 안 함. Outbox 시트에서만 관리
    Next i
End Sub

' ────────────── 큐잉 헬퍼 ──────────────

' Dirty 행을 안전하게 큐에 추가(한 행 단위)
Public Sub EnqueueUpsertRow(ws As Worksheet, ByVal rowIdx As Long)
    On Error GoTo eh
    If Not IsDataSheet(ws) Then Exit Sub
    If rowIdx < 3 Then Exit Sub

    Dim payload$: payload = BuildUpsertPayloadForRow(ws, rowIdx)
    If Len(payload) = 0 Then Exit Sub

    Dim ob As Worksheet: Set ob = EnsureOutboxSheet
    Dim id&: id = ob.Cells(ob.rows.count, 1).End(xlUp).row + 1

    ob.Cells(id, 1).Value = id - 1 ' 단순 시퀀스
    ob.Cells(id, 2).Value = Now
    ob.Cells(id, 3).Value = ws.name
    ob.Cells(id, 4).Value = ws.name
    ob.Cells(id, 5).Value = rowIdx
    ob.Cells(id, 6).Value = payload
    ob.Cells(id, 7).Value = 0
    ob.Cells(id, 8).Value = Now
    ob.Cells(id, 9).Value = ""

    ' 더티 마킹은 유지(성공 시 Tick에서 클리어)
    Exit Sub
eh:
    ' 무시
End Sub

' 한 행의 UpsertRowInput JSON 생성
Private Function BuildUpsertPayloadForRow(ws As Worksheet, ByVal r As Long) As String
    Dim vcol&, dcol&, lc&, c&, k$, v
    vcol = XQL_VerCol(ws): dcol = XQL_DirtyCol(ws): lc = XQL_LastCol(ws)
    If vcol = 0 Or dcol = 0 Then XQL_EnsureAuxCols ws: vcol = XQL_VerCol(ws): dcol = XQL_DirtyCol(ws)
    If CLng(Nz(ws.Cells(r, 1).Value2, 0)) = 0 Then BuildUpsertPayloadForRow = "": Exit Function

    Dim basev&: basev = CLng(Nz(ws.Cells(r, vcol).Value2, 0))
    Dim data$, first As Boolean: first = True
    data = "{"
    For c = 2 To lc
        k = CStr(ws.Cells(1, c).Value2)
        If Len(k) = 0 Then GoTo cont
        v = ws.Cells(r, c).Value
        If Not first Then data = data & ","
        data = data & """" & XQL_EscapeJson(k) & """:"
        If IsEmpty(v) Or v = "" Then
            data = data & "null"
        ElseIf IsNumeric(v) And Not IsDate(v) Then
            data = data & CStr(v)
        ElseIf VarType(v) = vbBoolean Then
            data = data & IIf(CBool(v), "true", "false")
        Else
            data = data & """" & XQL_EscapeJson(CStr(v)) & """"
        End If
        first = False
cont:
    Next c
    data = data & "}"

    BuildUpsertPayloadForRow = "{""id"":" & CLng(ws.Cells(r, 1).Value2) & ",""base_row_version"":" & basev & ",""data"":" & data & "}"
End Function

' ────────────── 유틸 ──────────────

Private Function EnsureOutboxSheet() As Worksheet
    On Error Resume Next
    Set EnsureOutboxSheet = Sheets(SHEET_OB)
    On Error GoTo 0
    If EnsureOutboxSheet Is Nothing Then
        Set EnsureOutboxSheet = Sheets.add(After:=Sheets(Sheets.count))
        EnsureOutboxSheet.name = SHEET_OB
        With EnsureOutboxSheet
            .Range("A1:I1").Value = Array("id", "ts", "table", "sheet", "row_idx", "payload_json", "tries", "next_at", "last_error")
            .rows(1).Font.Bold = True
            .Columns("A:I").ColumnWidth = 18
            .Visible = xlSheetVeryHidden
        End With
    End If
End Function

Private Sub ScheduleTick(ByVal immediate As Boolean)
    On Error Resume Next
    If mTickAt <> 0 Then Application.OnTime mTickAt, "XQL_Outbox_Tick", , False
    Dim sec As Double: sec = IIf(immediate, 0.5, XQL_OutboxRetrySec())
    mTickAt = Now + (sec / 86400#)
    Application.OnTime mTickAt, "XQL_Outbox_Tick"
End Sub

Private Function BackoffSec(ByVal tries As Long) As Double
    Dim base As Double: base = XQL_OutboxRetrySec()
    Dim s As Double: s = base * (2 ^ Application.WorksheetFunction.Min(tries - 1, 6)) ' 최대 64배
    Dim jitter As Double: Randomize: jitter = (Rnd() * (JITTER_MS / 1000#))
    BackoffSec = s + jitter
    If BackoffSec > 300# Then BackoffSec = 300# ' 상한 5분
End Function

Private Function IsDataSheet(ws As Worksheet) As Boolean
    Dim n$: n = ws.name
    IsDataSheet = Not (n = "XQLite" Or n = "XQLite_Conflicts" Or n = "XQLite_Presence" Or n = "XQLite_EnumCache" Or n = "XQLite_Enums" Or n = "XQLite_Check" Or n = "XQLite_Query" Or n = "XQLite_Audit" Or n = "XQLite_Outbox")
End Function

Private Function Nz(ByVal v As Variant, ByVal fb As Variant) As Variant
    If IsEmpty(v) Or v = "" Or v Is Nothing Then Nz = fb Else Nz = v
End Function

'Module XQL_Permissions
' === Module: XQL_Permissions ===
Option Explicit

' 규칙 시트
Private Const SHEET_PERM As String = "XQLite_Perms"

' ─────────────────────────────────────────────
' 엔트리

' 규칙 시트 생성
Public Sub XQL_Perms_Setup()
    Dim ws As Worksheet
    On Error Resume Next: Set ws = Sheets(SHEET_PERM): On Error GoTo 0
    If ws Is Nothing Then
        Set ws = Sheets.add(After:=Sheets(Sheets.count))
        ws.name = SHEET_PERM
    Else
        ws.Cells.Clear
    End If

    With ws
        .Range("A1:F1").Value = Array( _
            "target_sheet", _               ' 시트명 또는 * (모든 데이터 시트)
            "selector", _                   ' ALL | COLUMN:<헤더> | RANGE:<A1표기>
            "allow_nicknames", _            ' 편집 허용 닉네임 CSV (예: alice,bob,*)
            "hide", _                       ' TRUE면 숨김 (COLUMN/RANGE 대상만)
            "note", _                       ' 메모
            "enabled" _                     ' TRUE/FALSE
        )
        .rows(1).Font.Bold = True
        .Columns("A:F").ColumnWidth = 28

        ' 예시 규칙
        .Range("A2:F2").Value = Array("items", "COLUMN:rarity", "planner_1,planner_2", "FALSE", "레어리티는 기획만 편집", "TRUE")
        .Range("A3:F3").Value = Array("items", "COLUMN:atk_lv50", "balancer_1,balancer_2", "FALSE", "밸런스팀 전용", "TRUE")
        .Range("A4:F4").Value = Array("items", "COLUMN:secret_flag", "", "TRUE", "민감 컬럼 숨김", "TRUE")
        .Range("A5:F5").Value = Array("*", "ALL", "admin", "FALSE", "관리자만 편집, 나머지는 읽기전용", "FALSE")
    End With

    MsgBox SHEET_PERM & " 시트를 만들었습니다." & vbCrLf & _
           "- target_sheet: 시트명 또는 *" & vbCrLf & _
           "- selector: ALL | COLUMN:<헤더> | RANGE:<주소>" & vbCrLf & _
           "- allow_nicknames: 허용 닉네임 CSV, *는 모두 허용" & vbCrLf & _
           "- hide: TRUE면 숨김 적용(선택)" & vbCrLf & _
           "- enabled: TRUE일 때만 적용", vbInformation
End Sub

' 활성 시트에 적용
Public Sub XQL_Perms_ApplyActive()
    ApplyForSheet ActiveSheet
    MsgBox "권한 프리셋을 적용했습니다: " & ActiveSheet.name, vbInformation
End Sub

' 모든 데이터 시트에 적용
Public Sub XQL_Perms_ApplyAll()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsDataSheet(ws) Then ApplyForSheet ws
    Next ws
    MsgBox "모든 데이터 시트에 권한 프리셋을 적용했습니다.", vbInformation
End Sub

' ─────────────────────────────────────────────
' 내부 본체

Private Sub ApplyForSheet(ws As Worksheet)
    If Not IsDataSheet(ws) Then Exit Sub

    Dim rules As Object: Set rules = LoadRulesFor(ws.name)
    Dim pass$: pass = XQL_ProtectPassword()
    Dim lastRow&, lastCol&: lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row: If lastRow < 3 Then lastRow = 3
    lastCol = XQL_LastCol(ws): If lastCol < 1 Then lastCol = 1

    ' 보호 해제 후 기본 잠금=TRUE, 선택=Unlocked만
    On Error Resume Next
    ws.Unprotect Password:=pass
    ws.Cells.Locked = True
    ws.EnableSelection = xlUnlockedCells
    On Error GoTo 0

    ' 숨김/언숨김 원복(일단 모두 언숨김 후 규칙대로 숨김)
    Dim c&: For c = 1 To 200
        On Error Resume Next
        ws.Columns(c).Hidden = False
        On Error GoTo 0
    Next c

    ' 허용 범위(=Unlock) 표집합
    Dim allowRanges As Collection: Set allowRanges = New Collection

    ' 규칙 적용
    Dim i&, r As Object
    For i = 1 To rules.count
        Set r = rules(i)
        Dim allowMe As Boolean: allowMe = IsAllowedForCurrent(r("allow"))
        Dim selType$, selVal$: selType = r("stype"): selVal = r("sval")

        If selType = "ALL" Then
            If allowMe Or r("allow") = "*" Then
                allowRanges.add ws.Range(ws.Cells(3, 1), ws.Cells(lastRow, lastCol))
            End If
        ElseIf selType = "COLUMN" Then
            Dim col&: col = FindColumnByHeader(ws, selVal)
            If col > 0 Then
                If r("hide") Then
                    On Error Resume Next: ws.Columns(col).Hidden = True: On Error GoTo 0
                End If
                If allowMe Or r("allow") = "*" Then
                    allowRanges.add ws.Range(ws.Cells(3, col), ws.Cells(lastRow, col))
                End If
            End If
        ElseIf selType = "RANGE" Then
            On Error Resume Next
            Dim rg As Range: Set rg = Nothing
            Set rg = ws.Range(selVal)
            On Error GoTo 0
            If Not rg Is Nothing Then
                If r("hide") Then
                    On Error Resume Next: rg.EntireColumn.Hidden = True: On Error GoTo 0
                End If
                If allowMe Or r("allow") = "*" Then
                    ' 데이터 행만 교집합
                    Dim rr As Range
                    Set rr = Intersect(rg, ws.Range(ws.Cells(3, 1), ws.Cells(lastRow, lastCol)))
                    If Not rr Is Nothing Then allowRanges.add rr
                End If
            End If
        End If
    Next i

    ' 허용 범위를 Unlock
    Dim ar As Variant
    For Each ar In allowRanges
        On Error Resume Next
        ar.Locked = False
        On Error GoTo 0
    Next ar

    ' 최종 Protect
    On Error Resume Next
    ws.Protect Password:=pass, UserInterfaceOnly:=True, AllowFormattingCells:=False, AllowSorting:=True, AllowFiltering:=True
    ws.EnableSelection = xlUnlockedCells
    On Error GoTo 0
End Sub

' 규칙 로드
Private Function LoadRulesFor(ByVal sheetName As String) As Collection
    Dim cfg As Worksheet
    On Error Resume Next: Set cfg = Sheets(SHEET_PERM): On Error GoTo 0
    Dim col As New Collection: Set LoadRulesFor = col
    If cfg Is Nothing Then Exit Function

    Dim last&, r&
    last = cfg.Cells(cfg.rows.count, 1).End(xlUp).row
    For r = 2 To last
        If UCase$(Trim$(CStr(cfg.Cells(r, 6).Value2))) <> "TRUE" Then GoTo cont
        Dim tgt$, sel$, allow$, hideB As Boolean
        tgt = Trim$(CStr(cfg.Cells(r, 1).Value2))
        sel = Trim$(CStr(cfg.Cells(r, 2).Value2))
        allow = Trim$(CStr(cfg.Cells(r, 3).Value2))
        hideB = (UCase$(CStr(cfg.Cells(r, 4).Value2)) = "TRUE")

        If Len(tgt) = 0 Or Len(sel) = 0 Then GoTo cont
        If Not (tgt = "*" Or LCase$(tgt) = LCase$(sheetName)) Then GoTo cont

        Dim st$, sv$
        If InStr(1, UCase$(sel), "ALL", vbTextCompare) = 1 Then
            st = "ALL": sv = ""
        ElseIf InStr(1, UCase$(sel), "COLUMN:", vbTextCompare) = 1 Then
            st = "COLUMN": sv = Mid$(sel, Len("COLUMN:") + 1)
        ElseIf InStr(1, UCase$(sel), "RANGE:", vbTextCompare) = 1 Then
            st = "RANGE": sv = Mid$(sel, Len("RANGE:") + 1)
        Else
            GoTo cont
        End If

        Dim o As Object: Set o = CreateObject("Scripting.Dictionary")
        o("stype") = st: o("sval") = sv: o("allow") = allow: o("hide") = hideB
        col.add o
cont:
    Next r
End Function

' 현재 닉네임 허용 여부
Private Function IsAllowedForCurrent(ByVal allowCsv As String) As Boolean
    Dim me$: me = LCase$(XQL_GetNickname())
    Dim s$: s = Trim$(LCase$(allowCsv))
    If Len(s) = 0 Then IsAllowedForCurrent = False: Exit Function
    If s = "*" Then IsAllowedForCurrent = True: Exit Function
    Dim parts() As String, i&
    parts = Split(s, ",")
    For i = LBound(parts) To UBound(parts)
        If Trim$(parts(i)) = Me Then IsAllowedForCurrent = True: Exit Function
    Next i
    IsAllowedForCurrent = False
End Function

' 데이터 시트 판별
Private Function IsDataSheet(ws As Worksheet) As Boolean
    Dim n$: n = ws.name
    IsDataSheet = Not (n = "XQLite" Or n = "XQLite_Conflicts" Or n = "XQLite_Presence" Or n = "XQLite_EnumCache" _
                       Or n = "XQLite_Enums" Or n = "XQLite_Check" Or n = "XQLite_Query" Or n = "XQLite_Audit" _
                       Or n = "XQLite_Outbox" Or n = "XQLite_Resolve" Or n = "XQLite_Perms")
End Function

' 헤더로 컬럼 찾기
Private Function FindColumnByHeader(ws As Worksheet, ByVal header As String) As Long
    Dim lastCol&, c&
    lastCol = XQL_LastCol(ws)
    For c = 1 To lastCol
        If LCase$(CStr(ws.Cells(1, c).Value2)) = LCase$(header) Then
            FindColumnByHeader = c: Exit Function
        End If
    Next
    FindColumnByHeader = 0
End Function

'Module XQL_PresenceLock
Option Explicit

' 3초마다 호출(Workbook에서 예약)
Public Sub XQL_Heartbeat()
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim addr$: addr = Selection.Address(False, False)
    Dim q$, payload$, resp$
    q = "mutation($n:String!,$s:String,$c:String){presenceHeartbeat(nickname:$n,sheet:$s,cell:$c)}"
    payload = "{""n"":""" & XQL_GetNickname() & """,""s"":""" & ws.name & """,""c"":""" & addr & """}"
    resp = XQL_GraphQL(q, payload)
    ' 다음 타이머 예약
    ThisWorkbook.XQL_StartHeartbeat
End Sub

Public Sub XQL_AcquireLockSelection()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim addr$: addr = Selection.Address(False, False)
    Dim q$, payload$, resp$
    q = "mutation($s:String!,$c:String!,$n:String!){acquireLock(sheet:$s,cell:$c,nickname:$n)}"
    payload = "{""s"":""" & ws.name & """,""c"":""" & addr & """,""n"":""" & XQL_GetNickname() & """}"
    resp = XQL_GraphQL(q, payload)
    If InStr(1, resp, "true", vbTextCompare) > 0 Then
        Selection.Interior.Color = RGB(220, 240, 255)
        MsgBox "락 획득", vbInformation
    Else
        MsgBox "락 획득 실패(다른 사용자가 점유중)", vbExclamation
    End If
End Sub

Public Sub XQL_ReleaseLockSelection()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim addr$: addr = Selection.Address(False, False)
    Dim q$, payload$, resp$
    q = "mutation($s:String!,$c:String!,$n:String!){releaseLock(sheet:$s,cell:$c,nickname:$n)}"
    payload = "{""s"":""" & ws.name & """,""c"":""" & addr & """,""n"":""" & XQL_GetNickname() & """}"
    resp = XQL_GraphQL(q, payload)
    Selection.Interior.ColorIndex = xlNone
    MsgBox "락 해제 시도 완료", vbInformation
End Sub

'Module XQL_PresenceView
Option Explicit

' 시각화용 시트 이름
Private Const SHEET_PRESENCE As String = "XQLite_Presence"

' 현재 락 인덱스: sheetName -> Dictionary(cellAddr -> nickname)
Private gLockIndex As Object          ' Scripting.Dictionary (late binding)
' 칠한 셀 기록: sheetName -> ArrayList(address)
Private gLockMarks As Object          ' Scripting.Dictionary (late binding)

' 타이머
Private mPresenceAt As Date

' ──────────────────────────────────────────────────────────────────
' 타이머 시작/정지

Public Sub XQL_StartPresenceView()
    On Error Resume Next
    If mPresenceAt <> 0 Then Application.OnTime mPresenceAt, "XQL_Presence_Tick", , False
    Dim sec As Double: sec = XQL_GetPresenceRefreshSec()
    mPresenceAt = Now + (sec / 86400#)
    Application.OnTime mPresenceAt, "XQL_Presence_Tick"
End Sub

Public Sub XQL_StopPresenceView()
    On Error Resume Next
    If mPresenceAt <> 0 Then Application.OnTime mPresenceAt, "XQL_Presence_Tick", , False
    mPresenceAt = 0
End Sub

' 주기적으로 호출: Presence/Locks 갱신
Public Sub XQL_Presence_Tick()
    On Error GoTo done
    Dim prs As Object: Set prs = XQL_FetchPresence()
    Dim lcks As Object: Set lcks = XQL_FetchLocks("") ' 전체 락

    XQL_RenderPresenceSheet prs, lcks
    XQL_UpdateLockShading lcks

done:
    ' 다음 주기 예약
    XQL_StartPresenceView
End Sub

' ──────────────────────────────────────────────────────────────────
' 서버 호출

Public Function XQL_FetchPresence() As Object
    On Error GoTo bad
    Dim q$, resp$, parsed As Object
    q = "query{ presence{ nickname sheet cell updated_at } }"
    resp = XQL_GraphQL(q, "{}")
    If XQL_HasErrors(resp) Then GoTo bad
    Set parsed = JsonConverter.ParseJson(resp)
    Set XQL_FetchPresence = parsed("data")("presence")
    Exit Function
bad:
    Set XQL_FetchPresence = CreateObject("Scripting.Dictionary") ' 빈 dict
End Function

' sheetName가 비어 있으면 전체 락, 아니면 해당 시트만
Public Function XQL_FetchLocks(ByVal sheetName As String) As Object
    On Error GoTo bad
    Dim q$, vars$, resp$, parsed As Object
    If Len(sheetName) = 0 Then
        q = "query{ locks{ sheet cell nickname updated_at } }"
        resp = XQL_GraphQL(q, "{}")
    Else
        q = "query($s:String){ locks(sheet:$s){ sheet cell nickname updated_at } }"
        vars = "{""s"":""" & sheetName & """}"
        resp = XQL_GraphQL(q, vars)
    End If
    If XQL_HasErrors(resp) Then GoTo bad
    Set parsed = JsonConverter.ParseJson(resp)
    Set XQL_FetchLocks = parsed("data")("locks")
    Exit Function
bad:
    Set XQL_FetchLocks = CreateObject("Scripting.Dictionary")
End Function

' ──────────────────────────────────────────────────────────────────
' 렌더링

Private Sub XQL_RenderPresenceSheet(prs As Object, lcks As Object)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets(SHEET_PRESENCE)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = Sheets.add(After:=Sheets(Sheets.count))
        ws.name = SHEET_PRESENCE
    End If

    Application.ScreenUpdating = False
    ws.Cells.Clear

    ' Presence 테이블
    ws.Range("A1:D1").Value = Array("nickname", "sheet", "cell", "updated_at")
    ws.rows(1).Font.Bold = True

    Dim i&, r&
    r = 2
    If Not prs Is Nothing Then
        For i = 1 To prs.count
            ws.Cells(r, 1).Value = prs(i)("nickname")
            ws.Cells(r, 2).Value = prs(i)("sheet")
            ws.Cells(r, 3).Value = prs(i)("cell")
            ws.Cells(r, 4).Value = prs(i)("updated_at")
            r = r + 1
        Next i
    End If

    ' Locks 테이블
    r = r + 1
    ws.Cells(r, 1).Value = "LOCKS"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1
    ws.Range("A" & r & ":D" & r).Value = Array("sheet", "cell", "nickname", "updated_at")
    ws.rows(r).Font.Bold = True
    r = r + 1

    If Not lcks Is Nothing Then
        For i = 1 To lcks.count
            ws.Cells(r, 1).Value = lcks(i)("sheet")
            ws.Cells(r, 2).Value = lcks(i)("cell")
            ws.Cells(r, 3).Value = lcks(i)("nickname")
            ws.Cells(r, 4).Value = lcks(i)("updated_at")
            r = r + 1
        Next i
    End If

    ws.Columns("A:D").AutoFit
    Application.ScreenUpdating = True
End Sub

' 다른 사용자의 락을 시트에 칠하기
Private Sub XQL_UpdateLockShading(lcks As Object)
    If Not XQL_ShowLockShading() Then Exit Sub

    If gLockIndex Is Nothing Then Set gLockIndex = CreateObject("Scripting.Dictionary")
    If gLockMarks Is Nothing Then Set gLockMarks = CreateObject("Scripting.Dictionary")

    ' ? 여기서 한 번만 선언
    Dim ws As Worksheet
    Dim rng As Range
    Dim k As Variant
    Dim arr As Object
    Dim i As Long
    Dim meName As String
    Dim sheetName As String, cellAddr As String, nick As String
    Dim perSheet As Object
    Dim marks As Object

    ' 1) 기존 칠한 것 지우기
    For Each k In gLockMarks.keys
        Set ws = Nothing
        On Error Resume Next: Set ws = Sheets(CStr(k)): On Error GoTo 0
        If Not ws Is Nothing Then
            Set arr = gLockMarks(k) ' ArrayList
            For i = 0 To arr.count - 1
                On Error Resume Next
                Set rng = ws.Range(CStr(arr(i)))
                If Not rng Is Nothing Then
                    If rng.Interior.Color = RGB(255, 230, 230) Then rng.Interior.ColorIndex = xlNone
                End If
                On Error GoTo 0
            Next i
        End If
    Next k
    On Error Resume Next
    gLockMarks.RemoveAll
    gLockIndex.RemoveAll
    On Error GoTo 0

    ' 2) 현재 락 반영
    meName = XQL_GetNickname()
    For i = 1 To lcks.count
        sheetName = CStr(lcks(i)("sheet"))
        cellAddr = CStr(lcks(i)("cell"))
        nick = CStr(lcks(i)("nickname"))
        If LCase$(nick) = LCase$(meName) Then GoTo conti ' 내 락은 하이라이트 안 함

        ' 인덱스 저장
        If Not gLockIndex.Exists(sheetName) Then
            Set perSheet = CreateObject("Scripting.Dictionary")
            gLockIndex.add sheetName, perSheet
        Else
            Set perSheet = gLockIndex(sheetName)
        End If
        perSheet(NormAddr(cellAddr)) = nick

        ' 셀 칠하기(존재할 때, 색이 비어있는 경우만)
        Set ws = Nothing
        On Error Resume Next: Set ws = Sheets(sheetName): On Error GoTo 0
        If Not ws Is Nothing Then
            On Error Resume Next: Set rng = ws.Range(cellAddr): On Error GoTo 0
            If Not rng Is Nothing Then
                If rng.Interior.ColorIndex = xlNone Then
                    rng.Interior.Color = RGB(255, 230, 230) ' 연한 빨강톤: 타인 락
                    ' 기록
                    If Not gLockMarks.Exists(sheetName) Then
                        Set marks = CreateObject("System.Collections.ArrayList")
                        gLockMarks.add sheetName, marks
                    Else
                        Set marks = gLockMarks(sheetName)
                    End If
                    marks.add rng.Address(False, False)
                End If
            End If
        End If
conti:
    Next i
End Sub


' 선택한 범위에 타인 락이 있으면 경고
Public Sub XQL_CheckSelectionLock()
    If Not XQL_WarnOnLockedSelect() Then Exit Sub
    If gLockIndex Is Nothing Then Exit Sub

    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim idx As Object
    If Not gLockIndex.Exists(ws.name) Then Exit Sub
    Set idx = gLockIndex(ws.name)

    Dim c As Range
    For Each c In Selection.Cells
        Dim key$: key = NormAddr(c.Address(False, False))
        If idx.Exists(key) Then
            Dim nick$: nick = CStr(idx(key))
            MsgBox "이 셀은 " & nick & " 님이 잠금 중입니다." & vbCrLf & "(" & ws.name & "!" & key & ")", vbExclamation
            Exit Sub
        End If
    Next c
End Sub

' ──────────────────────────────────────────────────────────────────
' 소도구

Private Function NormAddr(ByVal addr As String) As String
    addr = Replace(addr, "$", "")
    NormAddr = UCase$(addr)
End Function

'Module XQL_Pull
' === Module: XQL_Pull ===
Option Explicit

Private mTickAt As Date

' ─────────────────────── 엔트리 ───────────────────────

Public Sub XQL_Pull_Start()
    If Not XQL_PullEnabled() Then Exit Sub
    ScheduleTick True
End Sub

Public Sub XQL_Pull_Stop()
    On Error Resume Next
    If mTickAt <> 0 Then Application.OnTime mTickAt, "XQL_Pull_Tick", , False
    mTickAt = 0
End Sub

' 수동 즉시 실행(버튼용)
Public Sub XQL_Pull_Now()
    XQL_Pull_Tick
End Sub

' 주기 처리
Public Sub XQL_Pull_Tick()
    On Error GoTo done
    ProcessAllSheets
done:
    ScheduleTick False
End Sub

' ─────────────────────── 본체 ───────────────────────

Private Sub ProcessAllSheets()
    Dim ws As Worksheet
    Dim updatedTables As Object: Set updatedTables = CreateObject("Scripting.Dictionary") ' table -> 1

    For Each ws In ThisWorkbook.Worksheets
        If IsDataSheet(ws) Then
            If ApplyChangesForSheet(ws) > 0 Then
                updatedTables(ws.name) = 1
            End If
        End If
    Next ws

    ' 참조 테이블이 바뀌었으면 ENUM 캐시 자동 리프레시(대상 좁히기 귀찮으면 전체 리프레시)
    If updatedTables.count > 0 Then
        On Error Resume Next
        ' 간단: 전체 리프레시 (필요하면 이후에 영향받은 sheet만 선별 로직 추가 가능)
        XQL_Enums_RefreshAll
        On Error GoTo 0
    End If
End Sub

' 한 시트(=테이블)에 대해 changes를 가져와 적용하고, 적용 건수를 반환
Private Function ApplyChangesForSheet(ws As Worksheet) As Long
    Dim table$: table = ws.name
    Dim since&: since = CLng(GetKV("pull_since:" & table, 0))
    Dim limit&: limit = XQL_PullBatch()

    ' GraphQL 호출
    Dim q$, vars$, resp$, parsed As Object, arr As Object
    q = "query($t:String!,$sv:Int!,$l:Int){changes(table:$t,since_version:$sv,limit:$l){row row_version op}}"
    vars = "{""t"":""" & table & """,""sv"":" & CStr(since) & ",""l"":" & CStr(limit) & "}"
    resp = XQL_GraphQL(q, vars)

    If XQL_HasErrors(resp) Then Exit Function

    Set parsed = JsonConverter.ParseJson(resp)
    Set arr = parsed("data")("changes")
    If arr Is Nothing Or arr.count = 0 Then Exit Function

    Dim applied&, i&, maxrv&: applied = 0: maxrv = since

    Application.ScreenUpdating = False
    XQL_EnsureAuxCols ws
    Dim vcol&, dcol&: vcol = XQL_VerCol(ws): dcol = XQL_DirtyCol(ws)

    For i = 1 To arr.count
        Dim ch As Object: Set ch = arr(i)
        Dim op$, rv&, row As Object, id&
        op = CStr(ch("op"))
        rv = CLng(ch("row_version"))
        If rv > maxrv Then maxrv = rv
        Set row = ch("row")
        id = CLng(val(row("id")))
        If id <= 0 Then GoTo conti

        Select Case UCase$(op)
            Case "UPSERT": applied = applied + ApplyUpsert(ws, id, row, rv, vcol, dcol)
            Case "DELETE": applied = applied + ApplyDelete(ws, id, rv, vcol, dcol)
        End Select
conti:
    Next i
    Application.ScreenUpdating = True

    ' since_version 갱신
    If maxrv > since Then Call SetKV("pull_since:" & table, CStr(maxrv))

    ' 글로벌 최신 버전도 갱신(표시용)
    Dim svrMax&: svrMax = FetchServerMaxRowVersion()
    If svrMax > 0 Then Sheets("XQLite").Range("A7").Offset(0, 1).Value = svrMax

    ApplyChangesForSheet = applied
End Function

' UPSERT 적용: 존재하면 갱신, 없으면 추가
Private Function ApplyUpsert(ws As Worksheet, ByVal id As Long, row As Object, ByVal rv As Long, _
                             ByVal vcol As Long, ByVal dcol As Long) As Long
    Dim rr&: rr = FindRowById(ws, id)
    Dim lc&: lc = XQL_LastCol(ws)

    If rr = 0 Then
        ' 새 행 append
        rr = Max(ws.Cells(ws.rows.count, 1).End(xlUp).row + 1, 3)
        ws.Cells(rr, 1).Value = id
    End If

    ' 각 컬럼 반영(더티 충돌은 건드리지 않고 노란색 표시만)
    Dim c&, k As Variant, curr, incoming
    For c = 2 To lc
        Dim name$: name = CStr(ws.Cells(1, c).Value2)
        If Len(name) = 0 Then GoTo contc
        If row.Exists(name) Then
            incoming = row(name)
            curr = ws.Cells(rr, c).Value

            If dcol > 0 And Len(ws.Cells(rr, dcol).Value2) > 0 Then
                ' 로컬 더티: 값이 다르면 충돌 표시(노랑), 값이 같으면 그대로 둠
                If Not EqLoose(curr, incoming) Then
                    ws.Cells(rr, c).Interior.Color = RGB(255, 255, 180)
                End If
            Else
                ' 로컬 더티 아님: 덮어쓰기
                SetCellValue ws.Cells(rr, c), incoming
                ' 충돌색 제거
                If ws.Cells(rr, c).Interior.Color = RGB(255, 255, 180) Then
                    ws.Cells(rr, c).Interior.ColorIndex = xlNone
                End If
            End If
        End If
contc:
    Next c

    ' deleted 메타가 들어오면 표시(있을 때만)
    On Error Resume Next
    If row.Exists("deleted") Then
        If CBool(row("deleted")) Then
            ws.rows(rr).Font.Strikethrough = True
            ws.rows(rr).Interior.Color = RGB(240, 240, 240)
        Else
            ws.rows(rr).Font.Strikethrough = False
            ws.rows(rr).Interior.ColorIndex = xlNone
        End If
    End If
    On Error GoTo 0

    ' _ver 갱신, 더티 클리어(서버 기준으로 덮어쓴 경우만)
    If vcol > 0 Then ws.Cells(rr, vcol).Value = rv
    ' 더티는 위에서 덮어쓴 셀이 하나라도 있으면 지워도 되지만,
    ' 보수적으로 전체 클리어는 하지 않음(사용자 편집 유지).
    ApplyUpsert = 1
End Function

' DELETE 적용: 소프트 삭제 표시
Private Function ApplyDelete(ws As Worksheet, ByVal id As Long, ByVal rv As Long, _
                             ByVal vcol As Long, ByVal dcol As Long) As Long
    Dim rr&: rr = FindRowById(ws, id)
    If rr = 0 Then
        ' 로컬에 없는 삭제라면 아무 것도 안 함(로그만)
        ApplyDelete = 0: Exit Function
    End If

    ' deleted 컬럼 있으면 true
    Dim delCol&: delCol = FindColumnByHeader(ws, "deleted")
    If delCol > 0 Then ws.Cells(rr, delCol).Value = True

    ' 스타일: 취소선+회색
    ws.rows(rr).Font.Strikethrough = True
    ws.rows(rr).Interior.Color = RGB(240, 240, 240)

    If vcol > 0 Then ws.Cells(rr, vcol).Value = rv
    ApplyDelete = 1
End Function

' ─────────────────────── 헬퍼 ───────────────────────

Private Sub ScheduleTick(ByVal immediate As Boolean)
    On Error Resume Next
    If mTickAt <> 0 Then Application.OnTime mTickAt, "XQL_Pull_Tick", , False
    Dim sec As Double: sec = IIf(immediate, 0.7, XQL_PullSec())
    mTickAt = Now + (sec / 86400#)
    Application.OnTime mTickAt, "XQL_Pull_Tick"
End Sub

Private Function IsDataSheet(ws As Worksheet) As Boolean
    Dim n$: n = ws.name
    IsDataSheet = Not (n = "XQLite" Or n = "XQLite_Conflicts" Or n = "XQLite_Presence" Or n = "XQLite_EnumCache" _
                       Or n = "XQLite_Enums" Or n = "XQLite_Check" Or n = "XQLite_Query" Or n = "XQLite_Audit" _
                       Or n = "XQLite_Outbox" Or n = "XQLite_Resolve")
End Function

' XQLite 시트를 K/V 스토어로 사용
Private Function GetKV(ByVal key As String, ByVal def As Variant) As Variant
    Dim cfg As Worksheet: Set cfg = Sheets("XQLite")
    Dim last&, r&
    last = cfg.Cells(cfg.rows.count, 1).End(xlUp).row
    For r = 1 To last
        If LCase$(CStr(cfg.Cells(r, 1).Value2)) = LCase$(key) Then
            GetKV = cfg.Cells(r, 2).Value
            If IsEmpty(GetKV) Or GetKV = "" Then GetKV = def
            Exit Function
        End If
    Next r
    GetKV = def
End Function

Private Sub SetKV(ByVal key As String, ByVal val As String)
    Dim cfg As Worksheet: Set cfg = Sheets("XQLite")
    Dim last&, r&, found As Boolean
    last = cfg.Cells(cfg.rows.count, 1).End(xlUp).row
    For r = 1 To last
        If LCase$(CStr(cfg.Cells(r, 1).Value2)) = LCase$(key) Then
            cfg.Cells(r, 2).Value = val
            found = True: Exit For
        End If
    Next r
    If Not found Then
        cfg.Cells(last + 1, 1).Value = key
        cfg.Cells(last + 1, 2).Value = val
    End If
End Sub

' 공용 유틸(다른 모듈과 동일 시그니처)

Private Function FindRowById(ws As Worksheet, ByVal id As Long) As Long
    Dim lastRow&, r&
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = 3 To lastRow
        If CLng(val(ws.Cells(r, 1).Value2)) = id Then FindRowById = r: Exit Function
    Next
    FindRowById = 0
End Function

Private Function FindColumnByHeader(ws As Worksheet, ByVal header As String) As Long
    Dim lastCol&, c&
    lastCol = XQL_LastCol(ws)
    For c = 1 To lastCol
        If LCase$(CStr(ws.Cells(1, c).Value2)) = LCase$(header) Then
            FindColumnByHeader = c: Exit Function
        End If
    Next
    FindColumnByHeader = 0
End Function

Private Function EqLoose(a As Variant, b As Variant) As Boolean
    If IsEmpty(a) And (IsEmpty(b) Or b = "" Or b Is Nothing) Then EqLoose = True: Exit Function
    If IsEmpty(b) And (IsEmpty(a) Or a = "" Or a Is Nothing) Then EqLoose = True: Exit Function
    On Error Resume Next
    EqLoose = (CStr(a) = CStr(b))
End Function

Private Sub SetCellValue(ByVal cell As Range, ByVal v As Variant)
    If IsObject(v) Then
        cell.Value = CStr(v)
    Else
        cell.Value = v
    End If
End Sub

Private Function FetchServerMaxRowVersion() As Long
    On Error GoTo bad
    Dim q$, resp$, parsed As Object
    q = "query{ meta{ max_row_version } }"
    resp = XQL_GraphQL(q, "{}")
    Set parsed = JsonConverter.ParseJson(resp)
    FetchServerMaxRowVersion = CLng(parsed("data")("meta")("max_row_version"))
    Exit Function
bad:
    FetchServerMaxRowVersion = 0
End Function

'Module XQL_PullDelta
' === Module: XQL_Pull ===
Option Explicit

Private mTickAt As Date

' ─────────────────────── 엔트리 ───────────────────────

Public Sub XQL_Pull_Start()
    If Not XQL_PullEnabled() Then Exit Sub
    ScheduleTick True
End Sub

Public Sub XQL_Pull_Stop()
    On Error Resume Next
    If mTickAt <> 0 Then Application.OnTime mTickAt, "XQL_Pull_Tick", , False
    mTickAt = 0
End Sub

' 수동 즉시 실행(버튼용)
Public Sub XQL_Pull_Now()
    XQL_Pull_Tick
End Sub

' 주기 처리
Public Sub XQL_Pull_Tick()
    On Error GoTo done
    ProcessAllSheets
done:
    ScheduleTick False
End Sub

' ─────────────────────── 본체 ───────────────────────

Private Sub ProcessAllSheets()
    Dim ws As Worksheet
    Dim updatedTables As Object: Set updatedTables = CreateObject("Scripting.Dictionary") ' table -> 1

    For Each ws In ThisWorkbook.Worksheets
        If IsDataSheet(ws) Then
            If ApplyChangesForSheet(ws) > 0 Then
                updatedTables(ws.name) = 1
            End If
        End If
    Next ws

    ' 참조 테이블이 바뀌었으면 ENUM 캐시 자동 리프레시(대상 좁히기 귀찮으면 전체 리프레시)
    If updatedTables.count > 0 Then
        On Error Resume Next
        ' 간단: 전체 리프레시 (필요하면 이후에 영향받은 sheet만 선별 로직 추가 가능)
        XQL_Enums_RefreshAll
        On Error GoTo 0
    End If
End Sub

' 한 시트(=테이블)에 대해 changes를 가져와 적용하고, 적용 건수를 반환
Private Function ApplyChangesForSheet(ws As Worksheet) As Long
    Dim table$: table = ws.name
    Dim since&: since = CLng(GetKV("pull_since:" & table, 0))
    Dim limit&: limit = XQL_PullBatch()

    ' GraphQL 호출
    Dim q$, vars$, resp$, parsed As Object, arr As Object
    q = "query($t:String!,$sv:Int!,$l:Int){changes(table:$t,since_version:$sv,limit:$l){row row_version op}}"
    vars = "{""t"":""" & table & """,""sv"":" & CStr(since) & ",""l"":" & CStr(limit) & "}"
    resp = XQL_GraphQL(q, vars)

    If XQL_HasErrors(resp) Then Exit Function

    Set parsed = JsonConverter.ParseJson(resp)
    Set arr = parsed("data")("changes")
    If arr Is Nothing Or arr.count = 0 Then Exit Function

    Dim applied&, i&, maxrv&: applied = 0: maxrv = since

    Application.ScreenUpdating = False
    XQL_EnsureAuxCols ws
    Dim vcol&, dcol&: vcol = XQL_VerCol(ws): dcol = XQL_DirtyCol(ws)

    For i = 1 To arr.count
        Dim ch As Object: Set ch = arr(i)
        Dim op$, rv&, row As Object, id&
        op = CStr(ch("op"))
        rv = CLng(ch("row_version"))
        If rv > maxrv Then maxrv = rv
        Set row = ch("row")
        id = CLng(val(row("id")))
        If id <= 0 Then GoTo conti

        Select Case UCase$(op)
            Case "UPSERT": applied = applied + ApplyUpsert(ws, id, row, rv, vcol, dcol)
            Case "DELETE": applied = applied + ApplyDelete(ws, id, rv, vcol, dcol)
        End Select
conti:
    Next i
    Application.ScreenUpdating = True

    ' since_version 갱신
    If maxrv > since Then Call SetKV("pull_since:" & table, CStr(maxrv))

    ' 글로벌 최신 버전도 갱신(표시용)
    Dim svrMax&: svrMax = FetchServerMaxRowVersion()
    If svrMax > 0 Then Sheets("XQLite").Range("A7").Offset(0, 1).Value = svrMax

    ApplyChangesForSheet = applied
End Function

' UPSERT 적용: 존재하면 갱신, 없으면 추가
Private Function ApplyUpsert(ws As Worksheet, ByVal id As Long, row As Object, ByVal rv As Long, _
                             ByVal vcol As Long, ByVal dcol As Long) As Long
    Dim rr&: rr = FindRowById(ws, id)
    Dim lc&: lc = XQL_LastCol(ws)

    If rr = 0 Then
        ' 새 행 append
        rr = Max(ws.Cells(ws.rows.count, 1).End(xlUp).row + 1, 3)
        ws.Cells(rr, 1).Value = id
    End If

    ' 각 컬럼 반영(더티 충돌은 건드리지 않고 노란색 표시만)
    Dim c&, k As Variant, curr, incoming
    For c = 2 To lc
        Dim name$: name = CStr(ws.Cells(1, c).Value2)
        If Len(name) = 0 Then GoTo contc
        If row.Exists(name) Then
            incoming = row(name)
            curr = ws.Cells(rr, c).Value

            If dcol > 0 And Len(ws.Cells(rr, dcol).Value2) > 0 Then
                ' 로컬 더티: 값이 다르면 충돌 표시(노랑), 값이 같으면 그대로 둠
                If Not EqLoose(curr, incoming) Then
                    ws.Cells(rr, c).Interior.Color = RGB(255, 255, 180)
                End If
            Else
                ' 로컬 더티 아님: 덮어쓰기
                SetCellValue ws.Cells(rr, c), incoming
                ' 충돌색 제거
                If ws.Cells(rr, c).Interior.Color = RGB(255, 255, 180) Then
                    ws.Cells(rr, c).Interior.ColorIndex = xlNone
                End If
            End If
        End If
contc:
    Next c

    ' deleted 메타가 들어오면 표시(있을 때만)
    On Error Resume Next
    If row.Exists("deleted") Then
        If CBool(row("deleted")) Then
            ws.rows(rr).Font.Strikethrough = True
            ws.rows(rr).Interior.Color = RGB(240, 240, 240)
        Else
            ws.rows(rr).Font.Strikethrough = False
            ws.rows(rr).Interior.ColorIndex = xlNone
        End If
    End If
    On Error GoTo 0

    ' _ver 갱신, 더티 클리어(서버 기준으로 덮어쓴 경우만)
    If vcol > 0 Then ws.Cells(rr, vcol).Value = rv
    ' 더티는 위에서 덮어쓴 셀이 하나라도 있으면 지워도 되지만,
    ' 보수적으로 전체 클리어는 하지 않음(사용자 편집 유지).
    ApplyUpsert = 1
End Function

' DELETE 적용: 소프트 삭제 표시
Private Function ApplyDelete(ws As Worksheet, ByVal id As Long, ByVal rv As Long, _
                             ByVal vcol As Long, ByVal dcol As Long) As Long
    Dim rr&: rr = FindRowById(ws, id)
    If rr = 0 Then
        ' 로컬에 없는 삭제라면 아무 것도 안 함(로그만)
        ApplyDelete = 0: Exit Function
    End If

    ' deleted 컬럼 있으면 true
    Dim delCol&: delCol = FindColumnByHeader(ws, "deleted")
    If delCol > 0 Then ws.Cells(rr, delCol).Value = True

    ' 스타일: 취소선+회색
    ws.rows(rr).Font.Strikethrough = True
    ws.rows(rr).Interior.Color = RGB(240, 240, 240)

    If vcol > 0 Then ws.Cells(rr, vcol).Value = rv
    ApplyDelete = 1
End Function

' ─────────────────────── 헬퍼 ───────────────────────

Private Sub ScheduleTick(ByVal immediate As Boolean)
    On Error Resume Next
    If mTickAt <> 0 Then Application.OnTime mTickAt, "XQL_Pull_Tick", , False
    Dim sec As Double: sec = IIf(immediate, 0.7, XQL_PullSec())
    mTickAt = Now + (sec / 86400#)
    Application.OnTime mTickAt, "XQL_Pull_Tick"
End Sub

Private Function IsDataSheet(ws As Worksheet) As Boolean
    Dim n$: n = ws.name
    IsDataSheet = Not (n = "XQLite" Or n = "XQLite_Conflicts" Or n = "XQLite_Presence" Or n = "XQLite_EnumCache" _
                       Or n = "XQLite_Enums" Or n = "XQLite_Check" Or n = "XQLite_Query" Or n = "XQLite_Audit" _
                       Or n = "XQLite_Outbox" Or n = "XQLite_Resolve")
End Function

' XQLite 시트를 K/V 스토어로 사용
Private Function GetKV(ByVal key As String, ByVal def As Variant) As Variant
    Dim cfg As Worksheet: Set cfg = Sheets("XQLite")
    Dim last&, r&
    last = cfg.Cells(cfg.rows.count, 1).End(xlUp).row
    For r = 1 To last
        If LCase$(CStr(cfg.Cells(r, 1).Value2)) = LCase$(key) Then
            GetKV = cfg.Cells(r, 2).Value
            If IsEmpty(GetKV) Or GetKV = "" Then GetKV = def
            Exit Function
        End If
    Next r
    GetKV = def
End Function

Private Sub SetKV(ByVal key As String, ByVal val As String)
    Dim cfg As Worksheet: Set cfg = Sheets("XQLite")
    Dim last&, r&, found As Boolean
    last = cfg.Cells(cfg.rows.count, 1).End(xlUp).row
    For r = 1 To last
        If LCase$(CStr(cfg.Cells(r, 1).Value2)) = LCase$(key) Then
            cfg.Cells(r, 2).Value = val
            found = True: Exit For
        End If
    Next r
    If Not found Then
        cfg.Cells(last + 1, 1).Value = key
        cfg.Cells(last + 1, 2).Value = val
    End If
End Sub

' 공용 유틸(다른 모듈과 동일 시그니처)

Private Function FindRowById(ws As Worksheet, ByVal id As Long) As Long
    Dim lastRow&, r&
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = 3 To lastRow
        If CLng(val(ws.Cells(r, 1).Value2)) = id Then FindRowById = r: Exit Function
    Next
    FindRowById = 0
End Function

Private Function FindColumnByHeader(ws As Worksheet, ByVal header As String) As Long
    Dim lastCol&, c&
    lastCol = XQL_LastCol(ws)
    For c = 1 To lastCol
        If LCase$(CStr(ws.Cells(1, c).Value2)) = LCase$(header) Then
            FindColumnByHeader = c: Exit Function
        End If
    Next
    FindColumnByHeader = 0
End Function

Private Function EqLoose(a As Variant, b As Variant) As Boolean
    If IsEmpty(a) And (IsEmpty(b) Or b = "" Or b Is Nothing) Then EqLoose = True: Exit Function
    If IsEmpty(b) And (IsEmpty(a) Or a = "" Or a Is Nothing) Then EqLoose = True: Exit Function
    On Error Resume Next
    EqLoose = (CStr(a) = CStr(b))
End Function

Private Sub SetCellValue(ByVal cell As Range, ByVal v As Variant)
    If IsObject(v) Then
        cell.Value = CStr(v)
    Else
        cell.Value = v
    End If
End Sub

Private Function FetchServerMaxRowVersion() As Long
    On Error GoTo bad
    Dim q$, resp$, parsed As Object
    q = "query{ meta{ max_row_version } }"
    resp = XQL_GraphQL(q, "{}")
    Set parsed = JsonConverter.ParseJson(resp)
    FetchServerMaxRowVersion = CLng(parsed("data")("meta")("max_row_version"))
    Exit Function
bad:
    FetchServerMaxRowVersion = 0
End Function

'Module XQL_QueryPanel
' === Module: XQL_QueryPanel ===
Option Explicit

Private Const SHEET_Q As String = "XQLite_Query"

' ─────────────────────────────────────────────────────────────────
' 1) 템플릿 생성

Public Sub XQL_Query_Setup()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets(SHEET_Q)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = Sheets.add(After:=Sheets(Sheets.count))
        ws.name = SHEET_Q
    Else
        ws.Cells.Clear
    End If

    With ws
        .Range("A1").Value = "Table"
        .Range("B1").Value = IIf(ActiveSheet.name <> SHEET_Q And ActiveSheet.name <> "XQLite", ActiveSheet.name, "")
        .Range("A2").Value = "Include Deleted"
        .Range("B2").Value = "FALSE"
        .Range("A3").Value = "Limit"
        .Range("B3").Value = 100
        .Range("A4").Value = "Offset"
        .Range("B4").Value = 0

        .Range("A6").Value = "Filters"
        .Range("A7:D7").Value = Array("Column", "Op(=,!=,>,>=,<,<=,LIKE,IN)", "Value(콤마로 다중)", "Type(AUTO|NUM|TEXT)")

        Dim r As Long
        For r = 8 To 17
            .Range("A" & r & ":D" & r).ClearContents
        Next r

        .Range("A19").Value = "Order By"
        .Range("A20:B20").Value = Array("Column", "Dir(ASC|DESC)")
        .Range("A21:B22").ClearContents
        .Range("A23:B23").ClearContents

        .Range("A26").Value = "Actions"
        .Range("A27").Value = "▶ Run Query"
        .Range("B27").Value = "◀ Prev"
        .Range("C27").Value = "Next ▶"
        .Range("D27").Value = "? Reset"

        .Columns("A:G").AutoFit
        .rows("1:7").Font.Bold = True
    End With

    ' 단추 바인딩(양식 컨트롤 없이 셀-더블클릭으로도 가능하지만, 매크로에 바로 매핑 추천)
    MsgBox "XQLite_Query 시트를 생성했습니다." & vbCrLf & _
           "- B1: 테이블명, B2: include_deleted, B3/B4: limit/offset" & vbCrLf & _
           "- A8:D17: 필터 그리드" & vbCrLf & _
           "- A20:B23: 정렬 조건" & vbCrLf & _
           "- A27:D27: 실행/페이지/초기화", vbInformation
           
    ' --- VIEW 토글/이름(우측 H열) ---
    With ws
        .Range("H1").Value = "USE_VIEW"
        .Range("H2").Value = "FALSE"
        .Range("H3").Value = "VIEW_NAME"
        .Range("H4").Value = ""
    End With
End Sub

' ─────────────────────────────────────────────────────────────────
' 2) 실행/페이지/초기화 엔트리

Public Sub XQL_Query_Run()
    Dim ws As Worksheet: Set ws = EnsureQuerySheet()
    Dim table$, wraw$, ob$, lim&, off&, incDel As Boolean
    
    Dim useView As Boolean, viewName$
    useView = (UCase$(CStr(ws.Range("H2").Value2)) = "TRUE")
    viewName = Trim$(CStr(ws.Range("H4").Value2))

    If Not ReadHeader(ws, table, incDel, lim, off) Then Exit Sub
    wraw = BuildWhereRaw(ws)
    ob = BuildOrderBy(ws)

    RunAndRender ws, table, wraw, ob, lim, off, incDel, useView, viewName
End Sub

Public Sub XQL_Query_Next()
    Dim ws As Worksheet: Set ws = EnsureQuerySheet()
    Dim off&, lim&: lim = CLng(Nz(ws.Range("B3").Value2, 100))
    off = CLng(Nz(ws.Range("B4").Value2, 0)) + lim
    ws.Range("B4").Value = off
    XQL_Query_Run
End Sub

Public Sub XQL_Query_Prev()
    Dim ws As Worksheet: Set ws = EnsureQuerySheet()
    Dim off&, lim&: lim = CLng(Nz(ws.Range("B3").Value2, 100))
    off = CLng(Nz(ws.Range("B4").Value2, 0)) - lim
    If off < 0 Then off = 0
    ws.Range("B4").Value = off
    XQL_Query_Run
End Sub

Public Sub XQL_Query_Reset()
    Dim ws As Worksheet: Set ws = EnsureQuerySheet()
    ws.Range("B2").Value = "FALSE"
    ws.Range("B3").Value = 100
    ws.Range("B4").Value = 0
    ws.Range("A8:D17").ClearContents
    ws.Range("A21:B23").ClearContents
    ClearResult ws
End Sub

Public Sub XQL_Query_SavePreset()
    Dim ws As Worksheet: Set ws = EnsureQuerySheet()
    Dim name$: name = InputBox("프리셋 이름을 입력하세요:", "Save Preset", "")
    If Len(name) = 0 Then Exit Sub

    Dim p As Worksheet: Set p = QP_EnsurePresetSheet()
    p.Visible = xlSheetVisible

    Dim useView As Boolean, source$, incDel As Boolean, lim&, off&
    useView = (UCase$(CStr(ws.Range("H2").Value2)) = "TRUE")
    source = IIf(useView, Trim$(CStr(ws.Range("H4").Value2)), Trim$(CStr(ws.Range("B1").Value2)))
    If Len(source) = 0 Then MsgBox "테이블명(B1) 또는 VIEW_NAME(H4)을 채워주세요.", vbExclamation: GoTo done

    incDel = (UCase$(CStr(ws.Range("B2").Value2)) = "TRUE")
    lim = CLng(Nz(ws.Range("B3").Value2, 100))
    off = CLng(Nz(ws.Range("B4").Value2, 0))

    ' 필터/정렬 직렬화
    Dim r&, filters$, order$
    For r = 8 To 17
        Dim c1$, c2$, c3$, c4$
        c1 = Trim$(CStr(ws.Range("A" & r).Value2))
        c2 = Trim$(CStr(ws.Range("B" & r).Value2))
        c3 = Trim$(CStr(ws.Range("C" & r).Value2))
        c4 = Trim$(CStr(ws.Range("D" & r).Value2))
        If Len(c1) > 0 Or Len(c2) > 0 Or Len(c3) > 0 Or Len(c4) > 0 Then
            If Len(filters) > 0 Then filters = filters & ";"
            filters = filters & QP_Enc(c1 & "|" & c2 & "|" & c3 & "|" & c4)
        End If
    Next r
    For r = 21 To 23
        Dim oc$, od$
        oc = Trim$(CStr(ws.Range("A" & r).Value2))
        od = Trim$(CStr(ws.Range("B" & r).Value2))
        If Len(oc) > 0 Then
            If Len(order) > 0 Then order = order & ";"
            order = order & QP_Enc(oc & "|" & od)
        End If
    Next r

    ' 존재 시 업데이트, 없으면 append
    Dim last&, i&, found As Boolean: last = p.Cells(p.rows.count, 1).End(xlUp).row
    For i = 2 To last
        If LCase$(CStr(p.Cells(i, 1).Value2)) = LCase$(name) Then
            found = True: Exit For
        End If
    Next i
    If Not found Then i = last + 1

    p.Cells(i, 1).Value = name
    p.Cells(i, 2).Value = IIf(useView, "TRUE", "FALSE")
    p.Cells(i, 3).Value = source
    p.Cells(i, 4).Value = IIf(incDel, "TRUE", "FALSE")
    p.Cells(i, 5).Value = lim
    p.Cells(i, 6).Value = off
    p.Cells(i, 7).Value = filters
    p.Cells(i, 8).Value = order

    MsgBox "프리셋 저장 완료: " & name, vbInformation
done:
    p.Visible = xlSheetVeryHidden
End Sub

Public Sub XQL_Query_LoadPreset()
    Dim p As Worksheet: Set p = QP_EnsurePresetSheet()
    p.Visible = xlSheetVisible
    Dim name$: name = InputBox("불러올 프리셋 이름:", "Load Preset", "")
    If Len(name) = 0 Then GoTo done

    Dim i&, last&: last = p.Cells(p.rows.count, 1).End(xlUp).row
    Dim row&: row = 0
    For i = 2 To last
        If LCase$(CStr(p.Cells(i, 1).Value2)) = LCase$(name) Then row = i: Exit For
    Next i
    If row = 0 Then MsgBox "프리셋을 찾을 수 없습니다: " & name, vbExclamation: GoTo done

    Dim ws As Worksheet: Set ws = EnsureQuerySheet()
    Dim useView As Boolean: useView = (UCase$(CStr(p.Cells(row, 2).Value2)) = "TRUE")
    Dim source$, filters$, order$: source = CStr(p.Cells(row, 3).Value2)
    ws.Range("B2").Value = CStr(p.Cells(row, 4).Value2)
    ws.Range("B3").Value = CLng(Nz(p.Cells(row, 5).Value2, 100))
    ws.Range("B4").Value = CLng(Nz(p.Cells(row, 6).Value2, 0))
    filters = CStr(p.Cells(row, 7).Value2)
    order = CStr(p.Cells(row, 8).Value2)

    ' VIEW / TABLE 반영
    ws.Range("H2").Value = IIf(useView, "TRUE", "FALSE")
    If useView Then
        ws.Range("H4").Value = source
        ws.Range("B1").Value = ""        ' 테이블칸 비움
    Else
        ws.Range("B1").Value = source
        ws.Range("H4").Value = ""
    End If

    ' 그리드 클리어 후 반영
    ws.Range("A8:D17").ClearContents
    ws.Range("A21:B23").ClearContents

    Dim parts() As String, j&, seg$, cols() As String
    If Len(filters) > 0 Then
        parts = Split(filters, ";")
        For j = 0 To WorksheetFunction.Min(UBound(parts), 9)
            seg = QP_Dec(parts(j))
            cols = Split(seg, "|")
            If UBound(cols) >= 0 Then ws.Range("A" & (8 + j)).Value = cols(0)
            If UBound(cols) >= 1 Then ws.Range("B" & (8 + j)).Value = cols(1)
            If UBound(cols) >= 2 Then ws.Range("C" & (8 + j)).Value = cols(2)
            If UBound(cols) >= 3 Then ws.Range("D" & (8 + j)).Value = cols(3)
        Next j
    End If
    If Len(order) > 0 Then
        parts = Split(order, ";")
        For j = 0 To WorksheetFunction.Min(UBound(parts), 2)
            seg = QP_Dec(parts(j))
            cols = Split(seg, "|")
            If UBound(cols) >= 0 Then ws.Range("A" & (21 + j)).Value = cols(0)
            If UBound(cols) >= 1 Then ws.Range("B" & (21 + j)).Value = cols(1)
        Next j
    End If

    MsgBox "프리셋 불러오기 완료: " & name, vbInformation
done:
    p.Visible = xlSheetVeryHidden
End Sub

Public Sub XQL_Query_DeletePreset()
    Dim p As Worksheet: Set p = QP_EnsurePresetSheet()
    p.Visible = xlSheetVisible
    Dim name$: name = InputBox("삭제할 프리셋 이름:", "Delete Preset", "")
    If Len(name) = 0 Then GoTo done

    Dim i&, last&: last = p.Cells(p.rows.count, 1).End(xlUp).row
    For i = 2 To last
        If LCase$(CStr(p.Cells(i, 1).Value2)) = LCase$(name) Then
            p.rows(i).Delete
            MsgBox "삭제 완료: " & name, vbInformation
            GoTo done
        End If
    Next i
    MsgBox "프리셋을 찾지 못했습니다: " & name, vbExclamation
done:
    p.Visible = xlSheetVeryHidden
End Sub

Public Sub XQL_Query_ListPresets()
    Dim p As Worksheet: Set p = QP_EnsurePresetSheet()
    Dim last&, i&, s$
    last = p.Cells(p.rows.count, 1).End(xlUp).row
    If last < 2 Then MsgBox "(저장된 프리셋 없음)", vbInformation: Exit Sub
    For i = 2 To last
        s = s & "- " & CStr(p.Cells(i, 1).Value2) & vbCrLf
    Next i
    MsgBox s, vbInformation, "Query Presets"
End Sub


' ─────────────────────────────────────────────────────────────────
' 3) 내부: 실행 & 렌더

Private Sub RunAndRender(ws As Worksheet, ByVal table As String, ByVal wraw As String, ByVal ob As String, _
                         ByVal lim As Long, ByVal off As Long, ByVal incDel As Boolean, _
                         ByVal useView As Boolean, ByVal viewName As String)
    Dim q$, vars$, resp$, parsed As Object, arr As Object, mx&, ok As Boolean

    ClearResult ws

    If useView Then
        If Len(viewName) = 0 Then
            MsgBox "USE_VIEW=TRUE 인데 VIEW_NAME이 비어 있습니다.", vbExclamation
            Exit Sub
        End If
        ' 1) rowsView 시도
        q = "query($v:String!,$w:String,$o:String,$l:Int,$f:Int,$inc:Boolean){" & _
            "rowsView(view:$v,whereRaw:$w,orderBy:$o,limit:$l,offset:$f,include_deleted:$inc){rows max_row_version affected}}"
        vars = "{""v"":""" & viewName & """,""w"":" & JStrOrNull(wraw) & ",""o"":" & JStrOrNull(ob) & _
               ",""l"":" & CStr(lim) & ",""f"":" & CStr(off) & ",""inc"":" & LCase$(CStr(incDel)) & "}"
        resp = XQL_GraphQL(q, vars)
        If Not XQL_HasErrors(resp) Then
            ok = True
        ElseIf InStr(1, resp, "Cannot query field ""rowsView""", vbTextCompare) > 0 Then
            MsgBox "서버가 rowsView를 아직 지원하지 않습니다." & vbCrLf & _
                   "- VIEW 이름: " & viewName & vbCrLf & _
                   "- 안내: 서버에 rowsView(view, whereRaw, orderBy, limit, offset, include_deleted) 리졸버를 추가하세요.", vbInformation
            ok = False
        Else
            ' 다른 오류: 그대로 노출
            MsgBox "rowsView 실패: " & resp, vbExclamation
            Exit Sub
        End If
    End If

    If Not ok Then
        ' 2) TABLE 모드로 실행
        q = "query($t:String!,$w:String,$o:String,$l:Int,$f:Int,$inc:Boolean){" & _
            "rows(table:$t,whereRaw:$w,orderBy:$o,limit:$l,offset:$f,include_deleted:$inc){rows max_row_version affected}}"
        vars = "{""t"":""" & table & """,""w"":" & JStrOrNull(wraw) & ",""o"":" & JStrOrNull(ob) & _
               ",""l"":" & CStr(lim) & ",""f"":" & CStr(off) & ",""inc"":" & LCase$(CStr(incDel)) & "}"
        resp = XQL_GraphQL(q, vars)
        If XQL_HasErrors(resp) Then
            MsgBox "Query 실패: " & resp, vbExclamation
            Exit Sub
        End If
    End If

    Set parsed = JsonConverter.ParseJson(resp)
    If useView And ok Then
        Set arr = parsed("data")("rowsView")("rows")
        mx = CLng(Nz(parsed("data")("rowsView")("max_row_version"), 0))
    Else
        Set arr = parsed("data")("rows")("rows")
        mx = CLng(Nz(parsed("data")("rows")("max_row_version"), 0))
    End If

    If arr Is Nothing Or arr.count = 0 Then
        ws.Range("A30").Value = "(no rows)"
        ws.Range("A28").Value = "max_row_version:": ws.Range("B28").Value = mx
        Exit Sub
    End If

    Dim keys As Object: Set keys = CollectKeys(arr)
    RenderHeader ws, keys
    RenderRows ws, keys, arr
    ws.Range("A28").Value = "max_row_version:": ws.Range("B28").Value = mx
End Sub


Private Function CollectKeys(arr As Object) As Object
    Dim s As Object: Set s = CreateObject("System.Collections.ArrayList")
    ' 메타 우선
    s.add "id": s.add "row_version": s.add "updated_at": s.add "deleted"
    ' 첫 로우에서 나머지 키 수집
    Dim k As Variant, o As Object: Set o = arr(1)
    For Each k In o.keys
        If Not ContainsCI(s, CStr(k)) Then s.add CStr(k)
    Next k
    Set CollectKeys = s
End Function

Private Sub RenderHeader(ws As Worksheet, keys As Object)
    Dim i&, base&: base = 30
    For i = 0 To keys.count - 1
        ws.Cells(base, i + 1).Value = keys(i)
    Next i
    ws.rows(base).Font.Bold = True
End Sub

Private Sub RenderRows(ws As Worksheet, keys As Object, arr As Object)
    Dim r&, i&, base&: base = 31
    Dim row As Object, k As Variant, v

    Application.ScreenUpdating = False
    For r = 1 To arr.count
        Set row = arr(r)
        For i = 0 To keys.count - 1
            k = keys(i)
            If row.Exists(k) Then
                v = row(k)
                If IsObject(v) Then
                    ws.Cells(base + r - 1, i + 1).Value = CStr(v)
                Else
                    ws.Cells(base + r - 1, i + 1).Value = v
                End If
            Else
                ws.Cells(base + r - 1, i + 1).Value = Empty
            End If
        Next i
    Next r
    ws.Columns("A:Z").AutoFit
    Application.ScreenUpdating = True
End Sub

Private Sub ClearResult(ws As Worksheet)
    ws.Range("A28:Z1000000").ClearContents
    ws.rows(30).Font.Bold = True
End Sub

' ─────────────────────────────────────────────────────────────────
' 4) 빌더: whereRaw / orderBy

Private Function BuildWhereRaw(ws As Worksheet) As String
    Dim r As Long, out As String, col$, op$, val$, ty$
    For r = 8 To 17
        col = Trim$(CStr(ws.Range("A" & r).Value2))
        op = UCase$(Trim$(CStr(ws.Range("B" & r).Value2)))
        val = Trim$(CStr(ws.Range("C" & r).Value2))
        ty = UCase$(Trim$(CStr(ws.Range("D" & r).Value2)))
        If Len(col) = 0 Or Len(op) = 0 Or Len(val) = 0 Then GoTo cont

        If op = "LIKE" Then
            out = AppendCond(out, col & " LIKE " & JQuote("%" & val & "%"))
        ElseIf op = "IN" Then
            out = AppendCond(out, col & " IN (" & BuildInList(val, ty) & ")")
        ElseIf (op = "=" Or op = "!=" Or op = ">" Or op = ">=" Or op = "<" Or op = "<=") Then
            If IsNumType(ty) Or IsNumeric(val) Then
                out = AppendCond(out, col & " " & op & " " & CStr(val(val)))
            Else
                out = AppendCond(out, col & " " & op & " " & JQuote(val))
            End If
        End If
cont:
    Next r

    If Len(out) = 0 Then
        BuildWhereRaw = vbNullString
    Else
        BuildWhereRaw = out
    End If
End Function

Private Function BuildOrderBy(ws As Worksheet) As String
    Dim i As Long, col$, Dir$, out$
    For i = 21 To 23
        col = Trim$(CStr(ws.Range("A" & i).Value2))
        Dir = UCase$(Trim$(CStr(ws.Range("B" & i).Value2)))
        If Len(col) = 0 Then GoTo cont
        If Dir <> "DESC" Then Dir = "ASC"
        If Len(out) > 0 Then out = out & ", "
        out = out & col & " " & Dir
cont:
    Next i
    BuildOrderBy = out
End Function

Private Function AppendCond(ByVal base As String, ByVal cond As String) As String
    If Len(base) = 0 Then AppendCond = cond Else AppendCond = base & " AND " & cond
End Function

Private Function BuildInList(ByVal csv As String, ByVal ty As String) As String
    Dim parts() As String, i&, p$, out$
    parts = Split(csv, ",")
    For i = LBound(parts) To UBound(parts)
        p = Trim$(parts(i))
        If Len(p) = 0 Then GoTo cont
        If IsNumType(ty) Or IsNumeric(p) Then
            If Len(out) > 0 Then out = out & ","
            out = out & CStr(val(p))
        Else
            If Len(out) > 0 Then out = out & ","
            out = out & JQuote(p)
        End If
cont:
    Next i
    BuildInList = out
End Function

' ─────────────────────────────────────────────────────────────────
' 5) 유틸

Private Function EnsureQuerySheet() As Worksheet
    On Error Resume Next
    Set EnsureQuerySheet = Sheets(SHEET_Q)
    On Error GoTo 0
    If EnsureQuerySheet Is Nothing Then
        XQL_Query_Setup
        Set EnsureQuerySheet = Sheets(SHEET_Q)
    End If
End Function

Private Function ReadHeader(ws As Worksheet, ByRef table As String, ByRef incDel As Boolean, ByRef lim As Long, ByRef off As Long) As Boolean
    table = Trim$(CStr(ws.Range("B1").Value2))
    If Len(table) = 0 Then
        MsgBox "B1에 테이블명을 입력하세요.", vbExclamation
        ReadHeader = False: Exit Function
    End If
    incDel = CBool(IIf(UCase$(CStr(ws.Range("B2").Value2)) = "TRUE", True, False))
    lim = CLng(Nz(ws.Range("B3").Value2, 100))
    off = CLng(Nz(ws.Range("B4").Value2, 0))
    If lim <= 0 Then lim = 100
    If off < 0 Then off = 0
    ReadHeader = True
End Function

Private Function JQuote(ByVal s As String) As String
    JQuote = """" & Replace(s, """", "\""") & """"
End Function

Private Function JStrOrNull(ByVal s As String) As String
    If Len(Trim$(s)) = 0 Then
        JStrOrNull = "null"
    Else
        JStrOrNull = JQuote(s)
    End If
End Function

Private Function ContainsCI(list As Object, ByVal key As String) As Boolean
    Dim i&
    For i = 0 To list.count - 1
        If LCase$(CStr(list(i))) = LCase$(key) Then ContainsCI = True: Exit Function
    Next i
End Function

Private Function IsNumType(ByVal t As String) As Boolean
    Dim u$: u = UCase$(Trim$(t))
    IsNumType = (InStr(u, "INT") > 0 Or u = "REAL" Or u = "FLOAT" Or u = "DOUBLE" Or u = "NUM")
End Function

Private Function Nz(ByVal v As Variant, ByVal fb As Variant) As Variant
    If IsEmpty(v) Or v = "" Or v Is Nothing Then Nz = fb Else Nz = v
End Function

Private Function QP_Enc(ByVal s As String) As String
    s = Replace(s, "|", "%7C")
    s = Replace(s, ";", "%3B")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    QP_Enc = s
End Function

Private Function QP_Dec(ByVal s As String) As String
    s = Replace(s, "%7C", "|")
    s = Replace(s, "%3B", ";")
    s = Replace(s, "\n", vbLf)
    QP_Dec = s
End Function

Private Function QP_EnsurePresetSheet() As Worksheet
    On Error Resume Next: Set QP_EnsurePresetSheet = Sheets("XQLite_QueryPresets"): On Error GoTo 0
    If QP_EnsurePresetSheet Is Nothing Then
        Set QP_EnsurePresetSheet = Sheets.add(After:=Sheets(Sheets.count))
        QP_EnsurePresetSheet.name = "XQLite_QueryPresets"
        With QP_EnsurePresetSheet
            .Range("A1:H1").Value = Array("name", "use_view", "source", "include_deleted", "limit", "offset", "filters", "order")
            .rows(1).Font.Bold = True
            .Visible = xlSheetVeryHidden
        End With
    End If
End Function

'Module XQL_Resolve
Option Explicit

' 선택 셀의 충돌 표시 제거 + 코멘트 제거 (서버값 이미 반영된 셀에 사용)
Public Sub XQL_AcceptServerForSelection()
    Dim c As Range
    For Each c In Selection.Cells
        c.Interior.ColorIndex = xlNone
        On Error Resume Next
        If Not c.Comment Is Nothing Then c.Comment.Delete
        On Error GoTo 0
    Next c
End Sub

' 선택된 행의 로컬 값을 강제 푸시(낙관잠금: base_row_version는 숨김열 값)
Public Sub XQL_PushLocalForRow()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim r&: r = Selection.row
    Call XQL_UpsertRows(ws, Array(r))
End Sub

'Module XQL_Ribbon
' === Module: XQL_Ribbon ===
Option Explicit
Dim gRibbon As IRibbonUI

' 리본 로드 콜백
Public Sub XQL_Ribbon_OnLoad(rib As IRibbonUI)
    Set gRibbon = rib
End Sub

' 대시보드 열기
Public Sub XQL_Ribbon_ShowDashboard(control As IRibbonControl)
    XQL_ShowDashboard
End Sub

' 단축 실행(원하면 더 추가)
Public Sub XQL_Ribbon_PullNow(control As IRibbonControl): XQL_Pull_Now: End Sub
Public Sub XQL_Ribbon_OutboxOpen(control As IRibbonControl): XQL_Outbox_Open: End Sub
Public Sub XQL_Ribbon_Query(control As IRibbonControl): XQL_Query_Setup: End Sub
Public Sub XQL_Ribbon_Conflicts(control As IRibbonControl): XQL_Resolve_ScanConflicts: End Sub
Public Sub XQL_Ribbon_Integrity(control As IRibbonControl): XQL_Check_RunActive: End Sub

'Module XQL_SchemaSync
Option Explicit

' === 공개 엔트리포인트 ============================================

' 활성 시트 스키마 동기화
Public Sub XQL_Schema_SyncActiveSheet()
    Dim ws As Worksheet: Set ws = ActiveSheet
    If ws.name = "XQLite" Or ws.name = "XQLite_Conflicts" Then
        MsgBox "데이터 시트에서 실행하세요.", vbExclamation: Exit Sub
    End If

    Dim table$, cols() As Variant
    table = ws.name
    cols = ReadHeaderTypes(ws)
    If Not ValidateHeader(cols) Then Exit Sub

    Dim serverDef As Object: Set serverDef = FetchServerDef(table)
    If serverDef Is Nothing Then
        If CreateTableOnServer(table, cols) Then
            MsgBox "테이블 생성 완료: " & table, vbInformation
        End If
    Else
        Dim addList As Object: Set addList = DiffColumns(serverDef, cols)
        If addList.count = 0 Then
            MsgBox "서버 스키마와 일치합니다. 추가할 컬럼 없음.", vbInformation
        Else
            If AddColumnsOnServer(table, addList) Then
                MsgBox "컬럼 " & addList.count & "개 추가 완료.", vbInformation
            End If
        End If
    End If
End Sub

' 모든 데이터 시트 일괄 동기화
Public Sub XQL_Schema_SyncAllSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.name <> "XQLite" And ws.name <> "XQLite_Conflicts" Then
            ws.Activate
            XQL_Schema_SyncActiveSheet
        End If
    Next ws
End Sub

' === 내부 구현 =====================================================

' 1행=헤더, 2행=타입 토큰을 배열로 반환 [{name, type, notNull}]
Private Function ReadHeaderTypes(ws As Worksheet) As Variant
    Dim lc&, c&, name$, ty$, out As Object
    lc = XQL_LastCol(ws)
    Set out = CreateObject("System.Collections.ArrayList")
    For c = 1 To lc
        name = Trim$(CStr(ws.Cells(1, c).Value2))
        If Len(name) = 0 Then GoTo contc
        ty = Trim$(CStr(ws.Cells(2, c).Value2))
        If Len(ty) = 0 Then ty = "TEXT"
        Dim o As Object: Set o = CreateObject("Scripting.Dictionary")
        o("name") = SanitizeName(name)
        o("type") = NormalizeType(ty)
        o("notNull") = False
        out.add o
contc:
    Next c
    ReadHeaderTypes = out.ToArray
End Function

Private Function ValidateHeader(cols As Variant) As Boolean
    If UBound(cols) < 0 Then
        MsgBox "헤더가 비어 있습니다.", vbExclamation: ValidateHeader = False: Exit Function
    End If
    If UCase$(CStr(cols(0)("name"))) <> "ID" Then
        MsgBox "A1 헤더는 반드시 'id'여야 합니다.", vbExclamation: ValidateHeader = False: Exit Function
    End If
    ValidateHeader = True
End Function

' meta에서 schema:<table> JSON을 가져옴
Private Function FetchServerDef(ByVal table As String) As Object
    Dim wr$, q$, vars$, resp$, parsed As Object, rows As Object
    wr = "key='schema:" & EscapeSqlStr(table) & "'"
    q = "query($l:Int){rows(table:""meta"",whereRaw:""" & wr & """,limit:$l){rows}}"
    vars = "{""l"":1}"
    resp = XQL_GraphQL(q, vars)
    If XQL_HasErrors(resp) Then Exit Function

    Set parsed = JsonConverter.ParseJson(resp)
    Set rows = parsed("data")("rows")("rows")
    If rows.count = 0 Then Exit Function

    Dim v$: v = rows(1)("value")
    On Error Resume Next
    Set FetchServerDef = JsonConverter.ParseJson(CStr(v))
    On Error GoTo 0
End Function

' 서버 정의 vs 엑셀 정의 비교 → 추가해야 할 컬럼 목록(Dictionary[])
Private Function DiffColumns(serverDef As Object, cols As Variant) As Object
    Dim have As Object: Set have = CreateObject("Scripting.Dictionary")
    Dim i&
    For i = 0 To serverDef("columns").count - 1
        have(LCase$(CStr(serverDef("columns")(i)("name")))) = True
    Next i

    Dim out As Object: Set out = CreateObject("System.Collections.ArrayList")
    For i = 1 To UBound(cols) ' 0=id는 빼고 비교
        Dim nm$: nm = LCase$(CStr(cols(i)("name")))
        If Not have.Exists(nm) Then out.add cols(i)
    Next i
    Set DiffColumns = out
End Function

Private Function CreateTableOnServer(ByVal table As String, cols As Variant) As Boolean
    Dim list$, i&
    list = "["
    For i = 1 To UBound(cols) ' id 제외
        list = list & ColumnDefJson(cols(i))
        If i < UBound(cols) Then list = list & ","
    Next i
    list = list & "]"

    Dim q$, vars$, resp$
    q = "mutation($t:String!,$cols:[ColumnDefIn!]!){createTable(table:$t,columns:$cols)}"
    vars = "{""t"":""" & table & """,""cols"":" & list & "}"
    resp = XQL_GraphQL(q, vars)
    If XQL_HasErrors(resp) Then
        MsgBox "createTable 실패: " & resp, vbExclamation
        CreateTableOnServer = False
    Else
        CreateTableOnServer = True
    End If
End Function

Private Function AddColumnsOnServer(ByVal table As String, list As Object) As Boolean
    Dim arr$, i&
    arr = "["
    For i = 0 To list.count - 1
        arr = arr & ColumnDefJson(list(i))
        If i < list.count - 1 Then arr = arr & ","
    Next i
    arr = arr & "]"

    Dim q$, vars$, resp$
    q = "mutation($t:String!,$cols:[ColumnDefIn!]!){addColumns(table:$t,columns:$cols)}"
    vars = "{""t"":""" & table & """,""cols"":" & arr & "}"
    resp = XQL_GraphQL(q, vars)
    If XQL_HasErrors(resp) Then
        MsgBox "addColumns 실패: " & resp, vbExclamation
        AddColumnsOnServer = False
    Else
        AddColumnsOnServer = True
    End If
End Function

' === 헬퍼 ===========================================================

Private Function ColumnDefJson(col As Object) As String
    Dim name$, ty$
    name = CStr(col("name"))
    ty = CStr(col("type"))
    ColumnDefJson = "{""name"":""" & XQL_EscapeJson(name) & """,""type"":""" & XQL_EscapeJson(ty) & """}"
End Function

Private Function NormalizeType(ByVal t As String) As String
    Dim u$: u = UCase$(Trim$(t))
    If InStr(u, "INT") > 0 Then NormalizeType = "INTEGER": Exit Function
    If u = "REAL" Or u = "FLOAT" Or u = "DOUBLE" Then NormalizeType = "REAL": Exit Function
    If u = "BOOLEAN" Or u = "BOOL" Then NormalizeType = "BOOLEAN": Exit Function
    If u = "BLOB" Then NormalizeType = "BLOB": Exit Function
    NormalizeType = "TEXT"
End Function

Private Function SanitizeName(ByVal s As String) As String
    Dim i&, ch$, out$: out = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z0-9_]" Then
            out = out & ch
        ElseIf ch = " " Or ch = "-" Then
            out = out & "_"  ' 빈칸/하이픈 → _
        End If
    Next
    If Len(out) = 0 Then out = "col"
    SanitizeName = out
End Function

Private Function EscapeSqlStr(ByVal s As String) As String
    EscapeSqlStr = Replace(s, "'", "''")
End Function

'Module XQL_Sync
Option Explicit

' 2초 디바운스 후 호출됨 (ThisWorkbook에서 예약)
Public Sub XQL_FlushDirty()
    On Error GoTo eh

    Dim items As Collection
    Set items = ThisWorkbook.XQL_PopDirty()
    If items.count = 0 Then Exit Sub

    ' 시트별로 묶어서 업서트
    Dim i As Long, key As String, p As Long, sheetName As String, rowIdx As Long
    Dim groups As Object: Set groups = CreateObject("Scripting.Dictionary")  ' sheet -> ArrayList(rows)

    For i = 1 To items.count
        key = CStr(items(i))                ' e.g. "items!12"
        p = InStr(1, key, "!")
        If p <= 0 Then GoTo nexti
        sheetName = Left$(key, p - 1)
        rowIdx = CLng(Mid$(key, p + 1))

        If Not groups.Exists(sheetName) Then groups.add sheetName, CreateObject("System.Collections.ArrayList")
        groups(sheetName).add rowIdx
nexti:
    Next

    Dim ws As Worksheet, arr As Variant, keyVar As Variant
    For Each keyVar In groups.keys
        On Error Resume Next
        Set ws = Nothing
        Set ws = Sheets(CStr(keyVar))
        On Error GoTo eh
        If Not ws Is Nothing Then
            arr = groups(CStr(keyVar)).ToArray()
            XQL_UpsertRows ws, arr
        End If
    Next keyVar

    Exit Sub
eh:
    MsgBox "FlushDirty error: " & Err.Description, vbExclamation
End Sub

' 시트의 특정 행들 업서트
Public Sub XQL_UpsertRows(ws As Worksheet, rowsArr As Variant)
    On Error GoTo eh
    If IsEmpty(rowsArr) Then Exit Sub
    If UBound(rowsArr) - LBound(rowsArr) + 1 <= 0 Then Exit Sub

    XQL_EnsureAuxCols ws

    Dim lc As Long, vcol As Long, dcol As Long
    lc = XQL_LastCol(ws): vcol = XQL_VerCol(ws): dcol = XQL_DirtyCol(ws)

    ' UpsertRowInput[] 생성 (안전한 콤마 처리)
    Dim buf As String: buf = "["
    Dim firstRow As Boolean: firstRow = True

    Dim i As Long, r As Long, idVal, basev As Long
    For i = LBound(rowsArr) To UBound(rowsArr)
        r = CLng(rowsArr(i))
        If r < 2 Then GoTo nexti

        idVal = ws.Cells(r, 1).Value2
        If Len(CStr(idVal)) = 0 Then GoTo nexti

        On Error Resume Next
        basev = CLng(ws.Cells(r, vcol).Value2)
        On Error GoTo 0

        ' data 오브젝트
        Dim data As String: data = ""
        Dim firstCol As Boolean: firstCol = True

        Dim c As Long, k As String, v
        For c = 1 To lc
            k = CStr(ws.Cells(1, c).Value2)  ' 헤더=컬럼명
            If Len(k) = 0 Then GoTo contc
            ' 메타 컬럼 제외: id, _ver, _dirty (컬럼 인덱스로 필터)
            If c = 1 Or c = vcol Or c = dcol Then GoTo contc

            v = ws.Cells(r, c).Value
            Dim pair As String: pair = ""

            If IsEmpty(v) Or v = "" Then
                pair = """" & XQL_EscapeJson(k) & """:null"
            ElseIf VarType(v) = vbBoolean Then
                pair = """" & XQL_EscapeJson(k) & """:" & IIf(CBool(v), "true", "false")
            ElseIf IsNumeric(v) And Not IsDate(v) Then
                pair = """" & XQL_EscapeJson(k) & """:" & CStr(v)
            Else
                pair = """" & XQL_EscapeJson(k) & """:""" & XQL_EscapeJson(CStr(v)) & """"
            End If

            If Len(pair) > 0 Then
                If Not firstCol Then data = data & ","
                data = data & pair
                firstCol = False
            End If
contc:
        Next c

        If Not firstRow Then buf = buf & ","
        buf = buf & "{""id"":" & CLng(idVal) & ",""base_row_version"":" & basev & ",""data"":{" & data & "}}"
        firstRow = False
nexti:
    Next i
    buf = buf & "]"

    Dim q As String, payload As String, resp As String
    q = "mutation($t:String!,$rows:[UpsertRowInput!]!,$a:String!){upsertRows(table:$t,rows:$rows,actor:$a){max_row_version affected conflicts errors}}"
    payload = "{""t"":""" & ws.name & """,""rows"":" & buf & ",""a"":""" & XQL_GetNickname() & """}"

    resp = XQL_GraphQL(q, payload)

    ' 결과 반영: 버전 갱신 + dirty 해제
    Dim mx As Long: mx = XQL_ExtractMaxRowVersion(resp)
    If mx > 0 Then
        For i = LBound(rowsArr) To UBound(rowsArr)
            r = CLng(rowsArr(i))
            If r >= 2 Then
                ws.Cells(r, vcol).Value = mx
                XQL_ClearDirty ws, r
            End If
        Next
        Sheets("XQLite").Range("A7").Offset(0, 1).Value = mx ' LAST_MAX_ROW_VERSION
    End If

    ' 충돌 단순 감지(상세 처리는 Conflict Resolver 사용)
    If InStr(1, resp, """conflicts"":", vbTextCompare) > 0 Then
        For i = LBound(rowsArr) To UBound(rowsArr)
            r = CLng(rowsArr(i))
            If r >= 2 Then ws.rows(r).Interior.Color = RGB(255, 255, 200) ' 노랑
        Next
        MsgBox "서버와 충돌(conflicts)이 감지되었습니다. Resolve 화면에서 처리하세요.", vbInformation
    End If

    Exit Sub
eh:
    MsgBox "UpsertRows error: " & Err.Description, vbExclamation
End Sub

' 수동 커밋 버튼(현재 시트의 모든 dirty 행 강제 업서트)
Public Sub XQL_CommitNow()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long, r As Long
    Dim list As Object: Set list = CreateObject("System.Collections.ArrayList")

    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        If Len(CStr(ws.Cells(r, XQL_DirtyCol(ws)).Value2)) > 0 Then list.add r
    Next

    If list.count = 0 Then
        MsgBox "업서트할 변경이 없습니다.", vbInformation
        Exit Sub
    End If

    XQL_UpsertRows ws, list.ToArray()
End Sub

' 원클릭 복구: 엑셀 → 서버 DB 재작성(스키마 동일 전제)
Public Sub XQL_RecoverServer()
    On Error GoTo eh
    Dim ws As Worksheet: Set ws = ActiveSheet
    XQL_EnsureAuxCols ws

    Dim lc As Long, lastRow As Long, r As Long, c As Long, k As String
    Dim rowsJson As String, q As String, payload As String, resp As String
    lc = XQL_LastCol(ws)
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row

    rowsJson = "["
    Dim firstRow As Boolean: firstRow = True

    For r = 2 To lastRow
        If Len(CStr(ws.Cells(r, 1).Value2)) = 0 Then GoTo nextr

        If Not firstRow Then rowsJson = rowsJson & ","
        rowsJson = rowsJson & "{"

        Dim firstCol As Boolean: firstCol = True
        For c = 1 To lc
            k = CStr(ws.Cells(1, c).Value2)
            If Len(k) = 0 Then GoTo contc
            Dim v: v = ws.Cells(r, c).Value

            Dim pair As String
            If IsNumeric(v) And Not IsDate(v) And Len(CStr(v)) > 0 Then
                pair = """" & XQL_EscapeJson(k) & """:" & CStr(v)
            Else
                pair = """" & XQL_EscapeJson(k) & """:""" & XQL_EscapeJson(CStr(v)) & """"
            End If

            If Not firstCol Then rowsJson = rowsJson & ","
            rowsJson = rowsJson & pair
            firstCol = False
contc:
        Next c
        rowsJson = rowsJson & "}"
        firstRow = False
nextr:
    Next r
    rowsJson = rowsJson & "]"

    ' 최신 스키마 해시 조회
    Dim metaQ As String, metaResp As String, schemaHash As String
    metaQ = "query{ meta{ schema_hash max_row_version } }"
    metaResp = XQL_GraphQL(metaQ, "{}")
    schemaHash = XQL_ExtractBetween(metaResp, "schema_hash"":""", """")

    q = "mutation($t:String!,$rows:[JSON!]!,$h:String!,$a:String!){recoverFromExcel(table:$t,rows:$rows,schema_hash:$h,actor:$a)}"
    payload = "{""t"":""" & ws.name & """,""rows"":" & rowsJson & ",""h"":""" & schemaHash & """,""a"":""" & XQL_GetNickname() & """}"
    resp = XQL_GraphQL(q, payload)

    MsgBox "서버 복구 요청 완료", vbInformation
    Exit Sub
eh:
    MsgBox "Recover error: " & Err.Description, vbExclamation
End Sub

Private Function XQL_ExtractBetween(ByVal s As String, ByVal a As String, ByVal b As String) As String
    Dim p As Long, q As Long
    p = InStr(1, s, a, vbTextCompare): If p = 0 Then Exit Function
    p = p + Len(a)
    q = InStr(p, s, b, vbTextCompare): If q = 0 Then q = Len(s) + 1
    XQL_ExtractBetween = Mid$(s, p, q - p)
End Function

'Module XQL_Template
' === Module: XQL_Template ===
Option Explicit

Private Type ColSpec
    name As String
    ty As String
End Type

' 엔트리: 새 시트 생성 + 스키마 동기화(있으면)
Public Sub XQL_NewTableWizard()
    Dim t$, spec$
    t = InputBox("테이블명(=시트명)을 입력하세요:", "XQLite", "items")
    If Len(t) = 0 Then Exit Sub
    spec = InputBox("컬럼과 타입을 입력하세요 (예: name:TEXT,rarity:INTEGER,atk_lv50:REAL,owned:BOOLEAN)" & vbCrLf & _
                    "타입 생략 시 TEXT로 처리됩니다. 'id'는 자동 추가됩니다.", _
                    "XQLite", "name,rarity:INTEGER,atk_lv50:REAL,owned:BOOLEAN")
    If Len(spec) = 0 Then Exit Sub

    Dim cols() As ColSpec
    cols = ParseSpec(spec)

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets(t)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = Sheets.add(After:=Sheets(Sheets.count))
        ws.name = t
    Else
        ws.Cells.Clear
    End If

    SetupHeaders ws, cols
    ApplyColumnValidation ws
    FormatSheet ws
    XQL_EnsureAuxCols ws     ' _ver/_dirty 생성·숨김

    ' 선택: 바로 서버 스키마 동기화
    On Error Resume Next
    XQL_Schema_SyncActiveSheet
    On Error GoTo 0

    ws.Activate
    MsgBox "시트 템플릿이 준비되었습니다." & vbCrLf & _
           "- 1행: 컬럼명, 2행: 타입" & vbCrLf & _
           "- A열 id는 정수 키" & vbCrLf & _
           "- _ver/_dirty 보조열 자동 생성(숨김)", vbInformation
End Sub

' 현재 시트(헤더/타입 기반)만 재검증(서식 유지)
Public Sub XQL_ReapplyValidationForActive()
    ApplyColumnValidation ActiveSheet
    MsgBox "데이터 유효성 규칙을 갱신했습니다.", vbInformation
End Sub

' ==== 내부 구현 ====

Private Function ParseSpec(ByVal s As String) As ColSpec()
    Dim parts() As String, i&, p$, name$, ty$
    Dim tmp() As ColSpec
    s = Trim$(s)
    If Len(s) = 0 Then ReDim tmp(0 To -1): ParseSpec = tmp: Exit Function

    parts = Split(s, ",")
    ReDim tmp(0 To UBound(parts))
    For i = 0 To UBound(parts)
        p = Trim$(parts(i))
        If InStr(p, ":") > 0 Then
            name = Trim$(Left$(p, InStr(p, ":") - 1))
            ty = Trim$(Mid$(p, InStr(p, ":") + 1))
        Else
            name = p: ty = "TEXT"
        End If
        If Len(name) = 0 Then name = "col" & CStr(i + 1)
        tmp(i).name = SanitizeName(name)
        tmp(i).ty = NormalizeType(ty)
    Next
    ParseSpec = tmp
End Function

Private Sub SetupHeaders(ws As Worksheet, cols() As ColSpec)
    Dim c&, i&
    ' A열 id 고정
    ws.Cells(1, 1).Value = "id"
    ws.Cells(2, 1).Value = "INTEGER"
    c = 2
    For i = LBound(cols) To UBound(cols)
        ws.Cells(1, c).Value = cols(i).name
        ws.Cells(2, c).Value = cols(i).ty
        c = c + 1
    Next
End Sub

Private Sub ApplyColumnValidation(ws As Worksheet)
    Dim lastCol&, c&, ty$
    lastCol = XQL_LastCol(ws)

    ' id 컬럼: 정수
    With ws.Range(ws.Cells(3, 1), ws.Cells(rows.count, 1))
        .Validation.Delete
        .Validation.add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="-2147483648", Formula2:="2147483647"
        .NumberFormat = "0"
    End With

    For c = 2 To lastCol
        ty = UCase$(CStr(ws.Cells(2, c).Value2))
        With ws.Range(ws.Cells(3, c), ws.Cells(rows.count, c))
            .Validation.Delete
            Select Case True
                Case InStr(ty, "INT") > 0
                    .Validation.add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                        Operator:=xlBetween, Formula1:="-9007199254740991", Formula2:="9007199254740991"
                    .NumberFormat = "0"
                Case ty = "REAL" Or ty = "FLOAT" Or ty = "DOUBLE"
                    .Validation.add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, _
                        Operator:=xlBetween, Formula1:="-1.0E+308", Formula2:="1.0E+308"
                    .NumberFormat = "General"
                Case ty = "BOOLEAN" Or ty = "BOOL"
                    .Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                        Formula1:="TRUE,FALSE"
                Case Else
                    ' TEXT/BLOB: 제한 없음
            End Select
        End With
    Next c
End Sub

Private Sub FormatSheet(ws As Worksheet)
    ' 머리행 디자인/고정
    ws.rows(1).Font.Bold = True
    ws.rows(2).Font.Italic = True
    ws.rows(2).Interior.Color = RGB(245, 245, 245)
    ws.Columns.AutoFit
    ' 보기 편하게 1~2행 고정
    ws.Activate
    ActiveWindow.SplitColumn = 0
    ActiveWindow.SplitRow = 2
    ActiveWindow.FreezePanes = True
End Sub

Private Function NormalizeType(ByVal t As String) As String
    Dim u$: u = UCase$(Trim$(t))
    If InStr(u, "INT") > 0 Then NormalizeType = "INTEGER": Exit Function
    If u = "REAL" Or u = "FLOAT" Or u = "DOUBLE" Then NormalizeType = "REAL": Exit Function
    If u = "BOOLEAN" Or u = "BOOL" Then NormalizeType = "BOOLEAN": Exit Function
    If u = "BLOB" Then NormalizeType = "BLOB": Exit Function
    NormalizeType = "TEXT"
End Function

Private Function SanitizeName(ByVal s As String) As String
    Dim i&, ch$, out$: out = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z0-9_]" Then
            out = out & ch
        ElseIf ch = " " Or ch = "-" Then
            out = out & "_"
        End If
    Next
    If Len(out) = 0 Then out = "col"
    SanitizeName = out
End Function

'Module XQL_Validate
Option Explicit

' Target 범위 내에서 타입 위반 셀을 찾아 표시하고 false 반환
Public Function XQL_ValidateRange(ws As Worksheet, rng As Range) As Boolean
    On Error GoTo bad
    Dim ok As Boolean: ok = True
    Dim r As Range, c As Range, tkn$
    Dim lastCol&: lastCol = XQL_LastCol(ws)
    For Each r In rng.rows
        If r.row < 2 Then GoTo nxr
        For Each c In r.Cells
            If c.Column > lastCol Then Exit For
            tkn = CStr(ws.Cells(2, c.Column).Value2)
            If Len(tkn) = 0 Then GoTo nxc
            If Not XQL_ValidateCell(c, tkn) Then
                c.Interior.Color = RGB(255, 230, 230) ' 빨강빛
                ok = False
            Else
                If c.row Mod 2 = 0 Then
                    c.Interior.ColorIndex = xlNone
                Else
                    c.Interior.ColorIndex = xlNone
                End If
            End If
nxc:
        Next c
nxr:
    Next r
    XQL_ValidateRange = ok
    Exit Function
bad:
    XQL_ValidateRange = True ' 검증 실패 시 막진 않음
End Function

Private Function XQL_ValidateCell(c As Range, ByVal tkn As String) As Boolean
    Dim v: v = c.Value
    Select Case UCase$(tkn)
        Case "INTEGER": If Len(v) = 0 Then XQL_ValidateCell = True Else XQL_ValidateCell = (IsNumeric(v) And v = CLng(v))
        Case "REAL":    If Len(v) = 0 Then XQL_ValidateCell = True Else XQL_ValidateCell = (IsNumeric(v))
        Case "BOOLEAN": If Len(v) = 0 Then XQL_ValidateCell = True Else XQL_ValidateCell = (UCase$(CStr(v)) = "TRUE" Or UCase$(CStr(v)) = "FALSE" Or v = 0 Or v = 1)
        Case Else:      XQL_ValidateCell = True ' TEXT or unknown
    End Select
End Function

'UserForm frmXQLite
Option Explicit

' ======== 초기화/상태 ========
Private Sub UserForm_Initialize()
    On Error Resume Next
    Me.Caption = "XQLite Dashboard"
    LoadSettingsToForm
    Me.lblStatus.Caption = "Ready."
End Sub

Private Sub LoadSettingsToForm()
    On Error Resume Next
    Me.txtNickname.Text = XQL_GetNickname()
    Me.txtServerUrl.Text = GetServerUrlGuess() ' 여러 키 시도

    Me.chkAutoLock.Value = GetKVb("AUTO_LOCK_ON_SELECT", True)
    Me.chkAutoRelease.Value = GetKVb("AUTO_RELEASE_ON_MOVE", True)
    Me.txtLockRefresh.Text = GetKVs("LOCK_REFRESH_SEC", "5")

    Me.chkPullEnabled.Value = GetKVb("PULL_ENABLED", True)
    Me.txtPullSec.Text = GetKVs("PULL_SEC", "10")

    Me.txtOutboxRetry.Text = GetKVs("OUTBOX_RETRY_SEC", "15")
    Me.txtOutboxMax.Text = GetKVs("OUTBOX_MAX_RETRY", "10")

    Me.chkPermsAuto.Value = GetKVb("PERMS_AUTO_APPLY", True)
End Sub

' ======== 저장 ========
Private Sub btnSave_Click()
    On Error GoTo eh
    SetKV "NICKNAME", Trim$(Me.txtNickname.Text)
    ' 서버 URL은 존재하는 키에 맞춰 저장 시도
    SaveServerUrlTrim Trim$(Me.txtServerUrl.Text)

    SetKV "AUTO_LOCK_ON_SELECT", IIf(Me.chkAutoLock.Value, "TRUE", "FALSE")
    SetKV "AUTO_RELEASE_ON_MOVE", IIf(Me.chkAutoRelease.Value, "TRUE", "FALSE")
    SetKV "LOCK_REFRESH_SEC", NzStr(Me.txtLockRefresh.Text, "5")

    SetKV "PULL_ENABLED", IIf(Me.chkPullEnabled.Value, "TRUE", "FALSE")
    SetKV "PULL_SEC", NzStr(Me.txtPullSec.Text, "10")

    SetKV "OUTBOX_RETRY_SEC", NzStr(Me.txtOutboxRetry.Text, "15")
    SetKV "OUTBOX_MAX_RETRY", NzStr(Me.txtOutboxMax.Text, "10")

    SetKV "PERMS_AUTO_APPLY", IIf(Me.chkPermsAuto.Value, "TRUE", "FALSE")

    Me.lblStatus.Caption = "Saved."
    Exit Sub
eh:
    Me.lblStatus.Caption = "Save failed: " & Err.Description
End Sub

' ======== 액션 버튼 ========
Private Sub btnPullNow_Click()
    On Error Resume Next
    XQL_Pull_Now
    Me.lblStatus.Caption = "Pulled."
End Sub

Private Sub btnPullStart_Click()
    On Error Resume Next
    XQL_Pull_Start
    Me.lblStatus.Caption = "Pull started."
End Sub

Private Sub btnPullStop_Click()
    On Error Resume Next
    XQL_Pull_Stop
    Me.lblStatus.Caption = "Pull stopped."
End Sub

Private Sub btnOutboxOpen_Click()
    On Error Resume Next
    XQL_Outbox_Open
    Me.lblStatus.Caption = "Outbox opened."
End Sub

Private Sub btnQuery_Click()
    On Error Resume Next
    XQL_Query_Setup
    Me.lblStatus.Caption = "Query panel ready."
End Sub

Private Sub btnConflicts_Click()
    On Error Resume Next
    XQL_Resolve_ScanConflicts
    Me.lblStatus.Caption = "Conflicts scanned."
End Sub

Private Sub btnEnums_Click()
    On Error Resume Next
    XQL_Enums_RefreshAll
    Me.lblStatus.Caption = "Enums refreshed."
End Sub

Private Sub btnAudit_Click()
    On Error Resume Next
    XQL_Audit_Setup
    XQL_Audit_Run
    Me.lblStatus.Caption = "Audit loaded."
End Sub

Private Sub btnIntegrity_Click()
    On Error Resume Next
    XQL_Check_RunActive
    Me.lblStatus.Caption = "Integrity report done."
End Sub

Private Sub btnFullPull_Click()
    On Error Resume Next
    XQL_Check_FullPullActive
    Me.lblStatus.Caption = "Full Pull complete."
End Sub

Private Sub btnFullPush_Click()
    On Error Resume Next
    XQL_Check_FullPushActive
    Me.lblStatus.Caption = "Full Push started."
End Sub

Private Sub btnPermsApply_Click()
    On Error Resume Next
    XQL_Perms_ApplyAll
    Me.lblStatus.Caption = "Permissions applied."
End Sub

' ======== KV/설정 유틸 ========
Private Function GetServerUrlGuess() As String
    Dim v As String
    v = GetKVs("GRAPHQL_URL", "")
    If Len(v) = 0 Then v = GetKVs("SERVER_URL", "")
    If Len(v) = 0 Then v = GetKVs("HOST", "")
    GetServerUrlGuess = v
End Function

Private Sub SaveServerUrlTrim(ByVal url As String)
    If Len(url) = 0 Then Exit Sub
    If Right$(url, 1) = "/" Then url = Left$(url, Len(url) - 1)
    ' 기존 키 중 존재하는 것으로 우선 저장
    If HasKey("GRAPHQL_URL") Then
        SetKV "GRAPHQL_URL", url: Exit Sub
    ElseIf HasKey("SERVER_URL") Then
        SetKV "SERVER_URL", url: Exit Sub
    ElseIf HasKey("HOST") Then
        SetKV "HOST", url: Exit Sub
    End If
    ' 없으면 GRAPHQL_URL로 신규 추가
    SetKV "GRAPHQL_URL", url
End Sub

Private Function HasKey(ByVal key As String) As Boolean
    Dim cfg As Worksheet: Set cfg = Sheets("XQLite")
    Dim last&, r&
    last = cfg.Cells(cfg.rows.count, 1).End(xlUp).row
    For r = 1 To last
        If LCase$(CStr(cfg.Cells(r, 1).Value2)) = LCase$(key) Then HasKey = True: Exit Function
    Next r
End Function

Private Function GetKVs(ByVal key As String, ByVal defVal As String) As String
    On Error Resume Next
    Dim cfg As Worksheet: Set cfg = Sheets("XQLite")
    Dim last&, r&
    last = cfg.Cells(cfg.rows.count, 1).End(xlUp).row
    For r = 1 To last
        If LCase$(CStr(cfg.Cells(r, 1).Value2)) = LCase$(key) Then
            GetKVs = CStr(cfg.Cells(r, 2).Value)
            If Len(GetKVs) = 0 Then GetKVs = defVal
            Exit Function
        End If
    Next r
    GetKVs = defVal
End Function

Private Function GetKVb(ByVal key As String, ByVal defVal As Boolean) As Boolean
    GetKVb = (UCase$(GetKVs(key, IIf(defVal, "TRUE", "FALSE"))) = "TRUE")
End Function

Private Sub SetKV(ByVal key As String, ByVal val As String)
    On Error Resume Next
    Dim cfg As Worksheet: Set cfg = Sheets("XQLite")
    Dim last&, r&, found As Boolean
    last = cfg.Cells(cfg.rows.count, 1).End(xlUp).row
    For r = 1 To last
        If LCase$(CStr(cfg.Cells(r, 1).Value2)) = LCase$(key) Then
            cfg.Cells(r, 2).Value = val: found = True: Exit For
        End If
    Next r
    If Not found Then
        cfg.Cells(last + 1, 1).Value = key
        cfg.Cells(last + 1, 2).Value = val
    End If
End Sub

Private Function NzStr(ByVal v As String, ByVal fb As String) As String
    If Len(Trim$(v)) = 0 Then NzStr = fb Else NzStr = Trim$(v)
End Function


