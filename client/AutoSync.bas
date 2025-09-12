Option Explicit

' === 설정 ===
Private Const API_URL As String = "http://localhost:4000/"
Private Const DEBOUNCE_MS As Long = 2000
Private Const PULL_SEC As Long = 10
Private Const PRESENCE_MS As Long = 2500

' 상태
Private lastChangeTick As Double
Private lastPullTick As Double
Private nickname As String
Private sinceVersion As Long

' ───── 시작/초기화 ─────
Sub XQLite_Init()
    nickname = Range("B1").Value ' 예: B1에 닉네임
    sinceVersion = 0
    Application.OnTime Now + TimeSerial(0,0,1), "XQLite_BackgroundPump"
End Sub

' ───── 백그라운드 루프 ─────
Sub XQLite_BackgroundPump()
    On Error Resume Next
    ' Presence Heartbeat
    Call SendPresence

    ' 주기 Pull
    If (Timer - lastPullTick) >= PULL_SEC Then
        lastPullTick = Timer
        Call PullIncremental
    End If

    ' 디바운스 업서트
    If lastChangeTick > 0 And (Timer - lastChangeTick) >= (DEBOUNCE_MS / 1000#) Then
        lastChangeTick = 0
        Call DebouncedUpsert
    End If

    Application.OnTime Now + TimeSerial(0,0,1), "XQLite_BackgroundPump"
End Sub

' ───── 변경 감지 ─────
Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, UsedRange) Is Nothing Then Exit Sub
    ' 타입/형식 간단 검증(예: 숫자 컬럼은 숫자인지)
    ' (실사용 시 헤더 타입 매핑표를 사용)
    lastChangeTick = Timer

    ' 셀 락 시도
    Call TryAcquireLock(Target)
End Sub

' ───── Commit 버튼용 ─────
Sub XQLite_Commit()
    Call DebouncedUpsert
End Sub

' ───── Sync 버튼용 ─────
Sub XQLite_Sync()
    sinceVersion = 0
    Call PullFull
End Sub

' ───── Recover 버튼용 ─────
Sub XQLite_Recover()
    Dim rows As String: rows = SheetToJson(ActiveSheet)
    Dim q As String
    q = "mutation($t:String!,$rs:[JSON!]!,$h:String!,$a:String!){recoverFromExcel(table:$t,rows:$rs,schema_hash:$h,actor:$a)}"
    Call PostGraphQL(q, "{""t"":""" & ActiveSheet.Name & """,""rs"":" & rows & ",""h"":""demo"",""a"":""" & nickname & """}")
End Sub

' ───── Presence ─────
Private Sub SendPresence()
    Dim sel As String: sel = ActiveCell.Address(False, False)
    Dim q As String
    q = "mutation($n:String!,$s:String,$c:String){presenceHeartbeat(nickname:$n,sheet:$s,cell:$c)}"
    Call PostGraphQL(q, "{""n"":""" & nickname & """,""s"":""" & ActiveSheet.Name & """,""c"":""" & sel & """}")
End Sub

' ───── 락 ─────
Private Sub TryAcquireLock(ByVal Target As Range)
    Dim cellAddr As String: cellAddr = Target(1,1).Address(False, False)
    Dim q As String
    q = "mutation($s:String!,$c:String!,$n:String!){acquireLock(sheet:$s,cell:$c,nickname:$n)}"
    Dim r As String
    r = PostGraphQL(q, "{""s"":""" & ActiveSheet.Name & """,""c"":""" & cellAddr & """,""n"":""" & nickname & """}")
    ' 실패 시 사용자 알림/되돌리기 등 처리
End Sub

' ───── 디바운스 업서트 ─────
Private Sub DebouncedUpsert()
    Dim rows As String: rows = SheetToJson(ActiveSheet)
    Dim q As String
    q = "mutation($t:String!,$rs:[JSON!]!,$a:String!){upsertRows(table:$t,rows:$rs,actor:$a){max_row_version affected errors}}"
    Dim resp As String
    resp = PostGraphQL(q, "{""t"":""" & ActiveSheet.Name & """,""rs"":" & rows & ",""a"":""" & nickname & """}")
    ' 응답에서 max_row_version 파싱하여 sinceVersion 갱신
    ' (간단히 전체 싱크를 소규모로는 다시 가져와도 됨)
End Sub

' ───── Pull (증분) ─────
Private Sub PullIncremental()
    Dim q As String
    q = "query($t:String!,$sv:Int){rows(table:$t,since_version:$sv){rows max_row_version}}"
    Dim resp As String
    resp = PostGraphQL(q, "{""t"":""" & ActiveSheet.Name & """,""sv"":" & CStr(sinceVersion) & "}")
    ' rows 적용 & sinceVersion 갱신
End Sub

' ───── Pull (풀) ─────
Private Sub PullFull()
    Dim q As String
    q = "query($t:String!){rows(table:$t){rows max_row_version}}"
    Dim resp As String
    resp = PostGraphQL(q, "{""t"":""" & ActiveSheet.Name & """}")
    ' 시트 전체 리렌더 + sinceVersion 갱신
End Sub

' ───── 유틸: 시트→JSON (헤더=컬럼) ─────
Private Function SheetToJson(ws As Worksheet) As String
    Dim lastRow&, lastCol&, r&, c&
    lastRow = ws.UsedRange.Rows.Count
    lastCol = ws.UsedRange.Columns.Count
    Dim headers() As String
    ReDim headers(1 To lastCol)
    For c = 1 To lastCol
        headers(c) = CStr(ws.Cells(1, c).Value2)
    Next
    Dim sb As String: sb = "["
    For r = 2 To lastRow
        sb = sb & "{"
        For c = 1 To lastCol
            Dim key$, val$
            key = headers(c)
            val = Replace(CStr(ws.Cells(r, c).Value2), """", "\""")
            If key = "id" Or IsNumeric(val) Then
                sb = sb & """" & key & """:" & IIf(val="", "0", val)
            Else
                sb = sb & """" & key & """:""" & val & """"
            End If
            If c < lastCol Then sb = sb & ","
        Next
        sb = sb & "}"
        If r < lastRow Then sb = sb & ","
    Next
    sb = sb & "]"
    SheetToJson = sb
End Function

' ───── 유틸: GraphQL POST (간소화) ─────
Private Function PostGraphQL(query As String, variablesJson As String) As String
    Dim xhr As Object: Set xhr = CreateObject("MSXML2.XMLHTTP")
    Dim payload As String
    payload = "{""query"":""" & Replace(query, """", "\""") & """,""variables"":" & variablesJson & "}"
    xhr.Open "POST", API_URL, False
    xhr.setRequestHeader "Content-Type", "application/json"
    xhr.send payload
    PostGraphQL = xhr.responseText
End Function
