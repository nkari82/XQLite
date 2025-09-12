Option Explicit

Private Const API_URL As String = "http://localhost:4000/"
Private Const DEBOUNCE_MS As Long = 2000

' 숨김열에 버전 저장하는 열 인덱스(예: 마지막 열 + 1)
Private Function VersionCol(ws As Worksheet) As Long
    VersionCol = ws.UsedRange.Columns.Count + 1
End Function

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Row = 1 Then Exit Sub ' 헤더
    Cells(Target.Row, VersionCol(ActiveSheet)).Value = Cells(Target.Row, VersionCol(ActiveSheet)).Value ' 존재보장
    ScheduleDebounced
End Sub

Private Sub ScheduleDebounced()
    Static pending As Boolean
    If pending Then Exit Sub
    pending = True
    Application.OnTime Now + TimeSerial(0, 0, 2), "DoUpsertV2"
    pending = False
End Sub

Sub DoUpsertV2()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow&, lastCol&, r&, c&, vcol&
    lastRow = ws.UsedRange.Rows.Count
    lastCol = ws.UsedRange.Columns.Count
    vcol = VersionCol(ws)

    Dim headers() As String: ReDim headers(1 To lastCol)
    For c = 1 To lastCol: headers(c) = CStr(ws.Cells(1, c).Value2): Next

    Dim arr As String: arr = "["
    For r = 2 To lastRow
        Dim basev&: basev = CLng(0)
        On Error Resume Next
        basev = CLng(ws.Cells(r, vcol).Value2)
        On Error GoTo 0

        arr = arr & "{""id"":" & CLng(ws.Cells(r, 1).Value2) & ",""base_row_version"":" & basev & ",""data"":{"
        For c = 1 To lastCol
            Dim k$, val$
            k = headers(c)
            If Len(k) = 0 Then GoTo nextc
            val = CStr(ws.Cells(r, c).Value2)
            If IsNumeric(val) Then
                arr = arr & """" & k & """:" & val
            Else
                arr = arr & """" & k & """:""" & Replace(val, """", "\""") & """"
            End If
            If c < lastCol Then arr = arr & ","
nextc:
        Next
        arr = arr & "}}"
        If r < lastRow Then arr = arr & ","
    Next
    arr = arr & "]"

    Dim q$, payload$, resp$
    q = "mutation($t:String!,$rs:[UpsertRowInput!]!,$a:String!){upsertRowsV2(table:$t,rows:$rs,actor:$a){max_row_version affected conflicts}}"
    payload = "{""query"":""" & Replace(q, """", "\""") & """,""variables"":{""t"":""" & ws.Name & """,""rs"":" & arr & ",""a"":""excel-user""}}"
    resp = PostJson(payload)

    ' TODO: JSON 파서로 conflicts 처리 (VBA-JSON 등을 사용)
    ' 여기서는 단순 감지: "conflicts":[ 가 있으면 노란색 표시
    If InStr(1, resp, """conflicts"":[", vbTextCompare) > 0 Then
        ' 간단 표시(실사용은 파싱 후 특정 셀 하이라이트)
        ws.Rows("2:" & lastRow).Interior.Color = RGB(255, 255, 200)
    End If

    ' 성공 시 각 행의 base_row_version 갱신(여기선 서버 max만 사용 예시)
    Dim mx&: mx = ExtractMaxRowVersion(resp)
    If mx > 0 Then
        For r = 2 To lastRow
            ws.Cells(r, vcol).Value = mx
        Next
    End If
End Sub

Private Function ExtractMaxRowVersion(ByVal resp As String) As Long
    Dim p&, q&, s$
    p = InStr(resp, "max_row_version"":")
    If p = 0 Then Exit Function
    q = InStr(p + 16, resp, ",")
    s = Mid(resp, p + 16, q - (p + 16))
    ExtractMaxRowVersion = CLng(Val(s))
End Function

Private Function PostJson(ByVal json As String) As String
    Dim xhr As Object: Set xhr = CreateObject("MSXML2.XMLHTTP")
    xhr.Open "POST", API_URL, False
    xhr.setRequestHeader "Content-Type", "application/json"
    xhr.setRequestHeader "x-api-key", "dev-secret-change-me"
    xhr.send json
    PostJson = xhr.responseText
End Function
