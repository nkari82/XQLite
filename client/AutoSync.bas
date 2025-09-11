Option Explicit

Private gBASE As String, gAPI As String
Private gPullSec As Double, gPushDebounceSec As Double
Private gNick As String, gPresenceSec As Double
Private gNextPullAt As Date, gNextPushAt As Date, gNextPresenceAt As Date
Private gPullScheduled As Boolean, gPushScheduled As Boolean, gPresenceScheduled As Boolean
Private gRunning As Boolean

Private Const META_ID As String = "id"
Private Const META_VER As String = "row_version"
Private Const META_UPD As String = "updated_at"
Private Const META_DEL As String = "deleted"
Private Const HIDDEN_CONFLICT_SHEET As String = "_Conflicts"

Public Sub AutoSync_Init(ByVal baseUrl As String, ByVal apiKey As String, _
                         ByVal pullIntervalSec As Double, ByVal pushDebounceSec As Double, _
                         Optional ByVal serverWins As Boolean = True, _
                         Optional ByVal nick As String = "", Optional ByVal presenceSec As Double = 3)
    gBASE = baseUrl: gAPI = apiKey
    gPullSec = pullIntervalSec: gPushDebounceSec = pushDebounceSec
    gPresenceSec = presenceSec
    GQL_BASE = gBASE & "/graphql": GQL_API = gAPI

    If Len(nick) = 0 Then gNick = GetOrPromptNickname() Else gNick = nick: SaveNickname gNick
    EnsureHiddenSheet HIDDEN_CONFLICT_SHEET
End Sub

Public Sub AutoSync_Start()
    If gRunning Then Exit Sub
    gRunning = True

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Not IsHiddenSheet(ws) Then
            EnsureSheetSchema ws
            PullOnce ws, 0
        End If
    Next ws

    SchedulePull: SchedulePresence
End Sub

Public Sub AutoSync_Stop()
    gRunning = False
    On Error Resume Next
    If gPullScheduled Then Application.OnTime gNextPullAt, "AutoSync_PullTick", , False
    If gPushScheduled Then Application.OnTime gNextPushAt, "AutoSync_PushTick", , False
    If gPresenceScheduled Then Application.OnTime gNextPresenceAt, "AutoSync_PresenceTick", , False
End Sub

Private Sub SchedulePull()
    If Not gRunning Or gPullScheduled Then Exit Sub
    gNextPullAt = Now + TimeSerial(0, 0, gPullSec)
    Application.OnTime gNextPullAt, "AutoSync_PullTick": gPullScheduled = True
End Sub
Public Sub AutoSync_PullTick()
    gPullScheduled = False
    If Not gRunning Then Exit Sub
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Not IsHiddenSheet(ws) Then PullOnce ws, SheetMaxVersion(ws)
    Next ws
    SchedulePull
End Sub

Private Sub SchedulePush()
    If Not gRunning Then Exit Sub
    gNextPushAt = Now + TimeSerial(0, 0, gPushDebounceSec)
    If Not gPushScheduled Then
        Application.OnTime gNextPushAt, "AutoSync_PushTick": gPushScheduled = True
    End If
End Sub
Public Sub AutoSync_PushTick()
    gPushScheduled = False
    If Not gRunning Then Exit Sub
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Not IsHiddenSheet(ws) Then PushDirty ws
    Next ws
End Sub

Private Sub SchedulePresence()
    If Not gRunning Or gPresenceScheduled Then Exit Sub
    gNextPresenceAt = Now + TimeSerial(0, 0, gPresenceSec)
    Application.OnTime gNextPresenceAt, "AutoSync_PresenceTick": gPresenceScheduled = True
End Sub
Public Sub AutoSync_PresenceTick()
    gPresenceScheduled = False
    If Not gRunning Then Exit Sub
    HeartbeatActiveSelection
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Not IsHiddenSheet(ws) Then RefreshPresenceMarkers ws
    Next ws
    SchedulePresence
End Sub

Public Sub AutoSync_SheetChanged(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    If Not gRunning Or TypeName(Sh) <> "Worksheet" Then Exit Sub
    Dim ws As Worksheet: Set ws = Sh
    If IsHiddenSheet(ws) Then Exit Sub

    EnsureSheetSchema ws
    Dim lo As ListObject: Set lo = FirstTable(ws)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    If Intersect(Target, lo.DataBodyRange) Is Nothing Then Exit Sub

    SchedulePush
    HeartbeatActiveSelection
End Sub

Private Sub EnsureSheetSchema(ByVal ws As Worksheet)
    Dim lo As ListObject
    If ws.ListObjects.Count = 0 Then
        ws.Cells.Clear
        ws.Range("A1").Value = META_ID
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(2, 1), , xlYes)
        lo.Name = "Table1"
    Else
        Set lo = ws.ListObjects(1)
    End If

    EnsureMetaColumns lo, Array(META_ID, META_VER, META_UPD, META_DEL)

    Dim headers As Variant: headers = lo.HeaderRowRange.Value
    Dim firstRow As Variant: If lo.DataBodyRange Is Nothing Then firstRow = Empty Else firstRow = lo.DataBodyRange.Resize(1).Value

    On Error Resume Next
    Gql_CreateTable ws.Name, headers, firstRow
    Gql_AddColumns ws.Name, headers, firstRow
    On Error GoTo 0
End Sub

Private Sub EnsureMetaColumns(ByVal lo As ListObject, ByVal metaArr As Variant)
    Dim have As Object: Set have = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        have(UCase$(CStr(lo.HeaderRowRange.Cells(1, i).Value))) = True
    Next i
    For i = LBound(metaArr) To UBound(metaArr)
        Dim nm As String: nm = CStr(metaArr(i))
        If Not have.Exists(UCase$(nm)) Then
            lo.HeaderRowRange.Cells(1, lo.HeaderRowRange.Columns.Count + 1).Value = nm
        End If
    Next i
End Sub

Private Sub PullOnce(ByVal ws As Worksheet, ByVal sinceVer As Long)
    Dim arr As Object: Set arr = Gql_LoadRows(ws.Name, sinceVer)
    If arr Is Nothing Then Exit Sub
    If arr.Count = 0 And sinceVer > 0 Then Exit Sub
    ApplySnapshot ws, arr
End Sub

Private Sub PushDirty(ByVal ws As Worksheet)
    Dim lo As ListObject: Set lo = FirstTable(ws)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub

    Dim idx As Object: Set idx = HeaderIndexMap(lo)
    Dim rowsJson As String: rowsJson = ""
    Dim r As Range
    For Each r In lo.DataBodyRange.Rows
        Dim one As String: one = RowToUpsertJson(r, idx)
        If Len(one) > 0 Then
            If Len(rowsJson) > 0 Then rowsJson = rowsJson & ","
            rowsJson = rowsJson & one
        End If
    Next r
    If Len(rowsJson) = 0 Then Exit Sub

    Dim payload As Object: Set payload = Gql_UpsertRows(ws.Name, Environ$("USERNAME"), rowsJson)
    Dim results As Object: Set results = payload("results")
    Dim snapshot As Object: Set snapshot = payload("snapshot")

    HandleResults ws, results, lo, idx
    ApplySnapshot ws, snapshot
End Sub

Private Sub HandleResults(ByVal ws As Worksheet, ByVal results As Object, ByVal lo As ListObject, ByVal idx As Object)
    If results Is Nothing Then Exit Sub
    Dim backupWS As Worksheet: Set backupWS = EnsureHiddenSheet(HIDDEN_CONFLICT_SHEET)

    Dim idToRow As Object: Set idToRow = CreateObject("Scripting.Dictionary")
    Dim rr As Range
    If Not lo.DataBodyRange Is Nothing Then
        For Each rr In lo.DataBodyRange.Rows
            Dim idv As Variant: idv = SafeCell(rr, idx(META_ID))
            If Len(CStr(idv)) > 0 Then idToRow(CStr(idv)) = rr.Row
        Next rr
    End If

    Dim it As Variant
    For Each it In results
        Dim st As String: st = it("status")
        Dim rid As Variant: rid = it("id")
        If IsNull(rid) Then GoTo NextIt
        If idToRow.Exists(CStr(rid)) Then
            Dim rownum As Long: rownum = CLng(idToRow(CStr(rid)))
            Select Case st
                Case "ok": Rows(rownum).Interior.Color = RGB(220, 255, 220)
                Case "conflict"
                    BackupRow backupWS, ws, rownum
                    Rows(rownum).Interior.Color = RGB(255, 245, 170)
                Case Else
                    Rows(rownum).Interior.Color = RGB(255, 200, 200)
            End Select
        End If
NextIt:
    Next it
End Sub

Private Sub BackupRow(ByVal backupWS As Worksheet, ByVal srcWS As Worksheet, ByVal rowNum As Long)
    Dim lastRow As Long: lastRow = backupWS.Cells(backupWS.Rows.Count, 1).End(xlUp).Row
    If lastRow < 1 Then lastRow = 1
    backupWS.Cells(lastRow + 1, 1).Value = Now
    backupWS.Cells(lastRow + 1, 2).Value = srcWS.Name
    backupWS.Cells(lastRow + 1, 3).Value = rowNum
    backupWS.Cells(lastRow + 1, 4).Value = Join(GetRowValues(srcWS.Rows(rowNum)), "|")
End Sub

Private Function GetRowValues(ByVal rng As Range) As Variant
    Dim arr() As String, i As Long, n As Long: n = rng.Columns.Count
    ReDim arr(1 To n)
    For i = 1 To n
        arr(i) = CStr(rng.Cells(1, i).Value)
    Next i
    GetRowValues = arr
End Function

Private Sub ApplySnapshot(ByVal ws As Worksheet, ByVal rows As Object)
    If rows Is Nothing Then Exit Sub
    Dim lo As ListObject: Set lo = FirstTable(ws)
    If lo Is Nothing Then Exit Sub

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo FIN

    If rows.Count = 0 Then GoTo FIN

    Dim first As Object: Set first = rows(1)
    Dim k As Variant, i As Long: i = 0
    For Each k In first.Keys
        i = i + 1: ws.Cells(1, i).Value = CStr(k)
    Next k

    If lo.DataBodyRange Is Nothing Then
        lo.Resize ws.Range(ws.Cells(1, 1), ws.Cells(2, i))
    Else
        lo.DataBodyRange.Delete
    End If

    Dim r As Long, c As Long
    For r = 1 To rows.Count
        Dim row As Object: Set row = rows(r)
        Dim lr As ListRow: Set lr = lo.ListRows.Add
        c = 0
        For Each k In first.Keys
            c = c + 1
            lr.Range.Cells(1, c).Value = IIf(row.Exists(k), row(k), Empty)
        Next k
    Next r

FIN:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Private Sub HeartbeatActiveSelection()
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ActiveSheet
    If ws Is Nothing Or IsHiddenSheet(ws) Then Exit Sub
    Dim lo As ListObject: Set lo = FirstTable(ws): If lo Is Nothing Then Exit Sub

    Dim addr As String: addr = ""
    Dim rid As Variant: rid = Empty
    Dim colname As String: colname = ""

    If Not Selection Is Nothing Then
        addr = Selection.Cells(1, 1).Address(False, False)
        If Not lo.DataBodyRange Is Nothing Then
            If Not Intersect(Selection, lo.DataBodyRange) Is Nothing Then
                Dim idx As Object: Set idx = HeaderIndexMap(lo)
                If idx.Exists("id") Then
                    rid = Selection.Cells(1, idx("id")).EntireRow.Cells(1, idx("id")).Value
                    colname = CStr(lo.HeaderRowRange.Cells(1, Selection.Column - lo.HeaderRowRange.Column + 1).Value)
                End If
            End If
        End If
    End If

    ' GraphQL presenceHeartbeat 호출
    Dim q As String, vars As String
    q = "mutation HB($u:String!,$t:String!,$ca:String,$rid:Int,$cn:String){ presenceHeartbeat(user:$u,table:$t,cell_addr:$ca,row_id:$rid,col_name:$cn) }"
    Dim parts As String: parts = """u"":" & JsonQuote(gNick) & ",""t"":" & JsonQuote(ws.Name)
    If Len(addr) > 0 Then parts = parts & ",""ca"":" & JsonQuote(addr)
    If Not IsEmpty(rid) And Len(CStr(rid)) > 0 Then parts = parts & ",""rid"":" & CStr(CLng(rid))
    If Len(colname) > 0 Then parts = parts & ",""cn"":" & JsonQuote(colname)
    vars = "{" & parts & "}"
    On Error Resume Next: Call GqlCall(q, vars): On Error GoTo 0
End Sub

Private Sub RefreshPresenceMarkers(ByVal ws As Worksheet)
    Dim lo As ListObject: Set lo = FirstTable(ws): If lo Is Nothing Then Exit Sub

    Dim sh As Shape
    For Each sh In ws.Shapes
        If Left$(sh.Name, 12) = "presenceDot_" Then sh.Delete
    Next sh

    Dim q As String, vars As String
    q = "query P($t:String!){ presence(table:$t){ user table_name cell_addr row_id col_name ts } }"
    vars = "{""t"":" & JsonQuote(ws.Name) & "}"
    Dim data As Object: On Error Resume Next: Set data = GqlCall(q, vars): On Error GoTo 0
    If data Is Nothing Then Exit Sub
    Dim arr As Object: Set arr = data("presence")
    If arr Is Nothing Or arr.Count = 0 Then Exit Sub

    Dim it As Variant
    For Each it In arr
        Dim nick As String: nick = it("user")
        If nick = gNick Then GoTo NextIt
        Dim cellAddr As String: cellAddr = it("cell_addr")
        If Len(cellAddr) = 0 Then GoTo NextIt
        Dim tgt As Range: Set tgt = Nothing
        On Error Resume Next: Set tgt = ws.Range(cellAddr): On Error GoTo 0
        If tgt Is Nothing Then GoTo NextIt

        On Error Resume Next: If Not tgt.Comment Is Nothing Then tgt.Comment.Delete: On Error GoTo 0
        tgt.AddComment "편집중: " & nick: tgt.Comment.Visible = False

        Dim dot As Shape
        Set dot = ws.Shapes.AddShape(msoShapeOval, tgt.Left + 2, tgt.Top + 2, 6, 6)
        dot.Name = "presenceDot_" & nick & "_" & cellAddr
        dot.Fill.ForeColor.RGB = HashColor(nick)
        dot.Line.Visible = msoFalse
NextIt:
    Next it
End Sub

Private Function FirstTable(ByVal ws As Worksheet) As ListObject
    If ws.ListObjects.Count > 0 Then Set FirstTable = ws.ListObjects(1) Else Set FirstTable = Nothing
End Function

Private Function HeaderIndexMap(ByVal lo As ListObject) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        d(CStr(lo.HeaderRowRange.Cells(1, i).Value)) = i
    Next i
    Set HeaderIndexMap = d
End Function

Private Function SafeCell(ByVal r As Range, ByVal colIdx As Variant) As Variant
    If IsEmpty(colIdx) Then SafeCell = Empty: Exit Function
    If colIdx <= 0 Or colIdx > r.Columns.Count Then SafeCell = Empty: Exit Function
    SafeCell = r.Cells(1, colIdx).Value
End Function

Private Function SheetMaxVersion(ByVal ws As Worksheet) As Long
    Dim lo As ListObject: Set lo = FirstTable(ws)
    If lo Is Nothing Then SheetMaxVersion = 0: Exit Function
    Dim idx As Object: Set idx = HeaderIndexMap(lo)
    If idx Is Nothing Or Not idx.Exists(META_VER) Then SheetMaxVersion = 0: Exit Function
    Dim maxv As Double: maxv = 0
    If Not lo.DataBodyRange Is Nothing Then
        Dim c As Range
        For Each c In lo.DataBodyRange.Columns(idx(META_VER)).Cells
            If IsNumeric(c.Value) Then If CDbl(c.Value) > maxv Then maxv = CDbl(c.Value)
        Next c
    End If
    SheetMaxVersion = CLng(maxv)
End Function

Private Function RowToUpsertJson(ByVal r As Range, ByVal idx As Object) As String
    Dim idv As Variant: idv = SafeCell(r, idx(META_ID))
    Dim ver As Variant: ver = SafeCell(r, idx(META_VER))
    Dim sb As String: sb = ""
    Dim i As Long
    For i = 1 To r.Columns.Count
        Dim name As String: name = CStr(r.ListObject.HeaderRowRange.Cells(1, i).Value)
        If name <> META_ID And name <> META_UPD And name <> META_VER And name <> META_DEL Then
            If Len(sb) > 0 Then sb = sb & ","
            sb = sb & """" & Replace(name, """", "\""") & """: " & JVal(r.Cells(1, i).Value)
        End If
    Next i
    Dim dataJson As String: dataJson = "{" & sb & "}"
    If Len(CStr(idv)) = 0 Then
        RowToUpsertJson = "{""data"":" & dataJson & "}"
    Else
        RowToUpsertJson = "{""id"":" & CStr(CLng(idv)) & ",""row_version"":" & CStr(CLng(ZeroIfEmpty(ver))) & ",""data"":" & dataJson & "}"
    End If
End Function

Private Function JVal(v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Then
        JVal = "null"
    ElseIf VarType(v) = vbBoolean Then
        JVal = IIf(v, "true", "false")
    ElseIf IsNumeric(v) Then
        JVal = CStr(v)
    Else
        JVal = """" & Replace(CStr(v), """", "\""") & """"
    End If
End Function

Private Function EnsureHiddenSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set EnsureHiddenSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If EnsureHiddenSheet Is Nothing Then
        Set EnsureHiddenSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        EnsureHiddenSheet.Name = name
        EnsureHiddenSheet.Visible = xlSheetVeryHidden
        EnsureHiddenSheet.Cells(1, 1).Value = "ts"
        EnsureHiddenSheet.Cells(1, 2).Value = "sheet"
        EnsureHiddenSheet.Cells(1, 3).Value = "row"
        EnsureHiddenSheet.Cells(1, 4).Value = "snapshot"
    End If
End Function

Private Function IsHiddenSheet(ByVal ws As Worksheet) As Boolean
    IsHiddenSheet = (ws.Name Like "_" & "*")
End Function

Private Function ZeroIfEmpty(ByVal v As Variant) As Long
    If IsEmpty(v) Or Not IsNumeric(v) Then ZeroIfEmpty = 0 Else ZeroIfEmpty = CLng(v)
End Function

Private Function HashColor(ByVal s As String) As Long
    Dim h As Long: h = 0
    Dim i As Long
    For i = 1 To Len(s)
        h = (h * 131 + AscW(Mid$(s, i, 1))) And &H7FFFFFFF
    Next i
    HashColor = RGB((h And &HFF), ((h \ 8) And &HFF), ((h \ 16) And &HFF))
End Function

Private Function GetOrPromptNickname() As String
    On Error Resume Next
    Dim nm As Name: Set nm = ThisWorkbook.Names("_Nick")
    On Error GoTo 0
    If Not nm Is Nothing Then
        GetOrPromptNickname = CStr(nm.RefersToRange.Value)
        If Len(GetOrPromptNickname) > 0 Then Exit Function
    End If
    Dim v As String
    Do
        v = InputBox("서버 접속 닉네임을 입력하세요:", "Nickname")
        If Len(v) = 0 Then MsgBox "닉네임이 필요합니다.", vbExclamation
    Loop While Len(v) = 0
    SaveNickname v
    GetOrPromptNickname = v
End Function

Private Sub SaveNickname(ByVal v As String)
    On Error Resume Next
    Dim nm As Name: Set nm = ThisWorkbook.Names("_Nick")
    If nm Is Nothing Then
        ThisWorkbook.Names.Add Name:="_Nick", RefersTo:=v
    Else
        nm.RefersTo = v
    End If
End Sub
