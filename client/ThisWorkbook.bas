Private Sub Workbook_Open()
    AutoSync_Init "http://localhost:8000", "devkey", 10, 2, True, "", 3
    AutoSync_Start
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    AutoSync_Stop
End Sub

' 셀 값 변경
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    AutoSync_SheetChanged Sh, Target
End Sub

' 표(ListObject) 구조/데이터 변경(행/열 추가·삭제 등)
Private Sub Workbook_SheetTableUpdate(ByVal Sh As Object, ByVal Target As TableObject)
    On Error Resume Next
    ' TableObject가 주어지므로 데이터 영역을 넘겨서 저장 트리거
    If Not Target Is Nothing Then
        AutoSync_SheetChanged Sh, Target.DataBodyRange
    Else
        AutoSync_SheetChanged Sh, Nothing
    End If
End Sub

