Option Explicit

Private Sub Workbook_Open()
    AutoSync_Init "http://localhost:8000", "devkey", 10, 2, True, "", 3
    AutoSync_Start
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    AutoSync_Stop
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    AutoSync_SheetChanged Sh, Target
End Sub

Private Sub Workbook_SheetTableUpdate(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    AutoSync_SheetChanged Sh, Target
End Sub
