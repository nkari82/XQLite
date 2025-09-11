Option Explicit

Private Sub Workbook_Open()
    AutoSync_Init "http://localhost:8000", "devkey", 10, 2, True
    AutoSync_Start
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    AutoSync_Stop
End Sub

' 통합문서의 모든 시트 변경을 한 곳에서 감지 (여러 시트 자동 지원)
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    AutoSync_SheetChanged Sh, Target
End Sub

' 테이블 행 삭제/삽입 등 범위 변화도 저장을 유도
Private Sub Workbook_SheetTableUpdate(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    AutoSync_SheetChanged Sh, Target
End Sub


