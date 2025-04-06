VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "현재_통합_문서"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

    ThisWorkbook.Unprotect ("")
    With Sheet1
        .Unprotect ("")
        .Activate
        With Cells(6, 1)
            .Value = "설정 로드 중"
            .Font.Bold = True
        End With
    End With
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Starter.Variable_Set '파일 이름 및 경로 가져오기
    
    Sheet1.Cells(6, 1).Clear
    Sheet1.Protect ("")
    ThisWorkbook.Protect ("")
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

