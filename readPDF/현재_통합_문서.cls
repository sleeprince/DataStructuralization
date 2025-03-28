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

    Let folder_path = ThisWorkbook.path
    Set sysObj = CreateObject("Scripting.FileSystemObject")
    Set folder = sysObj.getFolder(folder_path)
    
    
End Sub
