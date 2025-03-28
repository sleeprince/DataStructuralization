VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Dim num As Integer
    Dim file_name As String
    Dim this_sheet As String
    Dim prev_sheet As String
    
    For Each file In folder.Files
        
        If Right(file.name, 3) = "pdf" Then
            '테이블 개수 가져오기
            getNumOfTable query_name:="numOfTable", file_name:=file
            makeSheetWithTable sheet_name:="numOfTable", table_name:="numOfTable"
            num = Worksheets("numOfTable").Cells(2, 1).Value
            file_name = Replace(Replace(Replace(Replace(Left(file.name, InStrRev(file.name, ".") - 1), " ", "_"), "(", "_"), ")", "_"), ".", "_")
            If Len(file_name) > 24 Then
                file_name = Left(file_name, 12) & Right(file_name, 12)
            End If
            Debug.Print "▶" & file_name & ": " & num
            
            
            '테이블 가져오기
            For i = 0 To num - 1
                getTable query_name:="table" & i & "_" & file_name, file_name:=file, nth:=i
                makeSheetWithTable sheet_name:="sheet" & i & "_" & file_name, table_name:="table" & i & "_" & file_name
                Debug.Print "table" & i & "_" & file_name & " 만듦"
                deleteQuery query_name:="table" & i & "_" & file_name
            Next
            
        End If
        
    Next file
    
    deleteQuery query_name:="numOfTable"
    deleteSheet sheet_name:="numOfTable"
    
    mkFile name:=folder.name, path:=folder_path
    
    For Each sh In ThisWorkbook.Worksheets
        this_sheet = sh.name
        If Not sh.name = "Sheet1" Then
            sh.Move after:=Workbooks(folder.name & ".xlsx").Worksheets(prev_sheet)
        End If
        prev_sheet = this_sheet
    Next sh
    Workbooks(folder.name & ".xlsx").Worksheets("Sheet1").Delete
    Workbooks(folder.name & ".xlsx").Save
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
