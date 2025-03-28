Attribute VB_Name = "Utility"
Public sysObj As Object
Public folder As Object
Public folder_path As String
Public file_name As String
    
Public Function mkFile(ByVal name As String, ByVal path As String) As Boolean

Try:
    On Error GoTo Catch
    Workbooks.Add.SaveAs Filename:=path & "\" & name _
                        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    mkFile = True
    Exit Function
    
Catch:
    mkFile = False
    MsgBox "파일을 만들지 못했습니다."

End Function

Function deleteSheet(ByVal sheet_name As String)
    On Error Resume Next
    ThisWorkbook.Worksheets(sheet_name).Delete
    On Error GoTo 0
End Function

Function isSheet(ByVal sheet_name As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
     Set ws = ThisWorkbook.Worksheets(sheet_name)
    On Error GoTo 0
    isSheet = Not ws Is Nothing
End Function

Function makeSheetWithTable(ByVal sheet_name As String, ByVal table_name As String) As Boolean

    If isSheet(sheet_name) Then
        ThisWorkbook.Worksheets(sheet_name).Select
        ThisWorkbook.Worksheets(sheet_name).Cells.Clear
    Else
        ThisWorkbook.Worksheets.Add.name = sheet_name
        ThisWorkbook.Worksheets(sheet_name).Move after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    End If
    
Try:

    On Error GoTo Catch
    
    With Worksheets(sheet_name).ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & table_name & ";Extended Properties=""""", Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & table_name & "]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = table_name
        .Refresh BackgroundQuery:=False
    End With
    ActiveWindow.SmallScroll Down:=-55
    
    Worksheets(sheet_name).Tab.ColorIndex = xlColorIndexNone
    Debug.Print "잘 붙였다."
    makeSheetWithTable = True
    
    Exit Function
    
Catch:
    
    Worksheets(sheet_name).Tab.Color = 255
    Debug.Print "잘 못 붙였다."
    makeSheetWithTable = False

End Function

Sub getNumOfTable(ByVal query_name As String, ByVal file_name As String)

    deleteQuery (query_name)
    Dim query_str As String
    
    query_str = "let" & Chr(13) & Chr(10) & _
            Chr(9) & "Source = Pdf.Tables(File.Contents(""" & file_name & """), [Implementation=""1.3""])," & Chr(13) & Chr(10) & _
            Chr(9) & "SelectedTable = Table.SelectRows(Source, each [Kind] = ""Table"")," & Chr(13) & Chr(10) & _
            Chr(9) & "TableList = List.Generate(() => [x = 0, y = SelectedTable{x}[Data]], each [x] < Table.RowCount(SelectedTable), each [x = [x] + 1, y = SelectedTable{x}[Data]], each [y])," & Chr(13) & Chr(10) & _
            Chr(9) & "CountList = List.Count(TableList)" & Chr(13) & Chr(10) & _
            "in" & Chr(13) & Chr(10) & _
            Chr(9) & "CountList"
    
    ThisWorkbook.Queries.Add name:=query_name, Formula:=query_str

End Sub

Sub getTable(ByVal query_name As String, ByVal file_name As String, ByVal nth As Integer)

    deleteQuery (query_name)
    Dim query_str As String
    
    query_str = "let" & Chr(13) & Chr(10) & _
            Chr(9) & "Source = Pdf.Tables(File.Contents(""" & file_name & """), [Implementation=""1.3""])," & Chr(13) & Chr(10) & _
            Chr(9) & "SelectedTable = Table.SelectRows(Source, each [Kind] = ""Table"")," & Chr(13) & Chr(10) & _
            Chr(9) & "TableList = List.Generate(() => [x = 0, y = SelectedTable{x}[Data]], each [x] < Table.RowCount(SelectedTable), each [x = [x] + 1, y = SelectedTable{x}[Data]], each [y])," & Chr(13) & Chr(10) & _
            Chr(9) & "MyTable = SelectedTable{" & nth & "}[Data]," & Chr(13) & Chr(10) & _
            Chr(9) & "PromotedHeaders = Table.PromoteHeaders(MyTable, [PromoteAllScalars=true])" & Chr(13) & Chr(10) & _
            "in" & Chr(13) & Chr(10) & _
            Chr(9) & "PromotedHeaders"
    
    ThisWorkbook.Queries.Add name:=query_name, Formula:=query_str

End Sub

Function deleteQuery(ByVal query_name As String)
    On Error Resume Next
    ThisWorkbook.Queries.Item(query_name).Delete
    On Error GoTo 0
End Function
