Attribute VB_Name = "FileHandler"
  
Public Sub AddFiles(ByVal department As String, ByVal row_num As Integer)
    If numOfFile_dic.Item(department) = 0 Then
        GetFiles department, department_dic.Item(department) '�μ��� �������� �ϳ��� ���� ������ GetFiles�� ȣ��
    Else
        Dim itWorks As Boolean '��Ʈ�� �� ���� �Ǿ��°�

        Dim target_ws As Worksheet
            Set target_ws = ThisWorkbook.Worksheets(department)
            target_ws.Copy after:=target_ws '�ӽ� ��Ʈ�� ����
            ThisWorkbook.Worksheets(target_ws.Index + 1).Name = department & "_tmp" '�ӽ� ��Ʈ �̸� ����
            ThisWorkbook.Worksheets(department & "_tmp").ListObjects(1).Name = department & "_tmp" '���̺� �̸� ����
        
        getSheetData query_name:=department & "_tmp", sheet_name:=department & "_tmp" '�ӽ� ��Ʈ���� ���� ����

        Dim query_list As String '��� ���� �̸�
            query_list = department & "_tmp"
            
        Dim query_name As String '���� �̸�
        
        Dim sub_path As String '�μ� ���� ���
            sub_path = path & "\" & department
        Dim sub_folder As Object '�μ� ���� ��ü
            Set sub_folder = sysObj.getFolder(sub_path)
        
        Dim file_full_path As String '������ ������ ���
        
        Dim col As Integer
            col = numOfFile_dic.Item(department) + 2
        
        Dim flag As Boolean '�ֽ� �������� Ȯ��
            flag = False
            
        For Each file In sub_folder.Files
            
            If Not file_dic.exists(file.Name) Then
                
                Debug.Print (Chr(13) & Chr(10) & "�о���� ����:")
                
                flag = True
                
                file_full_path = sub_path & "\" & file.Name
                Debug.Print (Chr(9) & file_full_path)
                query_name = department & (col - 1)
                
                '���� ���̺� �����
                Select Case Mid(file.Name, InStrRev(file.Name, ".") + 1)
                Case "pdf"
                    Debug.Print Chr(9) & "����: PDF����"
                    query_list = query_list & ", " & query_name
                    getPdfData query_name:=query_name, path:=file_full_path, folder_name:=department
                Case "xls"
                    Debug.Print Chr(9) & "����: xls����"
                    query_list = query_list & ", " & query_name
                    getXlsData query_name:=query_name, path:=file_full_path, folder_name:=department
                Case "xlsx"
                    Debug.Print Chr(9) & "����: xlsx����"
                    query_list = query_list & ", " & query_name
                    getXlsxData query_name:=query_name, path:=file_full_path, folder_name:=department
                Case Else
                    Debug.Print Chr(9) & "��� �͵� �ƴϴ�."
                End Select
                
                With Sheet1.Cells(row_num, col)
                        .Value = file.Name '���� ��Ʈ�� ���
                        .WrapText = True
                        .VerticalAlignment = xlTop
                End With
                file_dic.Add Sheet1.Cells(row_num, col).Value, col '��ųʸ��� �߰�
                                
                col = col + 1
            End If
        
        Next file
        
        If flag = True Then '������ ������
            numOfFile_dic.Item(department) = col - 2
            
            Debug.Print ("�̾���� ǥ: " & query_list)
            unionTable query_name:=department, table_list:=query_list  '���̺� �̾���̱�
            
            itWorks = makeSheetWithTable(sheet_name:=department, table_name:=department) '��Ʈ�� ���̺� ����
            
            ThisWorkbook.Save '�����ϰ�
        
        Else
            
            itWorks = True
            MsgBox prompt:=department & "�� ������ �����ϴ�.", Buttons:=vbOKOnly Or vbInformation, Title:="�̰��� �˸��̿�."
        
        End If
        
        If itWorks Then
            '���� ����
            For Each q In Split(query_list, ", ")
                deleteQuery (q)
            Next q
            deleteQuery (department)
            '�ӽ� ��Ʈ ����
            deleteSheet (department & "_tmp")
        End If
    End If
    
End Sub
Public Sub AddAllFiles(Optional ByVal district As String)
    
    Dim error_times As Integer
        
    For Each department In department_dic.keys
    
        error_times = 0
        For Each thisSheet In ThisWorkbook.Worksheets
            If thisSheet.Tab.Color = 255 Then
                error_times = error_times + 1
            End If
        Next thisSheet
        
        If error_times > 2 Then
            Exit For
        End If
        
        Debug.Print "���� ��Ʈ �� " & error_times & "��"
        AddFiles department, department_dic.Item(department)
    Next department

End Sub
Public Sub GetFiles(ByVal department As String, ByVal row_num As Integer)
    
    Dim itWork As Boolean '��Ʈ�� �� ���� �Ǿ��°�
    
    Dim query_name As String '���� �̸�
    Dim query_list As String '��� ���� �̸�
        query_list = ""
    Dim sub_path As String '�μ� ���� ���
        sub_path = path & "\" & department
    Dim sub_folder As Object '�μ� ���� ��ü
        Set sub_folder = sysObj.getFolder(sub_path)
    
    Dim file_full_path As String '������ ������ ���
    Dim num As Integer
        num = 1
    
    Debug.Print (Chr(13) & Chr(10) & "�о���� ����:")
        
    For Each file In sub_folder.Files
        
        file_full_path = sub_path & "\" & file.Name
        Debug.Print (Chr(9) & file_full_path)
        query_name = department & num
        
        '���� ���̺� �����
        Select Case Mid(file.Name, InStrRev(file.Name, ".") + 1)
        Case "pdf"
            Debug.Print Chr(9) & "����: PDF����"
            If query_list = "" Then
                query_list = query_name
            Else
                query_list = query_list & ", " & query_name
            End If
            getPdfData query_name:=query_name, path:=file_full_path, folder_name:=department
        Case "xls"
            Debug.Print Chr(9) & "����: xls����"
            If query_list = "" Then
                query_list = query_name
            Else
                query_list = query_list & ", " & query_name
            End If
            getXlsData query_name:=query_name, path:=file_full_path, folder_name:=department
        Case "xlsx"
            Debug.Print Chr(9) & "����: xlsx����"
            If query_list = "" Then
                query_list = query_name
            Else
                query_list = query_list & ", " & query_name
            End If
            getXlsxData query_name:=query_name, path:=file_full_path, folder_name:=department
        Case Else
            Debug.Print Chr(9) & "��� �͵� �ƴϴ�."
        End Select
                
        With Sheet1.Cells(row_num, num + 1)
                .Value = file.Name '���� ��Ʈ�� ���
                .WrapText = True
                .VerticalAlignment = xlTop
        End With
        file_dic.Add Sheet1.Cells(row_num, num + 1).Value, num + 1 '��ųʸ��� �߰�
        
        num = num + 1
        
    Next file
    numOfFile_dic.Item(department) = num - 1
        
    Debug.Print ("�̾���� ǥ: " & query_list)
    unionTable query_name:=department, table_list:=query_list '���̺� �̾���̱�
    
    itWork = makeSheetWithTable(sheet_name:=department, table_name:=department) '��Ʈ�� ���̺� ����
    
    ThisWorkbook.Save '�����ϰ�
    
    If itWork Then
        '���� ����
        For Each q In Split(query_list, ", ")
            deleteQuery (q)
        Next q
        deleteQuery (department)
    End If
    
End Sub
Public Sub GetAllFiles(Optional ByVal district As String)
    
    Dim error_times As Integer
        
    For Each department In department_dic.keys
    
        error_times = 0
        For Each thisSheet In ThisWorkbook.Worksheets
            If thisSheet.Tab.Color = 255 Then
                error_times = error_times + 1
            End If
        Next thisSheet
        
        If error_times > 2 Then
            Exit For
        End If
        
        Debug.Print "���� ��Ʈ �� " & error_times & "��"
        GetFiles department, department_dic.Item(department)
    Next department

End Sub
Function makeSheetWithTable(ByVal sheet_name As String, ByVal table_name As String) As Boolean

    If isSheet(sheet_name) Then
        ThisWorkbook.Worksheets(sheet_name).Select
        ThisWorkbook.Worksheets(sheet_name).Cells.Clear
    Else
        ThisWorkbook.Worksheets.Add.Name = sheet_name
        ThisWorkbook.Worksheets(sheet_name).Move after:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
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
    Debug.Print "�� �ٿ���."
    makeSheetWithTable = True
    
    Exit Function
    
Catch:
    
    Worksheets(sheet_name).Tab.Color = 255
    Debug.Print "�� �� �ٿ���."
    makeSheetWithTable = False

End Function

Function deleteSheet(ByVal sheet_name As String)
    On Error Resume Next
    ActiveWorkbook.Worksheets(sheet_name).Delete
    On Error GoTo 0
End Function

Function isSheet(ByVal sheet_name As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
     Set ws = ActiveWorkbook.Worksheets(sheet_name)
    On Error GoTo 0
    isSheet = Not ws Is Nothing
End Function

Function getPdfData(ByVal query_name As String, ByVal path As String, ByVal folder_name)
    
    deleteQuery (query_name)
    Dim pre_q As String
    Dim mid_q As String
    Dim end_q As String
    Dim full_query As String
    Dim year As String
        year = Right(ThisWorkbook.path, 4)
    
    pre_q = "let" & Chr(13) & Chr(10) & _
            Chr(9) & "wait = (seconds as number, action as function) =>" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "if (List.Count(List.Generate(() => DateTimeZone.LocalNow() + #duration(0, 0, 0, seconds), (x) => DateTimeZone.LocalNow() < x,  (x) => x)) = 0) then null else action()," & Chr(13) & Chr(10) & _
            Chr(9) & "Pause = wait(" & TIME_BUFFER & ", DateTime.LocalNow)," & Chr(13) & Chr(10) & _
            Chr(9) & "Source = Pdf.Tables(File.Contents(""" & path & """), [Implementation=""1.3""])," & Chr(13) & Chr(10) & _
            Chr(9) & "SelectedTable = Table.SelectRows(Source, each [Kind] = ""Table"")," & Chr(13) & Chr(10) & _
            Chr(9) & "TableList = List.Generate(() => [x = 0, y = SelectedTable{x}[Data]], each [x] < Table.RowCount(SelectedTable), each [x = [x] + 1, y = SelectedTable{x}[Data]], each [y])," & Chr(13) & Chr(10) & _
            Chr(9) & "CombinedTable = Table.Combine(TableList)," & Chr(13) & Chr(10) & _
            Chr(9) & "NullIndexDelTable = Table.RemoveMatchingRows(CombinedTable, {[Column1 = null], [Column1 = """"]}, ""Column1"")," & Chr(13) & Chr(10) & _
            Chr(9) & "LastRow = List.Count(Table.ToList(NullIndexDelTable))," & Chr(13) & Chr(10) & _
            Chr(9) & "countNull = (myList as list, offset as number) =>" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "if offset < List.Count(myList) then" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "if myList{offset} = null or myList{offset} = """" then" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "@countNull(myList, offset + 1) + 1" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "else" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "@countNull(myList, offset + 1)" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "else 0," & Chr(13) & Chr(10) & _
            Chr(9) & "deleteNullRows = (myTable as table, offset as number, end as number) =>" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "if offset < end then" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "if countNull(Record.FieldValues(myTable{offset}), 0) > 3 then" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "@deleteNullRows(Table.RemoveRows(myTable, offset), offset + 1, end - 1)" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "else" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "@deleteNullRows(myTable, offset + 1, end)" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "else " & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "myTable," & Chr(13) & Chr(10)
    mid_q = Chr(9) & "NullDeletedTable = deleteNullRows(NullIndexDelTable, 0, LastRow)," & Chr(13) & Chr(10) & _
            Chr(9) & "PromotedHeaders = Table.PromoteHeaders(NullDeletedTable, [PromoteAllScalars=true])," & Chr(13) & Chr(10) & _
            Chr(9) & "ColumnList = Table.ColumnNames(PromotedHeaders)," & Chr(13) & Chr(10) & _
            Chr(9) & "RowCleanedTable = Table.RemoveMatchingRows(PromotedHeaders, ColumnList)," & Chr(13) & Chr(10) & _
            Chr(9) & "SpaceErasedColumnList = List.Generate(() => [x = 0, y = Text.Remove(Text.Clean(ColumnList{x}), "" "")], each [x] < List.Count(ColumnList), each [x = [x] + 1, y = Text.Remove(Text.Clean(ColumnList{x}), "" "")], each [y] )," & Chr(13) & Chr(10) & _
            Chr(9) & "exchangeColumn = (colText as text, comText as text) =>" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "if Text.Contains(colText, ""��ġ"") and not Text.Contains(comText, ""��ġ"") then ""��ġ""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "else if Text.Contains(colText, ""���"") or Text.Contains(colText, ""������"") or Text.Contains(colText, ""���ó"") or Text.Contains(colText, ""����"") or Text.Contains(colText, ""��ü"") or Text.Contains(colText, ""��ȣ"") and not Text.Contains(comText, ""���"") then ""���""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "else if Text.Contains(colText, ""����"") or Text.Contains(colText, ""����"") or Text.Contains(colText, ""����"") or Text.Contains(colText, ""����"") and not Text.Contains(comText, ""����"") then ""����""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "else if Text.Contains(colText, ""�μ�"") and not Text.Contains(comText, ""�μ�"") then ""�μ�""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "else if Text.Contains(colText, ""�����"") and not Text.Contains(comText, ""�����"") then ""�����""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "else if Text.Contains(colText, ""���"") and not Text.Contains(comText, ""���"") then ""���""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "else if Text.Contains(colText, ""�ο�"") or Text.Contains(colText, ""�����"") or Text.Contains(colText, ""������"") and not Text.Contains(comText, ""�ο�"") then ""�ο�""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "else if Text.Contains(colText, ""��"") and not Text.Contains(comText, ""�ݾ�"") then ""�ݾ�""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "else if Text.Contains(colText, ""����"") and not Text.Contains(comText, ""����"") then ""����""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "else if Text.Contains(colText, ""��"") and not Text.Contains(comText, ""��¥"") then ""��¥""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "else if Text.Contains(colText, ""��"") and not Text.Contains(comText, ""�ð�"") then ""�ð�""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "else colText," & Chr(13) & Chr(10) & _
            Chr(9) & "NewColumnList = List.Generate(() => [x = 0, y = exchangeColumn(SpaceErasedColumnList{x}, """"), z = y], each [x] < List.Count(SpaceErasedColumnList), each [x = [x] + 1, y = exchangeColumn(SpaceErasedColumnList{x}, [z]), z = [z] & y], each [y])," & Chr(13) & Chr(10) & _
            Chr(9) & "HeaderPairList = List.Generate(() => [x = 0, y = {ColumnList{x}, NewColumnList{x}}], each [x] < List.Count(NewColumnList), each [x = [x] + 1, y = {ColumnList{x}, NewColumnList{x}}], each [y])," & Chr(13) & Chr(10) & _
            Chr(9) & "RenamedTable = Table.RenameColumns(RowCleanedTable, HeaderPairList)," & Chr(13) & Chr(10) & _
            Chr(9) & "ColumnUnifiedTable = if List.Contains(NewColumnList, ""�μ�"") then RenamedTable else Table.AddColumn(RenamedTable, ""�μ�"", each """ & folder_name & """, type text)," & Chr(13) & Chr(10) & _
            Chr(9) & "extractNumber = (rowText) => if Type.Is(Value.Type(rowText), type text) then Text.Select(rowText, {""0""..""9""}) else rowText," & Chr(13) & Chr(10) & _
            Chr(9) & "TransfromedTable = Table.TransformColumns(ColumnUnifiedTable," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "{" & Chr(13) & Chr(10)
    end_q = Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""����"", Number.From}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""�ο�"", each extractNumber(_)}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""��¥"", each #date(" & year & ", Date.Month(DateTime.From(_)), Date.Day(DateTime.From(_)))}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""�ð�"", Text.From}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""�ݾ�"", each extractNumber(_)}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""�����"", Text.Clean}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""����"", Text.Clean}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""���"", Text.Clean}" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "Text.From," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "MissingField.UseNull" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & ")," & Chr(13) & Chr(10) & _
            Chr(9) & "TypeTransfromedTable = Table.TransformColumnTypes(TransfromedTable," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "{" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""����"", Int64.Type}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""�ο�"", Int64.Type}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""��¥"", type date}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""�ð�"", type time}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""�ݾ�"", Currency.Type}" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "}" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & ")," & Chr(13) & Chr(10) & _
            Chr(9) & "ErrorCorrectedTable = Table.ReplaceErrorValues(TypeTransfromedTable, {{""��¥"", null}, {""�ð�"", null}})," & Chr(13) & Chr(10) & _
            Chr(9) & "ErrorRemovedTable = Table.RemoveRowsWithErrors(ErrorCorrectedTable)" & Chr(13) & Chr(10) & _
            "in" & Chr(13) & Chr(10) & _
            Chr(9) & "ErrorRemovedTable"
    full_query = pre_q & mid_q & end_q
    
    ActiveWorkbook.Queries.Add Name:=query_name, Formula:=full_query

End Function

Function getXlsData(ByVal query_name As String, ByVal path As String, ByVal folder_name)
    
    deleteQuery (query_name)
    Dim pre_q As String
    Dim mid_q As String
    Dim end_q As String
    Dim full_query As String
    Dim year As String
        year = Right(ThisWorkbook.path, 4)
    
    pre_q = "let" & Chr(13) & Chr(10) & _
            Chr(9) & "wait = (seconds as number, action as function) =>" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "if (List.Count(List.Generate(() => DateTimeZone.LocalNow() + #duration(0, 0, 0, seconds), (x) => DateTimeZone.LocalNow() < x,  (x) => x)) = 0) then null else action()," & Chr(13) & Chr(10) & _
            Chr(9) & "Pause = wait(" & TIME_BUFFER & ", DateTime.LocalNow)," & Chr(13) & Chr(10) & _
            Chr(9) & "OriginalSource = Excel.Workbook(File.Contents(""" & path & """), null, true)," & Chr(13) & Chr(10) & _
            Chr(9) & "Source = Table.SelectRows(OriginalSource, each not Text.Contains([Name], ""$""))," & Chr(13) & Chr(10) & _
            Chr(9) & "SheetNameList = List.Generate(() => [x = 0, y = Source{x}[Name]], each [x] < Table.RowCount(Source), each [x = [x] + 1, y = Source{x}[Name]], each [y])," & Chr(13) & Chr(10) & _
            Chr(9) & "makeTable = (oneTable as table, name as text) =>" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "let" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "NullIndexDelTable = Table.RemoveMatchingRows(oneTable, {[Column1 = null], [Column1 = """"]}, ""Column1"")," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "RawColumnList = Table.ColumnNames(NullIndexDelTable)," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "deleteNullColumn = (myTable as table, myColumn as list, offset as number) =>" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "if offset < List.Count(myColumn) then" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "if List.NonNullCount(Table.Column(myTable, myColumn{offset})) = 0 then" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "@deleteNullColumn(Table.RemoveColumns(myTable, myColumn{offset}), myColumn, offset + 1)" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "@deleteNullColumn(myTable, myColumn, offset + 1)" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "myTable," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "NullColumnDelTable = deleteNullColumn(NullIndexDelTable, RawColumnList, 0)," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "LastRow = List.Count(Table.ToList(NullColumnDelTable))," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "countNull = (myList as list, offset as number) =>" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "if offset < List.Count(myList) then" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "if myList{offset} = null or myList{offset} = """" then" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "@countNull(myList, offset + 1) + 1" & Chr(13) & Chr(10)
    mid_q = Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "@countNull(myList, offset + 1)" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else 0," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "deleteNullRows = (myTable as table, offset as number, end as number) =>" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "if offset < end then" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "if countNull(Record.FieldValues(myTable{offset}), 0) > 3 then" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "@deleteNullRows(Table.RemoveRows(myTable, offset), offset, end - 1)" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "@deleteNullRows(myTable, offset + 1, end)" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else " & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "myTable," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "NullDeletedTable = deleteNullRows(NullColumnDelTable, 0, LastRow)," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "PromotedHeaders = Table.PromoteHeaders(NullDeletedTable, [PromoteAllScalars=true])," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "ColumnList = Table.ColumnNames(PromotedHeaders)," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "SpaceErasedColumnList = List.Generate(() => [x = 0, y = Text.Remove(Text.Clean(ColumnList{x}), "" "")], each [x] < List.Count(ColumnList), each [x = [x] + 1, y = Text.Remove(Text.Clean(ColumnList{x}), "" "")], each [y] )," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "exchangeColumn = (colText as text, comText as text) =>" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "if Text.Contains(colText, ""��ġ"") and not Text.Contains(comText, ""��ġ"") then ""��ġ""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""���"") or Text.Contains(colText, ""������"") or Text.Contains(colText, ""���ó"") or Text.Contains(colText, ""����"") or Text.Contains(colText, ""��ü"") or Text.Contains(colText, ""��ȣ"") and not Text.Contains(comText, ""���"") then ""���""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""����"") or Text.Contains(colText, ""����"") or Text.Contains(colText, ""����"") or Text.Contains(colText, ""����"") and not Text.Contains(comText, ""����"") then ""����""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""�μ�"") and not Text.Contains(comText, ""�μ�"") then ""�μ�""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""�����"") and not Text.Contains(comText, ""�����"") then ""�����""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""���"") and not Text.Contains(comText, ""���"") then ""���""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""�ο�"") or Text.Contains(colText, ""�����"") or Text.Contains(colText, ""������"") and not Text.Contains(comText, ""�ο�"") then ""�ο�""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""��"") and not Text.Contains(comText, ""�ݾ�"") then ""�ݾ�""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""����"") and not Text.Contains(comText, ""����"") then ""����""" & Chr(13) & Chr(10)
    end_q = Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""��"") and not Text.Contains(comText, ""��¥"") then ""��¥""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""��"") and not Text.Contains(comText, ""�ð�"") then ""�ð�""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else colText," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "NewColumnList = List.Generate(() => [x = 0, y = exchangeColumn(SpaceErasedColumnList{x}, """"), z = y], each [x] < List.Count(SpaceErasedColumnList), each [x = [x] + 1, y = exchangeColumn(SpaceErasedColumnList{x}, [z]), z = [z] & y], each [y])," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "HeaderPairList = List.Generate(() => [x = 0, y = {ColumnList{x}, NewColumnList{x}}], each [x] < List.Count(NewColumnList), each [x = [x] + 1, y = {ColumnList{x}, NewColumnList{x}}], each [y])," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "RenamedTable = Table.RenameColumns(PromotedHeaders, HeaderPairList)," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "UserAddedTable = " & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "if List.Contains(NewColumnList, ""�����"") then RenamedTable" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(name, ""sheet"", Comparer.OrdinalIgnoreCase) then RenamedTable" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else Table.AddColumn(RenamedTable, ""�����"", each name, type text)," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "ColumnUnifiedTable = if List.Contains(NewColumnList, ""�μ�"") then UserAddedTable else Table.AddColumn(UserAddedTable, ""�μ�"", each """ & folder_name & """, type text)," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "extractNumber = (rowText) => if Type.Is(Value.Type(rowText), type text) then Text.Select(rowText, {""0""..""9""}) else rowText," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "TransfromedTable = Table.TransformColumns(ColumnUnifiedTable," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""����"", Number.From}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""�ο�"", each extractNumber(_)}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""��¥"", each #date(" & year & ", Date.Month(DateTime.From(_)), Date.Day(DateTime.From(_)))}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""�ð�"", Text.From}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""�ݾ�"", each extractNumber(_)}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""�����"", Text.Clean}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""����"", Text.Clean}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""���"", Text.Clean}" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "Text.From," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "MissingField.UseNull" & Chr(13) & Chr(10)
    full_query = pre_q & mid_q & end_q & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & ")," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "TypeTransfromedTable = Table.TransformColumnTypes(TransfromedTable," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""����"", Int64.Type}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""�ο�"", Int64.Type}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""��¥"", type date}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""�ð�"", type time}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""�ݾ�"", Currency.Type}" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "}" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & ")," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "ErrorCorrectedTable = Table.ReplaceErrorValues(TypeTransfromedTable, {{""��¥"", null}, {""�ð�"", null}})" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "in" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "ErrorCorrectedTable," & Chr(13) & Chr(10) & _
            Chr(9) & "TableList = List.Generate(() => [x = 0, y = makeTable(Source{x}[Data], SheetNameList{x})], each [x] < List.Count(SheetNameList), each [x = [x] + 1, y = makeTable(Source{x}[Data], SheetNameList{x})], each [y])," & Chr(13) & Chr(10) & _
            Chr(9) & "CombinedTable = Table.Combine(TableList)" & Chr(13) & Chr(10) & _
            "in" & Chr(13) & Chr(10) & _
            Chr(9) & "CombinedTable"
            
    ActiveWorkbook.Queries.Add Name:=query_name, Formula:=full_query

End Function

Function getXlsxData(ByVal query_name As String, ByVal path As String, ByVal folder_name)
    
    deleteQuery (query_name)
    Dim pre_q As String
    Dim mid_q As String
    Dim end_q As String
    Dim full_query As String
    Dim year As String
        year = Right(ThisWorkbook.path, 4)
    
    pre_q = "let" & Chr(13) & Chr(10) & _
            Chr(9) & "wait = (seconds as number, action as function) =>" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "if (List.Count(List.Generate(() => DateTimeZone.LocalNow() + #duration(0, 0, 0, seconds), (x) => DateTimeZone.LocalNow() < x,  (x) => x)) = 0) then null else action()," & Chr(13) & Chr(10) & _
            Chr(9) & "Pause = wait(" & TIME_BUFFER & ", DateTime.LocalNow)," & Chr(13) & Chr(10) & _
            Chr(9) & "OriginalSource = Excel.Workbook(File.Contents(""" & path & """), null, true)," & Chr(13) & Chr(10) & _
            Chr(9) & "Source = Table.SelectRows(OriginalSource, each [Kind] = ""Sheet""), " & Chr(13) & Chr(10) & _
            Chr(9) & "SheetNameList = List.Generate(() => [x = 0, y = Source{x}[Name]], each [x] < Table.RowCount(Source), each [x = [x] + 1, y = Source{x}[Name]], each [y])," & Chr(13) & Chr(10) & _
            Chr(9) & "makeTable = (oneTable as table, name as text) =>" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "let" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "NullIndexDelTable = Table.RemoveMatchingRows(oneTable, {[Column1 = null], [Column1 = """"]}, ""Column1"")," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "RawColumnList = Table.ColumnNames(NullIndexDelTable)," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "deleteNullColumn = (myTable as table, myColumn as list, offset as number) =>" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "if offset < List.Count(myColumn) then" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "if List.NonNullCount(Table.Column(myTable, myColumn{offset})) = 0 then" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "@deleteNullColumn(Table.RemoveColumns(myTable, myColumn{offset}), myColumn, offset + 1)" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "@deleteNullColumn(myTable, myColumn, offset + 1)" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "myTable," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "NullColumnDelTable = deleteNullColumn(NullIndexDelTable, RawColumnList, 0)," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "LastRow = List.Count(Table.ToList(NullColumnDelTable))," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "countNull = (myList as list, offset as number) =>" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "if offset < List.Count(myList) then" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "if myList{offset} = null or myList{offset} = """" then" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "@countNull(myList, offset + 1) + 1" & Chr(13) & Chr(10)
    mid_q = Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "@countNull(myList, offset + 1)" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else 0," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "deleteNullRows = (myTable as table, offset as number, end as number) =>" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "if offset < end then" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "if countNull(Record.FieldValues(myTable{offset}), 0) > 3 then" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "@deleteNullRows(Table.RemoveRows(myTable, offset), offset, end - 1)" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "@deleteNullRows(myTable, offset + 1, end)" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else " & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "myTable," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "NullDeletedTable = deleteNullRows(NullColumnDelTable, 0, LastRow)," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "PromotedHeaders = Table.PromoteHeaders(NullDeletedTable, [PromoteAllScalars=true])," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "ColumnList = Table.ColumnNames(PromotedHeaders)," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "SpaceErasedColumnList = List.Generate(() => [x = 0, y = Text.Remove(Text.Clean(ColumnList{x}), "" "")], each [x] < List.Count(ColumnList), each [x = [x] + 1, y = Text.Remove(Text.Clean(ColumnList{x}), "" "")], each [y] )," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "exchangeColumn = (colText as text, comText as text) =>" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "if Text.Contains(colText, ""��ġ"") and not Text.Contains(comText, ""��ġ"") then ""��ġ""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""���"") or Text.Contains(colText, ""������"") or Text.Contains(colText, ""���ó"") or Text.Contains(colText, ""����"") or Text.Contains(colText, ""��ü"") or Text.Contains(colText, ""��ȣ"") and not Text.Contains(comText, ""���"") then ""���""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""����"") or Text.Contains(colText, ""����"") or Text.Contains(colText, ""����"") or Text.Contains(colText, ""����"") and not Text.Contains(comText, ""����"") then ""����""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""�μ�"") and not Text.Contains(comText, ""�μ�"") then ""�μ�""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""�����"") and not Text.Contains(comText, ""�����"") then ""�����""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""���"") and not Text.Contains(comText, ""���"") then ""���""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""�ο�"") or Text.Contains(colText, ""�����"") or Text.Contains(colText, ""������"") and not Text.Contains(comText, ""�ο�"") then ""�ο�""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""��"") and not Text.Contains(comText, ""�ݾ�"") then ""�ݾ�""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""����"") and not Text.Contains(comText, ""����"") then ""����""" & Chr(13) & Chr(10)
    end_q = Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""��"") and not Text.Contains(comText, ""��¥"") then ""��¥""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""��"") and not Text.Contains(comText, ""�ð�"") then ""�ð�""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else colText," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "NewColumnList = List.Generate(() => [x = 0, y = exchangeColumn(SpaceErasedColumnList{x}, """"), z = y], each [x] < List.Count(SpaceErasedColumnList), each [x = [x] + 1, y = exchangeColumn(SpaceErasedColumnList{x}, [z]), z = [z] & y], each [y])," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "HeaderPairList = List.Generate(() => [x = 0, y = {ColumnList{x}, NewColumnList{x}}], each [x] < List.Count(NewColumnList), each [x = [x] + 1, y = {ColumnList{x}, NewColumnList{x}}], each [y])," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "RenamedTable = Table.RenameColumns(PromotedHeaders, HeaderPairList)," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "UserAddedTable = " & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "if List.Contains(NewColumnList, ""�����"") then RenamedTable" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(name, ""sheet"", Comparer.OrdinalIgnoreCase) then RenamedTable" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else Table.AddColumn(RenamedTable, ""�����"", each name, type text)," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "ColumnUnifiedTable = if List.Contains(NewColumnList, ""�μ�"") then UserAddedTable else Table.AddColumn(UserAddedTable, ""�μ�"", each """ & folder_name & """, type text)," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "extractNumber = (rowText) => if Type.Is(Value.Type(rowText), type text) then Text.Select(rowText, {""0""..""9""}) else rowText," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "TransfromedTable = Table.TransformColumns(ColumnUnifiedTable," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""����"", Number.From}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""�ο�"", each extractNumber(_)}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""��¥"", each #date(" & year & ", Date.Month(DateTime.From(_)), Date.Day(DateTime.From(_)))}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""�ð�"", Text.From}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""�ݾ�"", each extractNumber(_)}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""�����"", Text.Clean}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""����"", Text.Clean}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""���"", Text.Clean}" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "Text.From," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "MissingField.UseNull" & Chr(13) & Chr(10)
    full_query = pre_q & mid_q & end_q & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & ")," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "TypeTransfromedTable = Table.TransformColumnTypes(TransfromedTable," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""����"", Int64.Type}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""�ο�"", Int64.Type}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""��¥"", type date}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""�ð�"", type time}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""�ݾ�"", Currency.Type}" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "}" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & ")," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "ErrorCorrectedTable = Table.ReplaceErrorValues(TypeTransfromedTable, {{""��¥"", null}, {""�ð�"", null}})" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "in" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "ErrorCorrectedTable," & Chr(13) & Chr(10) & _
            Chr(9) & "TableList = List.Generate(() => [x = 0, y = makeTable(Source{x}[Data], SheetNameList{x})], each [x] < List.Count(SheetNameList), each [x = [x] + 1, y = makeTable(Source{x}[Data], SheetNameList{x})], each [y])," & Chr(13) & Chr(10) & _
            Chr(9) & "CombinedTable = Table.Combine(TableList)" & Chr(13) & Chr(10) & _
            "in" & Chr(13) & Chr(10) & _
            Chr(9) & "CombinedTable"
    
    ActiveWorkbook.Queries.Add Name:=query_name, Formula:=full_query

End Function

Function getSheetData(ByVal query_name As String, ByVal sheet_name As String)
    deleteQuery (query_name)
    Dim full_query As String
    full_query = "let" & Chr(13) & Chr(10) & _
                Chr(9) & "Source = Excel.CurrentWorkbook(){[Name=""" & sheet_name & """]}[Content]," & Chr(13) & Chr(10) & _
                Chr(9) & "TypeTransfromedTable = Table.TransformColumnTypes(Source," & Chr(13) & Chr(10) & _
                Chr(9) & Chr(9) & "{" & Chr(13) & Chr(10) & _
                Chr(9) & Chr(9) & Chr(9) & "{""����"", Int64.Type}," & Chr(13) & Chr(10) & _
                Chr(9) & Chr(9) & Chr(9) & "{""�ο�"", Int64.Type}," & Chr(13) & Chr(10) & _
                Chr(9) & Chr(9) & Chr(9) & "{""��¥"", type date}," & Chr(13) & Chr(10) & _
                Chr(9) & Chr(9) & Chr(9) & "{""�ð�"", type time}," & Chr(13) & Chr(10) & _
                Chr(9) & Chr(9) & Chr(9) & "{""�ݾ�"", Currency.Type}," & Chr(13) & Chr(10) & _
                Chr(9) & Chr(9) & Chr(9) & "{""�����"", Text.Type}," & Chr(13) & Chr(10) & _
                Chr(9) & Chr(9) & Chr(9) & "{""���"", Text.Type}," & Chr(13) & Chr(10) & _
                Chr(9) & Chr(9) & Chr(9) & "{""�μ�"", Text.Type}," & Chr(13) & Chr(10) & _
                Chr(9) & Chr(9) & Chr(9) & "{""����"", Text.Type}" & Chr(13) & Chr(10) & _
                Chr(9) & Chr(9) & "}" & Chr(13) & Chr(10) & _
                Chr(9) & ")," & Chr(13) & Chr(10) & _
                Chr(9) & "ErrorCorrectedTable = Table.ReplaceErrorValues(TypeTransfromedTable, {{""��¥"", null}, {""�ð�"", null}})" & Chr(13) & Chr(10) & _
                "in" & Chr(13) & Chr(10) & _
                Chr(9) & "ErrorCorrectedTable" & Chr(13) & Chr(10)
    
    ActiveWorkbook.Queries.Add Name:=query_name, Formula:=full_query
    
End Function

Function unionTable(ByVal query_name As String, ByVal table_list As String)
    deleteQuery (query_name)
    Dim full_query As String
    full_query = "let" & Chr(13) & Chr(10) & _
                Chr(9) & "wait = (seconds as number, action as function) =>" & Chr(13) & Chr(10) & _
                Chr(9) & Chr(9) & "if (List.Count(List.Generate(() => DateTimeZone.LocalNow() + #duration(0, 0, 0, seconds), (x) => DateTimeZone.LocalNow() < x,  (x) => x)) = 0) then null else action()," & Chr(13) & Chr(10) & _
                Chr(9) & "Pause = wait(" & TIME_BUFFER & ", DateTime.LocalNow)," & Chr(13) & Chr(10) & _
                Chr(9) & "Source = Table.Combine({" & table_list & "})," & Chr(13) & Chr(10) & _
                Chr(9) & "OriginalColumns = Table.ColumnNames(Source)," & Chr(13) & Chr(10) & _
                Chr(9) & "PositionOfNum = List.PositionOf(OriginalColumns, ""����"")," & Chr(13) & Chr(10) & _
                Chr(9) & "OrderedTable = Table.Sort(Source, {{""��¥"", Order.Ascending}, {""�ð�"", Order.Ascending}})," & Chr(13) & Chr(10) & _
                Chr(9) & "TmpTable = Table.RemoveColumns(OrderedTable, OriginalColumns{PositionOfNum})," & Chr(13) & Chr(10) & _
                Chr(9) & "IndexAddedTable = Table.AddIndexColumn(TmpTable, OriginalColumns{PositionOfNum}, 1, 1, Int64.Type)," & Chr(13) & Chr(10) & _
                Chr(9) & "ReorderedTable = Table.ReorderColumns(IndexAddedTable, OriginalColumns)" & Chr(13) & Chr(10) & _
                "in" & Chr(13) & Chr(10) & _
                Chr(9) & "ReorderedTable"
                
    ActiveWorkbook.Queries.Add Name:=query_name, Formula:=full_query
    
End Function

Function deleteQuery(ByVal query_name As String)
    On Error Resume Next
    ActiveWorkbook.Queries.Item(query_name).Delete
    On Error GoTo 0
End Function

