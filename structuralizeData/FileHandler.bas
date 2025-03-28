Attribute VB_Name = "FileHandler"
  
Public Sub AddFiles(ByVal department As String, ByVal row_num As Integer)
    If numOfFile_dic.Item(department) = 0 Then
        GetFiles department, department_dic.Item(department) '부서내 아이템이 하나도 없는 때에는 GetFiles를 호출
    Else
        Dim itWorks As Boolean '시트에 잘 복사 되었는가

        Dim target_ws As Worksheet
            Set target_ws = ThisWorkbook.Worksheets(department)
            target_ws.Copy after:=target_ws '임시 시트로 복사
            ThisWorkbook.Worksheets(target_ws.Index + 1).Name = department & "_tmp" '임시 시트 이름 설정
            ThisWorkbook.Worksheets(department & "_tmp").ListObjects(1).Name = department & "_tmp" '테이블 이름 설정
        
        getSheetData query_name:=department & "_tmp", sheet_name:=department & "_tmp" '임시 시트에서 쿼리 생성

        Dim query_list As String '모든 쿼리 이름
            query_list = department & "_tmp"
            
        Dim query_name As String '쿼리 이름
        
        Dim sub_path As String '부서 폴더 경로
            sub_path = path & "\" & department
        Dim sub_folder As Object '부서 폴더 객체
            Set sub_folder = sysObj.getFolder(sub_path)
        
        Dim file_full_path As String '파일을 포함한 경로
        
        Dim col As Integer
            col = numOfFile_dic.Item(department) + 2
        
        Dim flag As Boolean '최신 상태인지 확인
            flag = False
            
        For Each file In sub_folder.Files
            
            If Not file_dic.exists(file.Name) Then
                
                Debug.Print (Chr(13) & Chr(10) & "읽어들인 파일:")
                
                flag = True
                
                file_full_path = sub_path & "\" & file.Name
                Debug.Print (Chr(9) & file_full_path)
                query_name = department & (col - 1)
                
                '쿼리 테이블 만들기
                Select Case Mid(file.Name, InStrRev(file.Name, ".") + 1)
                Case "pdf"
                    Debug.Print Chr(9) & "종류: PDF파일"
                    query_list = query_list & ", " & query_name
                    getPdfData query_name:=query_name, path:=file_full_path, folder_name:=department
                Case "xls"
                    Debug.Print Chr(9) & "종류: xls파일"
                    query_list = query_list & ", " & query_name
                    getXlsData query_name:=query_name, path:=file_full_path, folder_name:=department
                Case "xlsx"
                    Debug.Print Chr(9) & "종류: xlsx파일"
                    query_list = query_list & ", " & query_name
                    getXlsxData query_name:=query_name, path:=file_full_path, folder_name:=department
                Case Else
                    Debug.Print Chr(9) & "어느 것도 아니다."
                End Select
                
                With Sheet1.Cells(row_num, col)
                        .Value = file.Name '관리 시트에 기록
                        .WrapText = True
                        .VerticalAlignment = xlTop
                End With
                file_dic.Add Sheet1.Cells(row_num, col).Value, col '딕셔너리에 추가
                                
                col = col + 1
            End If
        
        Next file
        
        If flag = True Then '새것이 있으면
            numOfFile_dic.Item(department) = col - 2
            
            Debug.Print ("이어붙일 표: " & query_list)
            unionTable query_name:=department, table_list:=query_list  '테이블 이어붙이기
            
            itWorks = makeSheetWithTable(sheet_name:=department, table_name:=department) '시트에 테이블 복사
            
            ThisWorkbook.Save '저장하고
        
        Else
            
            itWorks = True
            MsgBox prompt:=department & "에 새것이 없습니다.", Buttons:=vbOKOnly Or vbInformation, Title:="이것은 알림이오."
        
        End If
        
        If itWorks Then
            '쿼리 삭제
            For Each q In Split(query_list, ", ")
                deleteQuery (q)
            Next q
            deleteQuery (department)
            '임시 시트 삭제
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
        
        Debug.Print "오류 시트 총 " & error_times & "개"
        AddFiles department, department_dic.Item(department)
    Next department

End Sub
Public Sub GetFiles(ByVal department As String, ByVal row_num As Integer)
    
    Dim itWork As Boolean '시트에 잘 복사 되었는가
    
    Dim query_name As String '쿼리 이름
    Dim query_list As String '모든 쿼리 이름
        query_list = ""
    Dim sub_path As String '부서 폴더 경로
        sub_path = path & "\" & department
    Dim sub_folder As Object '부서 폴더 객체
        Set sub_folder = sysObj.getFolder(sub_path)
    
    Dim file_full_path As String '파일을 포함한 경로
    Dim num As Integer
        num = 1
    
    Debug.Print (Chr(13) & Chr(10) & "읽어들인 파일:")
        
    For Each file In sub_folder.Files
        
        file_full_path = sub_path & "\" & file.Name
        Debug.Print (Chr(9) & file_full_path)
        query_name = department & num
        
        '쿼리 테이블 만들기
        Select Case Mid(file.Name, InStrRev(file.Name, ".") + 1)
        Case "pdf"
            Debug.Print Chr(9) & "종류: PDF파일"
            If query_list = "" Then
                query_list = query_name
            Else
                query_list = query_list & ", " & query_name
            End If
            getPdfData query_name:=query_name, path:=file_full_path, folder_name:=department
        Case "xls"
            Debug.Print Chr(9) & "종류: xls파일"
            If query_list = "" Then
                query_list = query_name
            Else
                query_list = query_list & ", " & query_name
            End If
            getXlsData query_name:=query_name, path:=file_full_path, folder_name:=department
        Case "xlsx"
            Debug.Print Chr(9) & "종류: xlsx파일"
            If query_list = "" Then
                query_list = query_name
            Else
                query_list = query_list & ", " & query_name
            End If
            getXlsxData query_name:=query_name, path:=file_full_path, folder_name:=department
        Case Else
            Debug.Print Chr(9) & "어느 것도 아니다."
        End Select
                
        With Sheet1.Cells(row_num, num + 1)
                .Value = file.Name '관리 시트에 기록
                .WrapText = True
                .VerticalAlignment = xlTop
        End With
        file_dic.Add Sheet1.Cells(row_num, num + 1).Value, num + 1 '딕셔너리에 추가
        
        num = num + 1
        
    Next file
    numOfFile_dic.Item(department) = num - 1
        
    Debug.Print ("이어붙일 표: " & query_list)
    unionTable query_name:=department, table_list:=query_list '테이블 이어붙이기
    
    itWork = makeSheetWithTable(sheet_name:=department, table_name:=department) '시트에 테이블 복사
    
    ThisWorkbook.Save '저장하고
    
    If itWork Then
        '쿼리 삭제
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
        
        Debug.Print "오류 시트 총 " & error_times & "개"
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
    Debug.Print "잘 붙였다."
    makeSheetWithTable = True
    
    Exit Function
    
Catch:
    
    Worksheets(sheet_name).Tab.Color = 255
    Debug.Print "잘 못 붙였다."
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
            Chr(9) & Chr(9) & "if Text.Contains(colText, ""위치"") and not Text.Contains(comText, ""위치"") then ""위치""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "else if Text.Contains(colText, ""장소"") or Text.Contains(colText, ""가맹점"") or Text.Contains(colText, ""사용처"") or Text.Contains(colText, ""업소"") or Text.Contains(colText, ""업체"") or Text.Contains(colText, ""상호"") and not Text.Contains(comText, ""장소"") then ""장소""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "else if Text.Contains(colText, ""목적"") or Text.Contains(colText, ""내역"") or Text.Contains(colText, ""내용"") or Text.Contains(colText, ""사유"") and not Text.Contains(comText, ""목적"") then ""목적""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "else if Text.Contains(colText, ""부서"") and not Text.Contains(comText, ""부서"") then ""부서""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "else if Text.Contains(colText, ""사용자"") and not Text.Contains(comText, ""사용자"") then ""사용자""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "else if Text.Contains(colText, ""방법"") and not Text.Contains(comText, ""방법"") then ""방법""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "else if Text.Contains(colText, ""인원"") or Text.Contains(colText, ""대상자"") or Text.Contains(colText, ""집행대상"") and not Text.Contains(comText, ""인원"") then ""인원""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "else if Text.Contains(colText, ""액"") and not Text.Contains(comText, ""금액"") then ""금액""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "else if Text.Contains(colText, ""순번"") and not Text.Contains(comText, ""연번"") then ""연번""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "else if Text.Contains(colText, ""일"") and not Text.Contains(comText, ""날짜"") then ""날짜""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "else if Text.Contains(colText, ""시"") and not Text.Contains(comText, ""시간"") then ""시간""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & "else colText," & Chr(13) & Chr(10) & _
            Chr(9) & "NewColumnList = List.Generate(() => [x = 0, y = exchangeColumn(SpaceErasedColumnList{x}, """"), z = y], each [x] < List.Count(SpaceErasedColumnList), each [x = [x] + 1, y = exchangeColumn(SpaceErasedColumnList{x}, [z]), z = [z] & y], each [y])," & Chr(13) & Chr(10) & _
            Chr(9) & "HeaderPairList = List.Generate(() => [x = 0, y = {ColumnList{x}, NewColumnList{x}}], each [x] < List.Count(NewColumnList), each [x = [x] + 1, y = {ColumnList{x}, NewColumnList{x}}], each [y])," & Chr(13) & Chr(10) & _
            Chr(9) & "RenamedTable = Table.RenameColumns(RowCleanedTable, HeaderPairList)," & Chr(13) & Chr(10) & _
            Chr(9) & "ColumnUnifiedTable = if List.Contains(NewColumnList, ""부서"") then RenamedTable else Table.AddColumn(RenamedTable, ""부서"", each """ & folder_name & """, type text)," & Chr(13) & Chr(10) & _
            Chr(9) & "extractNumber = (rowText) => if Type.Is(Value.Type(rowText), type text) then Text.Select(rowText, {""0""..""9""}) else rowText," & Chr(13) & Chr(10) & _
            Chr(9) & "TransfromedTable = Table.TransformColumns(ColumnUnifiedTable," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "{" & Chr(13) & Chr(10)
    end_q = Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""연번"", Number.From}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""인원"", each extractNumber(_)}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""날짜"", each #date(" & year & ", Date.Month(DateTime.From(_)), Date.Day(DateTime.From(_)))}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""시간"", Text.From}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""금액"", each extractNumber(_)}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""사용자"", Text.Clean}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""목적"", Text.Clean}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""장소"", Text.Clean}" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "Text.From," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "MissingField.UseNull" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & ")," & Chr(13) & Chr(10) & _
            Chr(9) & "TypeTransfromedTable = Table.TransformColumnTypes(TransfromedTable," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "{" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""연번"", Int64.Type}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""인원"", Int64.Type}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""날짜"", type date}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""시간"", type time}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""금액"", Currency.Type}" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "}" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & ")," & Chr(13) & Chr(10) & _
            Chr(9) & "ErrorCorrectedTable = Table.ReplaceErrorValues(TypeTransfromedTable, {{""날짜"", null}, {""시간"", null}})," & Chr(13) & Chr(10) & _
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
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "if Text.Contains(colText, ""위치"") and not Text.Contains(comText, ""위치"") then ""위치""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""장소"") or Text.Contains(colText, ""가맹점"") or Text.Contains(colText, ""사용처"") or Text.Contains(colText, ""업소"") or Text.Contains(colText, ""업체"") or Text.Contains(colText, ""상호"") and not Text.Contains(comText, ""장소"") then ""장소""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""목적"") or Text.Contains(colText, ""내역"") or Text.Contains(colText, ""내용"") or Text.Contains(colText, ""사유"") and not Text.Contains(comText, ""목적"") then ""목적""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""부서"") and not Text.Contains(comText, ""부서"") then ""부서""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""사용자"") and not Text.Contains(comText, ""사용자"") then ""사용자""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""방법"") and not Text.Contains(comText, ""방법"") then ""방법""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""인원"") or Text.Contains(colText, ""대상자"") or Text.Contains(colText, ""집행대상"") and not Text.Contains(comText, ""인원"") then ""인원""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""액"") and not Text.Contains(comText, ""금액"") then ""금액""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""순번"") and not Text.Contains(comText, ""연번"") then ""연번""" & Chr(13) & Chr(10)
    end_q = Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""일"") and not Text.Contains(comText, ""날짜"") then ""날짜""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""시"") and not Text.Contains(comText, ""시간"") then ""시간""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else colText," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "NewColumnList = List.Generate(() => [x = 0, y = exchangeColumn(SpaceErasedColumnList{x}, """"), z = y], each [x] < List.Count(SpaceErasedColumnList), each [x = [x] + 1, y = exchangeColumn(SpaceErasedColumnList{x}, [z]), z = [z] & y], each [y])," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "HeaderPairList = List.Generate(() => [x = 0, y = {ColumnList{x}, NewColumnList{x}}], each [x] < List.Count(NewColumnList), each [x = [x] + 1, y = {ColumnList{x}, NewColumnList{x}}], each [y])," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "RenamedTable = Table.RenameColumns(PromotedHeaders, HeaderPairList)," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "UserAddedTable = " & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "if List.Contains(NewColumnList, ""사용자"") then RenamedTable" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(name, ""sheet"", Comparer.OrdinalIgnoreCase) then RenamedTable" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else Table.AddColumn(RenamedTable, ""사용자"", each name, type text)," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "ColumnUnifiedTable = if List.Contains(NewColumnList, ""부서"") then UserAddedTable else Table.AddColumn(UserAddedTable, ""부서"", each """ & folder_name & """, type text)," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "extractNumber = (rowText) => if Type.Is(Value.Type(rowText), type text) then Text.Select(rowText, {""0""..""9""}) else rowText," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "TransfromedTable = Table.TransformColumns(ColumnUnifiedTable," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""연번"", Number.From}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""인원"", each extractNumber(_)}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""날짜"", each #date(" & year & ", Date.Month(DateTime.From(_)), Date.Day(DateTime.From(_)))}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""시간"", Text.From}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""금액"", each extractNumber(_)}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""사용자"", Text.Clean}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""목적"", Text.Clean}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""장소"", Text.Clean}" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "Text.From," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "MissingField.UseNull" & Chr(13) & Chr(10)
    full_query = pre_q & mid_q & end_q & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & ")," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "TypeTransfromedTable = Table.TransformColumnTypes(TransfromedTable," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""연번"", Int64.Type}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""인원"", Int64.Type}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""날짜"", type date}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""시간"", type time}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""금액"", Currency.Type}" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "}" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & ")," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "ErrorCorrectedTable = Table.ReplaceErrorValues(TypeTransfromedTable, {{""날짜"", null}, {""시간"", null}})" & Chr(13) & Chr(10) & _
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
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "if Text.Contains(colText, ""위치"") and not Text.Contains(comText, ""위치"") then ""위치""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""장소"") or Text.Contains(colText, ""가맹점"") or Text.Contains(colText, ""사용처"") or Text.Contains(colText, ""업소"") or Text.Contains(colText, ""업체"") or Text.Contains(colText, ""상호"") and not Text.Contains(comText, ""장소"") then ""장소""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""목적"") or Text.Contains(colText, ""내역"") or Text.Contains(colText, ""내용"") or Text.Contains(colText, ""사유"") and not Text.Contains(comText, ""목적"") then ""목적""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""부서"") and not Text.Contains(comText, ""부서"") then ""부서""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""사용자"") and not Text.Contains(comText, ""사용자"") then ""사용자""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""방법"") and not Text.Contains(comText, ""방법"") then ""방법""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""인원"") or Text.Contains(colText, ""대상자"") or Text.Contains(colText, ""집행대상"") and not Text.Contains(comText, ""인원"") then ""인원""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""액"") and not Text.Contains(comText, ""금액"") then ""금액""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""순번"") and not Text.Contains(comText, ""연번"") then ""연번""" & Chr(13) & Chr(10)
    end_q = Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""일"") and not Text.Contains(comText, ""날짜"") then ""날짜""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(colText, ""시"") and not Text.Contains(comText, ""시간"") then ""시간""" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else colText," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "NewColumnList = List.Generate(() => [x = 0, y = exchangeColumn(SpaceErasedColumnList{x}, """"), z = y], each [x] < List.Count(SpaceErasedColumnList), each [x = [x] + 1, y = exchangeColumn(SpaceErasedColumnList{x}, [z]), z = [z] & y], each [y])," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "HeaderPairList = List.Generate(() => [x = 0, y = {ColumnList{x}, NewColumnList{x}}], each [x] < List.Count(NewColumnList), each [x = [x] + 1, y = {ColumnList{x}, NewColumnList{x}}], each [y])," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "RenamedTable = Table.RenameColumns(PromotedHeaders, HeaderPairList)," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "UserAddedTable = " & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "if List.Contains(NewColumnList, ""사용자"") then RenamedTable" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else if Text.Contains(name, ""sheet"", Comparer.OrdinalIgnoreCase) then RenamedTable" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & "else Table.AddColumn(RenamedTable, ""사용자"", each name, type text)," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "ColumnUnifiedTable = if List.Contains(NewColumnList, ""부서"") then UserAddedTable else Table.AddColumn(UserAddedTable, ""부서"", each """ & folder_name & """, type text)," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "extractNumber = (rowText) => if Type.Is(Value.Type(rowText), type text) then Text.Select(rowText, {""0""..""9""}) else rowText," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "TransfromedTable = Table.TransformColumns(ColumnUnifiedTable," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""연번"", Number.From}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""인원"", each extractNumber(_)}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""날짜"", each #date(" & year & ", Date.Month(DateTime.From(_)), Date.Day(DateTime.From(_)))}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""시간"", Text.From}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""금액"", each extractNumber(_)}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""사용자"", Text.Clean}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""목적"", Text.Clean}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""장소"", Text.Clean}" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "Text.From," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "MissingField.UseNull" & Chr(13) & Chr(10)
    full_query = pre_q & mid_q & end_q & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & ")," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "TypeTransfromedTable = Table.TransformColumnTypes(TransfromedTable," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""연번"", Int64.Type}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""인원"", Int64.Type}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""날짜"", type date}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""시간"", type time}," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{""금액"", Currency.Type}" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "}" & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & Chr(9) & ")," & Chr(13) & Chr(10) & _
            Chr(9) & Chr(9) & Chr(9) & "ErrorCorrectedTable = Table.ReplaceErrorValues(TypeTransfromedTable, {{""날짜"", null}, {""시간"", null}})" & Chr(13) & Chr(10) & _
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
                Chr(9) & Chr(9) & Chr(9) & "{""연번"", Int64.Type}," & Chr(13) & Chr(10) & _
                Chr(9) & Chr(9) & Chr(9) & "{""인원"", Int64.Type}," & Chr(13) & Chr(10) & _
                Chr(9) & Chr(9) & Chr(9) & "{""날짜"", type date}," & Chr(13) & Chr(10) & _
                Chr(9) & Chr(9) & Chr(9) & "{""시간"", type time}," & Chr(13) & Chr(10) & _
                Chr(9) & Chr(9) & Chr(9) & "{""금액"", Currency.Type}," & Chr(13) & Chr(10) & _
                Chr(9) & Chr(9) & Chr(9) & "{""사용자"", Text.Type}," & Chr(13) & Chr(10) & _
                Chr(9) & Chr(9) & Chr(9) & "{""장소"", Text.Type}," & Chr(13) & Chr(10) & _
                Chr(9) & Chr(9) & Chr(9) & "{""부서"", Text.Type}," & Chr(13) & Chr(10) & _
                Chr(9) & Chr(9) & Chr(9) & "{""목적"", Text.Type}" & Chr(13) & Chr(10) & _
                Chr(9) & Chr(9) & "}" & Chr(13) & Chr(10) & _
                Chr(9) & ")," & Chr(13) & Chr(10) & _
                Chr(9) & "ErrorCorrectedTable = Table.ReplaceErrorValues(TypeTransfromedTable, {{""날짜"", null}, {""시간"", null}})" & Chr(13) & Chr(10) & _
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
                Chr(9) & "PositionOfNum = List.PositionOf(OriginalColumns, ""연번"")," & Chr(13) & Chr(10) & _
                Chr(9) & "OrderedTable = Table.Sort(Source, {{""날짜"", Order.Ascending}, {""시간"", Order.Ascending}})," & Chr(13) & Chr(10) & _
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

