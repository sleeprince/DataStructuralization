Attribute VB_Name = "Starter"

Public sysObj As Object '시스템 오브젝트
Public folder As Object '폴더 오브젝트
Public book_name As String  '이 파일 이름
Public path As String '총 경로
Public department_dic As Object
Public file_dic As Object
Public numOfFile_dic As Object
Public TIME_BUFFER As Integer
    

Public Sub Start(Optional ByRef book As Workbook)

Try:
    On Error GoTo Catch
    
    TIME_BUFFER = 5
    book_name = ThisWorkbook.Name
    Dim folder_name As String '파일 이름 = 폴더 이름으로
        folder_name = Left(book_name, InStrRev(book_name, ".") - 1)
    path = ThisWorkbook.path & "\" & folder_name
    
    Sheet1.Cells(1, 1).Value = folder_name & " 데이터 읽어 오기"
    
    Set sysObj = CreateObject("Scripting.FileSystemObject")
    Set folder = sysObj.getFolder(path)
    
    Set department_dic = CreateObject("Scripting.Dictionary") '부서 딕셔너리
    Set file_dic = CreateObject("Scripting.Dictionary") '읽은 파일 딕셔너리
    Set numOfFile_dic = CreateObject("Scripting.Dictionary") '부서별 읽은 파일 갯수
    
    Call Dictionary_Initialize
    
    Exit Sub
    
Catch:
    MsgBox folder_name & " 폴더가 없습니다.", vbCritical, "경고"
    
End Sub
Public Sub Dictionary_Initialize()

    Sheet1.ComboBox1.Clear
    Sheet1.ComboBox1.AddItem pvargItem:="전체"
    
    Dim row, col As Integer
        row = 10
    Dim folerNm As String
    
    Dim data_dic As Object
    Set data_dic = CreateObject("Scripting.Dictionary") '원래 기록되어 있던 부서 목록
    
    While Sheet1.Cells(row, 1).Value <> ""
        data_dic.Add Sheet1.Cells(row, 1).Value, row
        row = row + 1
    Wend
    
    Debug.Print ("원래 부서:")
    For Each datum In data_dic.keys
        Debug.Print ("key: " & datum & ", item: " & data_dic(datum))
    Next datum
    
    For Each sub_folder In folder.subFolders
    
        folderNm = sub_folder.Name
        
        Sheet1.ComboBox1.AddItem pvargItem:=folderNm
        
        col = 2
        If data_dic.exists(folderNm) Then
            While Sheet1.Cells(data_dic.Item(folderNm), col).Value <> ""
                file_dic.Add Sheet1.Cells(data_dic.Item(folderNm), col).Value, col
                col = col + 1
            Wend
            data_dic.Remove (folderNm)
        Else
            Sheet1.Cells(row, 1).Value = folderNm
            row = row + 1
        End If
        numOfFile_dic.Add folderNm, (col - 2)
        
    Next sub_folder
    
    Dim data_arr
    If data_dic.Count > 0 Then
        data_arr = data_dic.items
        Debug.Print ("지울 부서:")
        For i = UBound(data_arr) To 0 Step -1
            Debug.Print ("key: " & Sheet1.Cells(data_arr(i), 1).Value & ", item: " & data_arr(i))
            col = 2
            deleteQuery (Sheet1.Cells(data_arr(i), 1).Value)
            While Sheet1.Cells(data_arr(i), col).Value <> ""
                deleteQuery (Sheet1.Cells(data_arr(i), 1).Value & (col - 1))
                col = col + 1
            Wend
            deleteSheet (Sheet1.Cells(data_arr(i), 1).Value)
            data_dic.Remove (Sheet1.Cells(data_arr(i), 1).Value)
            Sheet1.Rows(data_arr(i)).Delete shift:=xlUp
            row = row - 1
        Next
    End If
    
    For rw = 10 To row - 1
        department_dic.Add Sheet1.Cells(rw, 1).Value, rw
    Next
    
'    Starter.test
    
    Sheet1.ComboBox1.Text = "전체"
End Sub

Public Function test()
    Dim n, m As Integer
    n = 1
    m = 1
    
    Debug.Print ("저장된 데이터:")
    
    For Each dep In department_dic.keys
        Debug.Print (m & ", key: " & dep & ", item: " & department_dic.Item(dep) & ", num: " & numOfFile_dic(dep))
        m = m + 1
    Next dep

    For Each file In file_dic.keys
        Debug.Print (n & ", key: " & file & ", item:" & file_dic.Item(file))
        n = n + 1
    Next
    
'    Dim lenOfQuery As Integer
'        lenOfQuery = Len(ThisWorkbook.Queries.Item("가족정책과").Formula)
'    Dim loc_open_braket As Integer
'        loc_open_braket = InStr(1, ThisWorkbook.Queries.Item("가족정책과").Formula, "{")
'    Dim loc_close_braket As Integer
'        loc_close_braket = InStr(1, ThisWorkbook.Queries.Item("가족정책과").Formula, "}")
'
'    Dim firstPart As String
'        firstPart = Left(ThisWorkbook.Queries.Item("가족정책과").Formula, loc_open_braket)
'    Dim query_list As String
'        query_list = Mid(ThisWorkbook.Queries.Item("가족정책과").Formula, loc_open_braket + 1, loc_close_braket - loc_open_braket - 1)
'    Dim lastPart As String
'        lastPart = Mid(ThisWorkbook.Queries.Item("가족정책과").Formula, loc_close_braket)
'
'    Debug.Print (ThisWorkbook.Queries.Item("가족정책과").Formula)
'    Debug.Print ("처음: " & firstPart)
'    Debug.Print ("리스트: " & query_list)
'    Debug.Print ("마지막: " & lastPart)
End Function
