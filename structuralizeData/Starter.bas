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

    Exit Sub
    
Catch:
    MsgBox folder_name & " 폴더가 없습니다.", vbCritical, "경고"
    
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
