Attribute VB_Name = "Starter"

Public sysObj As Object '�ý��� ������Ʈ
Public folder As Object '���� ������Ʈ
Public book_name As String  '�� ���� �̸�
Public path As String '�� ���
Public department_dic As Object
Public file_dic As Object
Public numOfFile_dic As Object
Public TIME_BUFFER As Integer
    

Public Sub Start(Optional ByRef book As Workbook)

Try:
    On Error GoTo Catch
    
    TIME_BUFFER = 5
    book_name = ThisWorkbook.Name
    Dim folder_name As String '���� �̸� = ���� �̸�����
        folder_name = Left(book_name, InStrRev(book_name, ".") - 1)
    path = ThisWorkbook.path & "\" & folder_name
    
    Sheet1.Cells(1, 1).Value = folder_name & " ������ �о� ����"
    
    Set sysObj = CreateObject("Scripting.FileSystemObject")
    Set folder = sysObj.getFolder(path)
    
    Set department_dic = CreateObject("Scripting.Dictionary") '�μ� ��ųʸ�
    Set file_dic = CreateObject("Scripting.Dictionary") '���� ���� ��ųʸ�
    Set numOfFile_dic = CreateObject("Scripting.Dictionary") '�μ��� ���� ���� ����

    Exit Sub
    
Catch:
    MsgBox folder_name & " ������ �����ϴ�.", vbCritical, "���"
    
End Sub

Public Function test()
    Dim n, m As Integer
    n = 1
    m = 1
    
    Debug.Print ("����� ������:")
    
    For Each dep In department_dic.keys
        Debug.Print (m & ", key: " & dep & ", item: " & department_dic.Item(dep) & ", num: " & numOfFile_dic(dep))
        m = m + 1
    Next dep

    For Each file In file_dic.keys
        Debug.Print (n & ", key: " & file & ", item:" & file_dic.Item(file))
        n = n + 1
    Next
    
'    Dim lenOfQuery As Integer
'        lenOfQuery = Len(ThisWorkbook.Queries.Item("������å��").Formula)
'    Dim loc_open_braket As Integer
'        loc_open_braket = InStr(1, ThisWorkbook.Queries.Item("������å��").Formula, "{")
'    Dim loc_close_braket As Integer
'        loc_close_braket = InStr(1, ThisWorkbook.Queries.Item("������å��").Formula, "}")
'
'    Dim firstPart As String
'        firstPart = Left(ThisWorkbook.Queries.Item("������å��").Formula, loc_open_braket)
'    Dim query_list As String
'        query_list = Mid(ThisWorkbook.Queries.Item("������å��").Formula, loc_open_braket + 1, loc_close_braket - loc_open_braket - 1)
'    Dim lastPart As String
'        lastPart = Mid(ThisWorkbook.Queries.Item("������å��").Formula, loc_close_braket)
'
'    Debug.Print (ThisWorkbook.Queries.Item("������å��").Formula)
'    Debug.Print ("ó��: " & firstPart)
'    Debug.Print ("����Ʈ: " & query_list)
'    Debug.Print ("������: " & lastPart)
End Function
