Attribute VB_Name = "DataService"
Public Sub insertFromExcelToMariaDB(ByVal file_path As String, ByVal file_name, ByVal sheet_name As String)

    Dim server As String
        server = "192.168.0.122"
    Dim port As String
        port = "13306"
    Dim database As String
        database = "test"
    Dim user As String
        user = "user15"
    Dim password As String
        password = "user15"
    Dim table As String
        table = "root_data"

    
    Dim mariaRS As ADODB.Recordset
    Dim excelRS As ADODB.Recordset
    Dim excelSQL As String
    
    ' mariaDB ����
    If Not openMariaDB(server, port, database, user, password) Then
        MsgBox "������DB ���ӿ� �����Ͽ����ϴ�.", vbCritical, "���"
        Exit Sub
    End If
    
    ' ���� ���� ����
    If Not openExcelFile(file_path, file_name) Then
        MsgBox file_name & " ���⿡ �����Ͽ����ϴ�.", vbCritical, "���"
        closeMariaDB
        Exit Sub
    End If
    
    'mariaDB ���ڵ弼Ʈ ��ü�� ����
    Set mariaRS = New ADODB.Recordset
    mariaRS.CursorLocation = adUseClient
    mariaRS.Open Source:=table, ActiveConnection:=mariaDBconn, CursorType:=adOpenForwardOnly, LockType:=adLockOptimistic, Options:=adCmdTable
    
    '���� ���� ���ڵ弼Ʈ ��ü�� ����
    Set excelRS = New ADODB.Recordset
    excelRS.CursorLocation = adUseClient
    excelSQL = "(SELECT `����`, `�����`, `��¥`, `�ð�`, `���`, `����`, `�ο�`, `�ݾ�`, `�μ�` FROM [" & sheet_name & "$]) AS `sub`"
    excelRS.Open Source:=excelSQL, ActiveConnection:=excelDBconn, CursorType:=adOpenForwardOnly, LockType:=adLockOptimistic, Options:=adCmdTable
    
    '���� ���� ù �ٺ��� ������ ������DB�� INSERT
    excelRS.MoveFirst
    Do While Not excelRS.EOF
        With mariaRS
            .AddNew
            For i = 1 To excelRS.Fields.count - 1
                .Fields(i) = excelRS.Fields(i)
            Next i
            .Fields(9) = Left(file_name, InStr(file_name, ".") - 1)
            .Update
        End With
        excelRS.MoveNext
    Loop
    
    '���� ���ڵ弼Ʈ ���� ���� �ݱ�
    Set excelRS = Nothing
    closeExcelFile
    '������DB ���ڵ弼Ʈ ���� DB �ݱ�
    Set mariaRS = Nothing
    closeMariaDB

End Sub

