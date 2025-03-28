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
    
    If openMariaDB(server, port, database, user, password) Then
    
        Set mariaRS = New ADODB.Recordset
        mariaRS.CursorLocation = adUseClient
        'mariaDB 레코드셑으로 열기
        mariaRS.Open Source:=table, ActiveConnection:=mariaDBconn, CursorType:=adOpenForwardOnly, LockType:=adLockOptimistic, Options:=adCmdTable
        
        If openExcelFile(file_path, file_name) Then
        
            Set excelRS = New ADODB.Recordset
            excelRS.CursorLocation = adUseClient
            
            excelSQL = "(SELECT `연번`, `사용자`, `날짜`, `시간`, `장소`, `목적`, `인원`, `금액`, `부서` FROM [" & sheet_name & "$]) AS `sub`"
            '엑셀 파일 레코드셑으로 열기
            excelRS.Open Source:=excelSQL, ActiveConnection:=excelDBconn, CursorType:=adOpenForwardOnly, LockType:=adLockOptimistic, Options:=adCmdTable
            
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
            
            Set excelRS = Nothing
            closeExcelFile
        Else
            MsgBox file_name & " 열기에 실패하였습니다.", vbCritical, "경고"
        End If
        
        Set mariaRS = Nothing
        closeMariaDB
    Else
        MsgBox "마리아DB 접속에 실패하였습니다.", vbCritical, "경고"
    End If

End Sub

