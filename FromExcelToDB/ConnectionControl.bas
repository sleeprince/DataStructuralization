Attribute VB_Name = "ConnectionControl"

Public Function openMariaDB(ByVal server As String, ByVal port As String, ByVal database As String, ByVal user As String, ByVal password As String) As Boolean
    
Try:
    On Error GoTo Catch
    
    Set mariaDBconn = New ADODB.Connection
    
    With mariaDBconn
        .ConnectionString = "Driver={MariaDB ODBC 3.2 Driver};" & _
                            "Server=" & server & ";" & _
                            "Port=" & port & ";" & _
                            "Database=" & database & ";" & _
                            "User=" & user & ";" & _
                            "Password=" & password & ";" & _
                            "Option=2;"
        .Open
    End With
    
    Debug.Print "YOUR ACCESS TO mariaDB SUCCEEDS"
    openMariaDB = True
        
    Exit Function
    
Catch:
    Debug.Print "YOUR ACCESS TO mariaDB FAILS"
    openMariaDB = False
    
End Function

Public Function openExcelFile(ByVal file_path As String, ByVal file_name As String) As Boolean

Try:
    On Error GoTo Catch
    
    Set excelDBconn = New ADODB.Connection
    
    With excelDBconn
        .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                            "Data Source=" & file_path & "\" & file_name & ";" & _
                            "Extended Properties=""Excel 12.0 xml;HDR=YES"""
        .Open
    End With
    
    Debug.Print "YOUR ACCESS TO " & file_name & " SUCCEEDS"
    openExcelFile = True
    
    Exit Function
    
Catch:

    Debug.Print "YOUR ACCESS TO " & file_name & " FAILS"
    openExcelFile = False
    
End Function

Public Function closeMariaDB() As Boolean

Try:
    On Error GoTo Catch
    
    mariaDBconn.Close
    Set mariaDBconn = Nothing
    
    Debug.Print "Closing mariaDB SUCCEEDS"
    closeMariaDB = True
    
    Exit Function
    
Catch:

    Set mariaDBconn = Nothing
    Debug.Print "Closing mariaDB Fails"
    closeMariaDB = False

End Function

Public Function closeExcelFile() As Boolean

Try:
    On Error GoTo Catch
    
    excelDBconn.Close
    Set excelDBconn = Nothing
    Debug.Print "Closing Excel File SUCCEEDS"
    closeExcelFile = True
    
    Exit Function
    
Catch:

    Set excelDBconn = Nothing
    Debug.Print "Closing Excel File FAILS"
    closeExcelFile = False
    
End Function

