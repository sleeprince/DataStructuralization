Attribute VB_Name = "Test"
Sub test1()
    getTable query_name:="test table", file_name:=folder_path & "\" & "23�� 4�� ��ȹ����� ���������� ��볻��.pdf", nth:=1
    makeSheetWithTable sheet_name:="sheet_test", table_name:="test table"
    'mkFile name:="test", path:=folder_path
'    getNumOfTable query_name:="numOfTable", file_name:=folder_path & "\" & "23�� 4�� ��ȹ����� ���������� ��볻��.pdf"
'    makeSheetWithTable sheet_name:="numOfTable", table_name:="numOfTable"
'    deleteQuery query_name:="numOfTable"
End Sub
