VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "����_����_����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

    ThisWorkbook.Unprotect ("")
    With Sheet1
        .Unprotect ("")
        .Activate
        With Cells(6, 1)
            .Value = "���� �ε� ��"
            .Font.Bold = True
        End With
    End With
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Start '���� �̸� �� ��� ��������
    
    Sheet1.ComboBox1.Clear
    Sheet1.ComboBox1.AddItem pvargItem:="��ü"
    
    Dim row, col As Integer
        row = 10
    Dim folerNm As String
    
    Dim data_dic As Object
    Set data_dic = CreateObject("Scripting.Dictionary") '���� ��ϵǾ� �ִ� �μ� ���
    
    While Sheet1.Cells(row, 1).Value <> ""
        data_dic.Add Sheet1.Cells(row, 1).Value, row
        row = row + 1
    Wend
    
    Debug.Print ("���� �μ�:")
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
        Debug.Print ("���� �μ�:")
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
    
    Starter.test
    
    Sheet1.ComboBox1.Text = "��ü"
    
    Sheet1.Cells(6, 1).Clear
    Sheet1.Protect ("")
    ThisWorkbook.Protect ("")
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
