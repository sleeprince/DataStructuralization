VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click() '�ٽ� �� ��������
    
    Dim startTime As Date
        startTime = Time()
    Dim endTime As Date
    Dim timeDiff As Long
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    ThisWorkbook.Unprotect ("")
    Sheet1.Unprotect ("")
    
    Dim row, col As Integer '����ȣ
    Dim res '��ư ��������
    Dim combo As String
        combo = ComboBox1.Value
    
    If combo = "��ü" Then
        
        res = MsgBox(prompt:="������ �Ͻ÷�����?" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "���� �ɸ� �ٵ���.", _
                Buttons:=vbYesNo Or vbCritical Or vbDefaultButton2, _
                Title:="�̰��� �����")
        If res = vbYes Then
            Debug.Print ("��ü �ٽ� �������� ����")
            '�ʱ�ȭ
            row = 10
            While Sheet1.Cells(row, 1).Value <> ""
                deleteQuery (Sheet1.Cells(row, 1).Value)
                deleteSheet (Sheet1.Cells(row, 1).Value)
                numOfFile_dic.Remove (Sheet1.Cells(row, 1).Value)
                col = 2
                While Sheet1.Cells(row, col).Value <> ""
                    deleteQuery (Sheet1.Cells(row, 1) & (col - 1))
                    file_dic.Remove (Sheet1.Cells(row, col).Value)
                    Sheet1.Cells(row, col).Clear
                    col = col + 1
                Wend
                row = row + 1
            Wend
            GetAllFiles '��� ��������
        Else
            Debug.Print ("��ü �������� ���")
        End If
        
    Else
        
        res = MsgBox(prompt:="�ִ� �����͸� ����ϴ�." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "�׷��� �Ͻðڽ��ϱ�?", _
                Buttons:=vbYesNo Or vbCritical Or vbDefaultButton2, _
                Title:="�̰��� �����")
        If res = vbYes Then
            Debug.Print ("�ٽ� �������� ����")
            '�ʱ�ȭ
            numOfFile_dic.Remove (combo)
            col = 2
            While Sheet1.Cells(department_dic.Item(combo), col).Value <> ""
                deleteQuery (combo & (col - 1))
                Debug.Print ("���� ���� �̸�: " & combo & (col - 1))
                Debug.Print ("���� ��ųʸ� Ű: " & Sheet1.Cells(department_dic.Item(combo), col).Value)
                file_dic.Remove (Sheet1.Cells(department_dic.Item(combo), col).Value)
                Sheet1.Cells(department_dic.Item(combo), col).Clear
                col = col + 1
            Wend
            
            '���� �ҷ�����
            GetFiles combo, department_dic.Item(combo)
        Else
            Debug.Print ("�������� ���")
        End If
        
    End If
    
    Sheet1.Activate
    Sheet1.Protect ("")
    ThisWorkbook.Protect ("")
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    endTime = Time()
    timeDiff = DateDiff("s", startTime, endTime)
    MsgBox prompt:="���� ��� ���ƽ��ϴ�!" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "�ɸ� �ð�: " & timeDiff \ 60 & "�� " & timeDiff Mod 60 & "��", Buttons:=vbOKOnly Or vbInformation, Title:="�̰��� �˸��̿�."
    Debug.Print
    
End Sub

Private Sub CommandButton2_Click() '���͸� ��������

    Dim startTime As Date
        startTime = Time()
    Dim endTime As Date
    Dim timeDiff As Long
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    ThisWorkbook.Unprotect ("")
    Sheet1.Unprotect ("")
    
    Dim combo As String
        combo = Sheet1.ComboBox1.Value
    
    If combo = "��ü" Then
        AddAllFiles
    Else
        AddFiles combo, department_dic.Item(combo)
    End If
        
    Sheet1.Activate
    Sheet1.Protect ("")
    ThisWorkbook.Protect ("")
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    endTime = Time()
    timeDiff = DateDiff("s", startTime, endTime)
    MsgBox prompt:="���� ��� ���ƽ��ϴ�!" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "�ɸ� �ð�: " & timeDiff \ 60 & "�� " & timeDiff Mod 60 & "��", Buttons:=vbOKOnly Or vbInformation, Title:="�̰��� �˸��̿�."
    
End Sub
