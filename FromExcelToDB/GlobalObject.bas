Attribute VB_Name = "GlobalObject"
Public mariaDBconn As ADODB.Connection
Public excelDBconn As ADODB.Connection
Public sysObj As Object '�ý��� ������Ʈ
Public thisFolder As Object 'raw data�� root ��ü
Public thisPath As String '�� ������ ���
Public myFiles As MyFile '��� ���� �� ��Ʈ ��θ� ���� array list ��ü ����
