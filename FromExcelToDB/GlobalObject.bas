Attribute VB_Name = "GlobalObject"
Public mariaDBconn As ADODB.Connection
Public excelDBconn As ADODB.Connection
Public sysObj As Object '시스템 오브젝트
Public thisFolder As Object 'raw data의 root 객체
Public thisPath As String '이 파일의 경로
Public myFiles As MyFile '모든 파일 및 시트 경로를 담은 array list 객체 선언
