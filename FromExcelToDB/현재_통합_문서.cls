VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "현재_통합_문서"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    thisPath = ThisWorkbook.path & "\raw data"
    Set sysObj = CreateObject("Scripting.FileSystemObject")
    Set thisFolder = sysObj.getFolder(thisPath)
    Set myFiles = New MyFile '모든 파일 및 시트 경로를 담은 array list 객체 초기화
    
    With Sheet1
        .ComboBox1.Clear
        .ComboBox1.AddItem pvargItem:="전체"
        .ComboBox1.Value = "전체"
        .ComboBox2.Clear
        .ComboBox2.AddItem pvargItem:="전체"
        .ComboBox2.Value = "전체"
        .ListBox1.Clear
        .ListBox2.Clear
    End With
    
    '루트 폴더 안 연도 폴더들
    For Each sub_year In thisFolder.subFolders
        '루트 폴더 안 연도 이름
        Dim sub_year_name As String
            sub_year_name = Mid(sub_year, InStrRev(sub_year, "\") + 1)
        '연도 콤보 박스에 입력
        Sheet1.ComboBox1.AddItem pvargItem:=sub_year_name
        '글로벌 연도 객체 선언 및 초기화
        Dim newYear As MyYear
            Set newYear = New MyYear
            newYear.name = sub_year_name
        '연도 폴더 안 파일들
        For Each sub_file In sub_year.Files
            '연도 폴더 안 파일 이름
            Dim sub_file_name As String
                sub_file_name = Mid(sub_file, InStrRev(sub_file, "\") + 1)
            
            If Mid(sub_file_name, InStrRev(sub_file_name, ".") + 1) = "xlsm" Then
                '지역 콤보 박스에 입력
                Sheet1.ListBox2.AddItem pvargItem:=sub_year_name & "\" & sub_file_name
                '글로벌 지역 객체 선언 및 초기화
                Dim newDistrict As MyDistrict
                    Set newDistrict = New MyDistrict
                    newDistrict.name = Left(sub_file_name, InStr(sub_file_name, ".") - 1)
                '파일 객체 열기
                Dim thisBook As Workbook
                    Set thisBook = GetObject(sub_file)
                '엑셀 파일 안 시트들
                For Each thisSheet In thisBook.Worksheets
                    If thisSheet.name <> "관리" Then
                        '글로벌 시트 객체 선언 및 초기화
                        Dim newSheet As MySheet
                            Set newSheet = New MySheet
                            newSheet.name = thisSheet.name
                        '지역 객체에 시트 객체 담기
                        newDistrict.addSheet newSheet
                        '시트 객체 소멸
                        Set newSheet = Nothing
                    End If
                Next thisSheet
                '연도 객체에 지역 객체 담기
                newYear.addDistrict newDistrict
                '지역 객체 소멸
                Set newDistrict = Nothing
                '파일 객체 닫기
                thisBook.Close SaveChanges:=False
                Set thisBook = Nothing
                Sheet1.Activate
            End If
            
        Next sub_file
        '연도 객체까지 담기
        myFiles.addYear newYear
        '연도 객체 소멸
        Set newYear = Nothing
    Next sub_year
    
    test_globalObect
    
End Sub

