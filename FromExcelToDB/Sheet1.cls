VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub ComboBox1_Change()
    
    Dim thisYear As String
        thisYear = Sheet1.ComboBox1.Value
    
    
    With Sheet1
        .ComboBox2.Clear
        .ComboBox2.AddItem pvargItem:="전체"
        .ComboBox2.Value = "전체"
        .ListBox1.Clear
        .ListBox2.Clear
    End With
    
    If thisYear = "전체" Then '콤보 박스가 전체일 때
        For Each sub_year In myFiles.Years
            For Each sub_district In sub_year.Districts
                    Sheet1.ListBox2.AddItem pvargItem:=sub_year.name & "\" & sub_district.name & ".xlsm"
            Next sub_district
        Next sub_year
    Else
        If Not myFiles.getYearByName(thisYear) Is Nothing Then '콤보 박스가 연도일 때
        
            For Each sub_district In myFiles.getYearByName(thisYear).Districts
            
                Sheet1.ComboBox2.AddItem pvargItem:=sub_district.name
                Sheet1.ListBox2.AddItem pvargItem:=sub_district.name & ".xlsm"
            
            Next sub_district
        End If
    End If
    
    
End Sub


Private Sub ComboBox2_Change()

    Dim thisYear As String
        thisYear = Sheet1.ComboBox1.Value
    
    Dim thisDistrict As String
        thisDistrict = Sheet1.ComboBox2.Value
        
    Dim thisBook As Workbook
    
    With Sheet1
        .ListBox1.Clear
        .ListBox2.Clear
    End With
    
    If thisYear = "전체" Then '연도 콤보 박스가 전체일 때
            
        If thisDistrict = "전체" Then '지역 콤보 박스가 전체일 때
            For Each sub_year In myFiles.Years
                For Each sub_district In sub_year.Districts
                    Sheet1.ListBox2.AddItem pvargItem:=sub_year.name & "\" & sub_district.name & ".xlsm"
                Next sub_district
            Next sub_year
        End If
            
    Else
        
        If thisDistrict = "전체" Then '지역 콤보 박스가 전체일 때
            If Not myFiles.getYearByName(thisYear) Is Nothing Then
                For Each sub_district In myFiles.getYearByName(thisYear).Districts
                    Sheet1.ListBox2.AddItem pvargItem:=sub_district.name & ".xlsm"
                Next
            End If
        Else
            If Not myFiles.getYearByName(thisYear) Is Nothing Then
                If Not myFiles.getYearByName(thisYear).getDistrictByName(thisDistrict) Is Nothing Then
                    For Each sub_sheet In myFiles.getYearByName(thisYear).getDistrictByName(thisDistrict).Sheets
                            Sheet1.ListBox2.AddItem pvargItem:=sub_sheet.name
                    Next sub_sheet
                End If
            End If
        End If
        
    End If
    
End Sub

Private Sub CommandButton1_Click()
    Dim text As String
    Dim index As Integer
        text = Sheet1.ListBox1.text
        index = Sheet1.ListBox1.ListIndex
    If index > -1 Then
        Sheet1.ListBox2.AddItem pvargItem:=text
        Sheet1.ListBox1.RemoveItem pvargIndex:=index
    Else
        MsgBox "올릴 것을 골라 주세요.", vbCritical, "오류!"
    End If
End Sub

Private Sub CommandButton2_Click()

    Dim text As String
    Dim index As Integer
        text = Sheet1.ListBox2.text
        index = Sheet1.ListBox2.ListIndex
    If index > -1 Then
        Sheet1.ListBox1.AddItem pvargItem:=text
        Sheet1.ListBox2.RemoveItem pvargIndex:=index
    Else
        MsgBox "올릴 것을 골라 주세요.", vbCritical, "오류!"
    End If
    
End Sub

Private Sub CommandButton3_Click()

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Dim startTime As Date
        startTime = Time()
    Dim endTime As Date
    Dim timeDiff As Long
    
    Dim thisYear As String
        thisYear = Sheet1.ComboBox1.Value
    
    Dim thisDistrict As String
        thisDistrict = Sheet1.ComboBox2.Value
    
    Dim file_name As String
    Dim in_year As String
    Dim in_district As String
    Dim in_sheet As String
    
    Dim path_string() As String
    
    If thisDistrict = "전체" Then
    
        For i = 0 To Sheet1.ListBox2.ListCount - 1
            If thisYear = "전체" Then
                path_string = Split(Sheet1.ListBox2.List(i), "\")
                in_year = path_string(0)
                file_name = path_string(1)
                in_district = Split(file_name, ".")(0)
            Else
                in_year = thisYear
                file_name = Sheet1.ListBox2.List(i)
                in_district = Split(file_name, ".")(0)
            End If
            'DB입력
            For Each sub_sheet In myFiles.getYearByName(in_year).getDistrictByName(in_district).Sheets
                insertFromExcelToMariaDB file_path:=thisPath & "\" & in_year, file_name:=file_name, sheet_name:=sub_sheet.name
            Next sub_sheet
        Next i
    Else
        For i = 0 To Sheet1.ListBox2.ListCount - 1
            in_year = thisYear
            file_name = thisDistrict & ".xlsm"
            in_district = thisDistrict
            in_sheet = Sheet1.ListBox2.List(i)
            'DB입력
            insertFromExcelToMariaDB file_path:=thisPath & "\" & in_year, file_name:=file_name, sheet_name:=in_sheet
        Next i
    End If
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    endTime = Time()
    timeDiff = DateDiff("s", startTime, endTime)
    MsgBox prompt:="일을 모두 마쳤습니다!" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "걸린 시간: " & timeDiff \ 60 & "분 " & timeDiff Mod 60 & "초", Buttons:=vbOKOnly Or vbInformation, title:="이것은 알림이오."
    
End Sub
