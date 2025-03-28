VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyDistrict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_name As String
Private m_sheets As Object
'GETTER
Property Get name() As String
    Let name = m_name
End Property
'SETTER
Property Let name(ByVal newName As String)
    Let m_name = newName
End Property
'GETTER
Property Get Sheets() As Object
    Set Sheets = m_sheets
End Property
'SETTER
Property Set Sheets(ByRef newSheets As Object)
    Set m_sheets = newSheets
End Property
'CONSTRUCTOR
Private Sub Class_Initialize()
    Set m_sheets = CreateObject("System.Collections.ArrayList")
End Sub

Public Sub addSheet(ByVal sheet As MySheet)
    m_sheets.Add sheet
End Sub

Public Function count() As Integer
    count = m_sheets.count
End Function

Public Function getSheetByIndex(ByVal num As Integer) As MySheet

Try:
    On Error GoTo Catch
    
    If num < m_sheets.count Then
        Set getSheetByIndex = m_sheets.Item(num)
    Else
        Set getSheetByIndex = Nothing
    End If
    
    Exit Function
    
Catch:
    Set getSheetByIndex = Nothing
End Function

Public Function getSheetByName(ByVal name As String) As MySheet

Try:
    On Error GoTo Catch
    Set getSheetByName = Nothing
    For Each m_sheet In m_sheets
        If m_sheet.name = name Then
            Set getSheetByName = m_sheet
            Exit For
        End If
    Next m_sheet
    
    Exit Function
Catch:
    Set getSheetByName = Nothing
End Function
