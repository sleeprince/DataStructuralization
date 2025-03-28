VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_years As Object
'GETTER
Property Get Years() As Object
    Set Years = m_years
End Property
'SETTER
Property Set Years(ByRef newYears As Object)
    Set m_years = newYears
End Property
'CONSTRUCTOR
Private Sub Class_Initialize()
    Set m_years = CreateObject("System.Collections.ArrayList")
End Sub

Public Sub addYear(ByVal year As MyYear)
    m_years.Add year
End Sub

Public Function count() As Integer
    count = m_years.count
End Function

Public Function getYearByIndex(ByVal num As Integer) As MyYear

Try:
    On Error GoTo Catch
    
    If num < m_years.count Then
        Set getYearByIndex = m_years.Item(num)
    Else
        Set getYearByIndex = Nothing
    End If
    
    Exit Function
    
Catch:
    Set getYearByIndex = Nothing
End Function

Public Function getYearByName(ByVal name As String) As MyYear

Try:
    On Error GoTo Catch
    Set getYearByName = Nothing
    For Each m_year In m_years
        If m_year.name = name Then
            Set getYearByName = m_year
            Exit For
        End If
    Next m_year
    
    Exit Function
Catch:
    Set getYearByName = Nothing
End Function
