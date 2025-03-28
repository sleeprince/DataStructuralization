VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_name As String
Private m_districts As Object
'GETTER
Property Get name() As String
    Let name = m_name
End Property
'SETTER
Property Let name(ByVal newName As String)
    Let m_name = newName
End Property
'GETTER
Property Get Districts() As Object
    Set Districts = m_districts
End Property
'SETTER
Property Set Districts(ByRef newDistricts As Object)
    Set m_districts = newDistricts
End Property
'CONSTRUCTOR
Private Sub Class_Initialize()
    Set m_districts = CreateObject("System.Collections.ArrayList")
End Sub

Public Sub addDistrict(ByVal district As MyDistrict)
    m_districts.Add district
End Sub

Public Function count() As Integer
    count = m_districts.count
End Function

Public Function getDistrictByIndex(ByVal num As Integer) As MyDistrict

Try:
    On Error GoTo Catch
    
    If num < m_districts.count Then
        Set getDistrictByIndex = m_districts.Item(num)
    Else
        Set getDistrictByIndex = Nothing
    End If
    
    Exit Function
    
Catch:
    Set getDistrictByIndex = Nothing
End Function

Public Function getDistrictByName(ByVal name As String) As MyDistrict

Try:
    On Error GoTo Catch
    Set getDistrictByName = Nothing
    For Each m_district In m_districts
        If m_district.name = name Then
            Set getDistrictByName = m_district
            Exit For
        End If
    Next m_district
    
    Exit Function
Catch:
    Set getDistrictByName = Nothing
End Function
