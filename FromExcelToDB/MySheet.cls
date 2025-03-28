VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MySheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_name As String
'GETTER
Property Get name() As String
    Let name = m_name
End Property
'SETTER
Property Let name(ByVal newName As String)
    Let m_name = newName
End Property
