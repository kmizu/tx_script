VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Name"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Value_ As String

Public Property Get Tag() As Integer
    Tag = ASTTag.TAG_NAME
End Property

Public Property Get Value() As String
    Value = Value_
End Property

Public Sub Initialize(ByVal Value As String)
    Value_ = Value
End Sub
