VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IntNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Value_ As Integer

Public Property Get Tag() As Integer
    Tag = ASTTag.TAG_INT_NODE
End Property

Public Property Get Value() As Integer
    Value = Value_
End Property

Public Sub Initialize(ByVal Value As Integer)
    Value_ = Value
End Sub

