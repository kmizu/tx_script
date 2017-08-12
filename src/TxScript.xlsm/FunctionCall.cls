VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FunctionCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Name_ As String
Private Arg_ As Variant

Public Property Get Tag() As Integer
    Tag = ASTTag.TAG_FUNCTION_CALL
End Property

Public Property Get Name() As String
    Name = Name_
End Property

Public Property Get Arg() As Variant
    Set Arg = Arg_
End Property

Public Sub Initialize(Name As String, Arg As Variant)
    Name_ = Name
    Set Arg_ = Arg
End Sub

