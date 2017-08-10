VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Token"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IString

Private Value_ As Variant
Private TokenType_ As Integer

Public Property Get Value() As Variant
    Value = Value_
End Property

Public Property Get TokenType() As Integer
    TokenType = TokenType_
End Property

Public Function IString_ToString() As String
    IString_ToString = "Token(" & Value_ & ", " & ShowTokenType(TokenType_) & ")"
End Function

Public Function ToString() As String
    ToString = IString_ToString()
End Function

Public Sub Initialize(ByVal Value As Variant, ByVal TokenType As Integer)
    Value_ = Value
    TokenType_ = TokenType
End Sub
