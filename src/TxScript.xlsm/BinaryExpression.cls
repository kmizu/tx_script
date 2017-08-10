VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BinaryExpression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private OpType_ As Integer
Private Lhs_ As Object
Private Rhs_ As Object

Public Property Get Tag() As Integer
    Tag = ASTTag.TAG_BINARY_EXPRESSION
End Property

Public Property Get OpType() As Integer
    OpType = OpType_
End Property

Public Property Get Lhs() As Object
    Set Lhs = Lhs_
End Property

Public Property Get Rhs() As Object
    Set Rhs = Rhs_
End Property

Public Sub Initialize(ByVal OpType As Integer, ByVal Lhs As Object, ByVal Rhs As Object)
    OpType_ = OpType
    Set Lhs_ = Lhs
    Set Rhs_ = Rhs
End Sub

