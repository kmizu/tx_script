VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Assignment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Name_ As String
Private Expression_ As Variant

Public Property Get Tag() As Integer
    Tag = ASTTag.TAG_ASSIGNMENT
End Property

Public Property Get Name() As String
    Name = Name_
End Property

Public Property Get Expression() As Variant
    Set Expression = Expression_
End Property

Public Sub Initialize(Name As String, Expression As Variant)
    Name_ = Name
    Set Expression_ = Expression
End Sub

