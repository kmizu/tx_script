VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Buffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Array_() As Variant

Public Sub Initialize()
    ReDim Array_(0)
End Sub

Public Sub Append(Element As Variant)
    ReDim Preserve Array_(UBound(Array_) + 1)
    If IsObject(Element) Then
        Set Array_(UBound(Array_)) = Element
    Else
        Array_(UBound(Array_)) = Element
    End If
End Sub

Public Function At(Index As Integer) As Variant
    If IsObject(Array_(Index)) Then
        Set At = Array_(Index)
    Else
        At = Array_(Index)
    End If
End Function

Public Property Get Length()
    Length = UBound(Array_)
End Property

