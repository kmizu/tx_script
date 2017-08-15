VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TxString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITxValue

Private Value_ As String

Public Property Get ITxValue_ValueTag() As Integer
    ITxValue_ValueTag = TxValueTag.TAG_STRING
End Property

Public Property Get ValueTag() As Integer
    ValueTag = Me.ITxValue_ValueTag
End Property

Public Property Get Value() As Integer
    Value = Value_
End Property

Public Function ITxValue_ToString() As String
    ITxValue_ToString = Value_
End Function

Public Function Equals(Obj As Variant) As Boolean
    Dim that As TxInt
    If Obj.ValueTag = TxValueTag.TAG_STRING Then
        Set that = Obj
        Equals = Me.Value = that.Value
    Else
        Equals = False
    End If
End Function

Public Function ToString() As String
    ToString = ITxValue_ToString()
End Function

Public Sub Initialize(ByVal Value As String)
    Value_ = Value
End Sub

