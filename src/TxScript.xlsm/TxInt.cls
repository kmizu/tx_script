VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TxInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITxValue

Private Value_ As Integer

Public Property Get ITxValue_ValueTag() As Integer
    ITxValue_ValueTag = TxValueTag.TAG_INT
End Property

Public Property Get ValueTag() As Integer
    ValueTag = Me.ITxValue_ValueTag
End Property

Public Property Get Value() As Integer
    Value = Value_
End Property

Public Sub Initialize(ByVal Value As Integer)
    Value_ = Value
End Sub