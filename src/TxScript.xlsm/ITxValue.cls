VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITxValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get ValueTag() As Integer
    ' Implemented by subclasses
End Property

Public Function ToString() As String
    ' Implemented by subclasses
End Function

