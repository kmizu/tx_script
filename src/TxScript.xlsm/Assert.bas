Attribute VB_Name = "Assert"
Option Explicit

Sub IsEqualTo(expected As Variant, actual As Variant)
    If expected <> actual Then
        Err.Raise 513, Description:="Expected: " & expected & " Actual: " & actual
    End If
End Sub

Sub IsTrue(expectedTrue As Boolean)
    If Not expectedTrue Then
        Err.Raise 513, Description:="Assertion failed"
    End If
End Sub

Sub Done(Message As String)
  MsgBox Message
End Sub
