Attribute VB_Name = "TxValues"
Option Explicit

Function NewTxInt(Value As Integer) As TxInt
    Dim newObject As New TxInt
    newObject.Initialize Value
    Set NewTxInt = newObject
End Function

Function NewTxString(Value As String) As TxInt
    Dim newObject As New TxString
    newObject.Initialize Value
    Set NewTxString = newObject
End Function
