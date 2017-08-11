Attribute VB_Name = "Exception"
Option Explicit
Sub Raise(ByVal Message As String)
    Error.Raise 515, Description:=Message
End Sub
