Attribute VB_Name = "Exception"
Option Explicit
Sub Raise(ByVal message As String)
    Error.Raise 515, Description:=message
End Sub
