VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Environment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Dictionary_ As Object

Public Sub Initialize()
    Set Dictionary_ = CreateObject("Scripting.Dictionary")
End Sub

Public Function Lookup(ByVal Name As String) As ITxValue
    Set Lookup = Dictionary_.Item(Name)
End Function

Public Sub Update(ByVal Name As String, ByVal Value As ITxValue)
    Dictionary_.Add Name, Value
End Sub

