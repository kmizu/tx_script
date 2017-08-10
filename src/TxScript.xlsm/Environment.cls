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

Public Function Lookup(name As String) As Variant
    Lookup = Dictionary_.Item(name)
End Function

Public Sub Update(name As String, Value As Variant)
    Dictionary_.Add name, Value
End Sub

