Attribute VB_Name = "AST"
Option Explicit

Function NewBinaryExpression(ByVal OpType As Integer, ByVal Lhs As Variant, ByVal Rhs As Variant) As BinaryExpression
    Dim newObject As New BinaryExpression
    newObject.Initialize OpType, Lhs, Rhs
    Set NewBinaryExpression = newObject
End Function

Function NewName(ByVal Value As String) As Name
    Dim newObject As New Name
    newObject.Initialize Value
    Set NewName = newObject
End Function

Function NewAssignment(ByVal Name As String, ByVal Expression As Variant) As Assignment
    Dim newObject As New Assignment
    newObject.Initialize Name, Expression
    Set NewAssignment = newObject
End Function

Function NewFunctionCall(ByVal Name As String, ByVal Arg As Variant) As FunctionCall
    Dim newObject As New FunctionCall
    newObject.Initialize Name, Arg
    Set NewFunctionCall = newObject
End Function

Function NewIntNode(ByVal Value As Integer) As IntNode
    Dim newObject As New IntNode
    newObject.Initialize Value
    Set NewIntNode = newObject
End Function
