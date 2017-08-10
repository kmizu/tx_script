Attribute VB_Name = "Objects"
Function NewTokenizer(ByVal Source As String) As Tokenizer
    Dim newObject As New Tokenizer
    newObject.Initialize Source
    Set NewTokenizer = newObject
End Function

Function NewToken(ByVal Value As Variant, ByVal TokenType As Integer) As Token
    Dim newObject As New Token
    newObject.Initialize Value, TokenType
    Set NewToken = newObject
End Function

Function NewParser(ByVal Tokenizer As Tokenizer) As Parser
    Dim newObject As New Parser
    newObject.Initialize Tokenizer
    Set NewParser = newObject
End Function

Function NewBinaryExpression(ByVal OpType As Integer, ByVal Lhs As Variant, ByVal Rhs As Variant) As BinaryExpression
    Dim newObject As New BinaryExpression
    newObject.Initialize OpType, Lhs, Rhs
    Set NewBinaryExpression = newObject
End Function

Function NewName(ByVal Value As String) As name
    Dim newObject As New name
    newObject.Initialize Value
    Set NewName = newObject
End Function

Function NewAssignment(ByVal name As String, ByVal Expression As Variant) As Assignment
    Dim newObject As New Assignment
    newObject.Initialize name, Expression
    Set NewAssignment = newObject
End Function

Function NewIntNode(ByVal Value As Integer) As IntNode
    Dim newObject As New IntNode
    newObject.Initialize Value
    Set NewIntNode = newObject
End Function

Function NewEvaluator() As Evaluator
    Dim newObject As New Evaluator
    newObject.Initialize
    Set NewEvaluator = newObject
End Function

Function NewEnvironment() As Environment
    Dim newObject As New Environment
    newObject.Initialize
    Set NewEnvironment = newObject
End Function

Function NewBuffer() As buffer
    Dim newObject As New buffer
    newObject.Initialize
    Set NewBuffer = newObject
End Function
