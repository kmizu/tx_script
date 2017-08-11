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
