VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Parser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Tokenizer_ As Tokenizer

Public Sub Initialize(ByVal Tokenizer As Tokenizer)
    Set Tokenizer_ = Tokenizer
End Sub

Public Function Lines() As buffer
    Dim buffer As buffer
    Dim terminates As Boolean
    
    Set buffer = Objects.NewBuffer()
    terminates = False
    Do While Not terminates
        Dim la As Token
        Set la = Tokenizer_.Lookahead
        If la.TokenType = TokenType.TEOF Then
            terminates = True
        Else
            buffer.Append Line()
        End If
    Loop
    Set Lines = buffer
End Function

Public Function Line() As Variant
    Dim expr As Variant
    Set expr = Expression()
    Tokenizer_.Consume TokenType.TSemicolon
    Set Line = expr
End Function

Public Function Expression() As Variant
    Set Expression = Additive()
End Function

Public Function Additive() As Variant
    Dim result As Variant
    Set result = Multitive()
    Dim continuable As Boolean:  continuable = True
    Do While continuable
        Dim la As Token:  Set la = Tokenizer_.Lookahead()
        If la.TokenType = TokenType.TPlus Then
            Tokenizer_.Consume TokenType.TPlus
            Dim Lhs, Rhs As Variant
            Set Lhs = result
            Set Rhs = Multitive()
            Set result = AST.NewBinaryExpression(OpType.OpPlus, Lhs, Rhs)
        ElseIf la.TokenType = TokenType.TMinus Then
            Tokenizer_.Consume TokenType.TMinus
            Set Lhs = result
            Set Rhs = Multitive()
            Set result = AST.NewBinaryExpression(OpType.OpMinus, Lhs, Rhs)
        Else
            continuable = False
        End If
    Loop
    Set Additive = result
End Function

Public Function Multitive() As Variant
    Dim result As Variant
    Set result = Primary()
    Dim continuable As Boolean:  continuable = True
    Do While continuable
        Dim la As Token:  Set la = Tokenizer_.Lookahead()
        If la.TokenType = TokenType.TStar Then
            Tokenizer_.Consume TokenType.TStar
            Set result = AST.NewBinaryExpression(OpType.OpMultiply, result, Primary())
        ElseIf la.TokenType = TokenType.TSlash Then
            Tokenizer_.Consume TokenType.TSlash
            Set result = AST.NewBinaryExpression(OpType.OpDivide, result, Primary())
        Else
            continuable = False
        End If
    Loop
    Set Multitive = result
End Function

Public Function Primary() As Variant
    Dim result As Variant
    Dim la As Token:  Set la = Tokenizer_.Lookahead()
    If la.TokenType = TokenType.TLParen Then
        Tokenizer_.Consume TokenType.TLParen
        Set result = Expression()
        Tokenizer_.Consume TokenType.TRParen
    ElseIf la.TokenType = TokenType.TInt Then
        Tokenizer_.Consume TokenType.TInt
        Set result = AST.NewIntNode(la.Value)
    ElseIf la.TokenType = TokenType.TString Then
        Tokenizer_.Consume TokenType.TString
        Set result = AST.NewStringNode(la.Value)
    ElseIf la.TokenType = TokenType.TName Then
        Dim Name As String
        Name = la.Value
        Tokenizer_.Consume TokenType.TName
        Set la = Tokenizer_.Lookahead
        If la.TokenType = TokenType.TEQ Then
            Tokenizer_.Consume TokenType.TEQ
            Dim expr As Variant: Set expr = Expression()
            Set result = AST.NewAssignment(Name, expr)
        ElseIf la.TokenType = TokenType.TLParen Then
            Tokenizer_.Consume TokenType.TLParen
            Dim Arg As Variant: Set Arg = Expression()
            Tokenizer_.Consume TokenType.TRParen
            Set result = AST.NewFunctionCall(Name, Arg)
        Else
            Set result = AST.NewName(Name)
        End If
    Else
        Exception.Raise "( or Integer is expected"
    End If
    Set Primary = result
End Function
