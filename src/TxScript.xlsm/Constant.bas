Attribute VB_Name = "Constant"
Option Explicit
Enum TokenType
  TInt
TDouble
  TName
  TString
  TLParen
  TRParen
  TSemicolon
  TColon
  TPlus
  TMinus
  TStar
  TSlash
  TIF
  TELSE
  TDEF
  TEQ
  TLT
  TGT
  TLTE
  TGTE
  TEOF
End Enum

Enum ASTTag
    TAG_BINARY_EXPRESSION
    TAG_NAME
    TAG_ASSIGNMENT
    TAG_INT_NODE
End Enum

Enum OpType
  OpPlus
  OpMinus
  OpMultiply
  OpDivide
End Enum

Function ShowTokenType(ByVal TokenType As Integer) As String
    Dim result As String
    Select Case TokenType
        Case TInt
            result = "Int"
        Case TDouble
            result = "Double"
        Case TName
            result = "Name"
        Case TString
            result = "String"
        Case TLParen
            result = "("
        Case TRParen
            result = ")"
        Case TSemicolon
            result = ";"
        Case TColon
            result = ":"
        Case TPlus
            result = "+"
        Case TMinus
            result = "-"
        Case TStar
            result = "*"
        Case TSlash
            result = "/"
        Case TEOF
            result = "EOF"
        Case Else
            result = "Unknown"
    End Select
    ShowTokenType = result
End Function
