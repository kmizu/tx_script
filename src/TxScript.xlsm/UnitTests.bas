Attribute VB_Name = "UnitTests"
Option Explicit
Sub Main()
  TestTokenizer
  TestEvaluator
  TestBuffer
  Assert.Done
End Sub

Sub TestTokenizer()
  Dim Tokenizer, Token 'Suppress automatic capitalize/decapitalize conversion

  Dim aTokenizer As Variant
  Dim aToken As Variant
  
  Set aTokenizer = Objects.NewTokenizer("abcde 123")
  Set aToken = aTokenizer.NextToken()
  Assert.IsEqualTo "abcde", aToken.Value
  Set aToken = aTokenizer.NextToken()
  Assert.IsEqualTo 123, aToken.Value
  
  Set aTokenizer = Objects.NewTokenizer("""hoge foo bar""")
  Set aToken = aTokenizer.NextToken()
  Assert.IsEqualTo "hoge foo bar", aToken.Value
  Assert.IsEqualTo TokenType.TString, aToken.TokenType
  
  Set aTokenizer = Objects.NewTokenizer("1+2")
  Set aToken = aTokenizer.NextToken()
  Assert.IsEqualTo 1, aToken.Value
  Assert.IsEqualTo TokenType.TInt, aToken.TokenType
  Set aToken = aTokenizer.NextToken()
  Assert.IsEqualTo "+", aToken.Value
  Assert.IsEqualTo TokenType.TPlus, aToken.TokenType
  Set aToken = aTokenizer.NextToken()
  Assert.IsEqualTo 2, aToken.Value
  Assert.IsEqualTo TokenType.TInt, aToken.TokenType
  
  Set aTokenizer = Objects.NewTokenizer("(")
  Set aToken = aTokenizer.NextToken()
  Assert.IsEqualTo "(", aToken.Value
  Assert.IsEqualTo TokenType.TLParen, aToken.TokenType
End Sub

Sub TestEvaluator()
    Dim aParser As Parser
    Dim anExpression As Variant
    Dim anEvaluator As Evaluator
    Dim aValue As Variant
    Dim expressions As buffer
    Dim i As Integer
    
    Set aParser = Objects.NewParser(Objects.NewTokenizer("a = (1 + 2) * 3; a;"))
    Set expressions = aParser.Lines()
    Set anEvaluator = Objects.NewEvaluator()
    For i = 1 To expressions.Length
        aValue = anEvaluator.Evaluate(expressions.At(i))
        Debug.Print aValue
    Next
    Assert.IsEqualTo 9, aValue
End Sub

Sub TestBuffer()
    Dim aBuffer As buffer
    Set aBuffer = Objects.NewBuffer()
    aBuffer.Append "A"
    aBuffer.Append "B"
    Assert.IsEqualTo "A", aBuffer.At(1)
    Assert.IsEqualTo "B", aBuffer.At(2)
End Sub
