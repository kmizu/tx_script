VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tokenizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Source_ As String
Private Position_ As Integer
Private LA_Token_ As Token

Public Property Get Source() As String
    Source = Source_
End Property

Public Property Get Position() As Integer
    Position = Position_
End Property

Public Property Get Rest() As String
    Rest = Mid(Source_, Position_)
End Property

Private Function TokenizeIdentifier() As Token
    Dim start As Integer:  start = Position_
    Dim current As Integer: current = start
    Dim found As Boolean:  found = False
    Do While Not found
        If Mid(Source_, current, 1) Like "[a-zA-Z0-9_]" Then
            current = current + 1
        Else
            found = True
        End If
    Loop
    Dim id As String: id = Mid(Source_, start, (current - start))
    Dim result As Token
    If id = "if" Then
        Set result = Objects.NewToken("if", TokenType.TIF)
    ElseIf id = "def" Then
        Set result = Objects.NewToken("def", TokenType.TDEF)
    ElseIf id = "else" Then
        Set result = Objects.NewToken("else", TokenType.TELSE)
    Else
        Set result = Objects.NewToken(id, TokenType.TName)
    End If
    Position_ = current
    Set TokenizeIdentifier = result
End Function

Private Function TokenizeInteger() As Token
    Dim start As Integer:  start = Position_
    Dim current As Integer: current = start
    Dim found As Boolean:  found = False
    Dim result As Integer: result = 0
    Do While Not found
        Dim ch As String:  ch = Mid(Source_, current, 1)
        If ch Like "[0-9]" Then
            result = result * 10 + CInt(ch)
            current = current + 1
        Else
            found = True
        End If
    Loop
    Position_ = current
    Set TokenizeInteger = Objects.NewToken(result, TokenType.TInt)
End Function

Private Function TokenizeString() As Token
    Dim ch As String:  ch = Mid(Source_, Position_, 1)
    If Not ch = """" Then
      Exception.Raise "first character must be """""
    End If
    Position_ = Position_ + 1
    Dim start As Integer:  start = Position_
    Dim current As Integer: current = start
    Dim terminated As Boolean:  terminated = False
    Dim result As String: result = ""
    Do While Not terminated
        ch = Mid(Source_, current, 1)
        If ch = "€" Then
            current = current + 1
            ch = Mid(Source_, current, 1)
            Select Case ch
                Case "r"
                    result = result & vbCr
                Case "n"
                    result = result & vbLf
                Case """"
                    result = result & """"
                Case "€"
                    result = result & "€"
                Case Else
                    Exception.Raise "unexpected escape sequence"
            End Select
            current = current + 1
        ElseIf ch = """" Then
            terminated = True
            current = current + 1
        Else
            result = result & ch
            current = current + 1
        End If
    Loop
    Position_ = current
    Set TokenizeString = Objects.NewToken(result, TokenType.TString)
End Function

Private Sub SkipWhitespaces()
    Dim isWhitespace As Boolean:   isWhitespace = True
    Do While isWhitespace
        If Len(Mid(Source_, Position_)) = 0 Then
            isWhitespace = False
        Else
            Dim ch As String:  ch = Mid(Source_, Position_, 1)
            If ch Like "[ ]" Then
                Position_ = Position_ + 1
            Else
                isWhitespace = False
            End If
        End If
    Loop
End Sub

Public Sub Consume(TokenType As Integer)
    Dim t As Token
    If LA_Token_ Is Nothing Then
        Set t = Me.NextToken()
    Else
        Set t = LA_Token_
    End If
    If Not t.TokenType = TokenType Then
        Exception.Raise "expected :" & ShowTokenType(t.TokenType)
    End If
    Set LA_Token_ = Nothing
End Sub

Public Function Lookahead() As Token
    If LA_Token_ Is Nothing Then
        Set LA_Token_ = Me.NextToken()
        Set Lookahead = LA_Token_
    Else
        Set Lookahead = LA_Token_
    End If
End Function

Public Function NextToken() As Token
    Set LA_Token_ = Nothing
    SkipWhitespaces
    Dim result As Token
    If Position > Len(Source_) Then
        Set result = Objects.NewToken("", TokenType.TEOF)
    Else
        Dim first As String: first = Mid(Source_, Position_, 1)
        If first Like "[0-9]" Then
            Set result = TokenizeInteger()
        ElseIf first = """" Then
            Set result = TokenizeString()
        ElseIf first Like "[a-zA-Z_]" Then
            Set result = TokenizeIdentifier()
        ElseIf first = "(" Then
            Set result = Objects.NewToken("(", TokenType.TLParen)
            Position_ = Position_ + 1
        ElseIf first = ")" Then
            Set result = Objects.NewToken(")", TokenType.TRParen)
            Position_ = Position_ + 1
        ElseIf first = "=" Then
            Set result = Objects.NewToken("=", TokenType.TEQ)
            Position_ = Position_ + 1
        ElseIf first = "<" Then
            Set result = Objects.NewToken("=", TokenType.TLT)
            Position_ = Position_ + 1
        ElseIf first = ">" Then
            Set result = Objects.NewToken("=", TokenType.TGT)
            Position_ = Position_ + 1
        ElseIf first = ";" Then
            Set result = Objects.NewToken(")", TokenType.TSemicolon)
            Position_ = Position_ + 1
        ElseIf first = ":" Then
            Set result = Objects.NewToken(":", TokenType.TColon)
            Position_ = Position_ + 1
        ElseIf first = "=" Then
            Set result = Objects.NewToken("=", TokenType.TEQ)
            Position_ = Position_ + 1
        ElseIf first = "+" Then
            Set result = Objects.NewToken("+", TokenType.TPlus)
            Position_ = Position_ + 1
        ElseIf first = "-" Then
            Set result = Objects.NewToken("-", TokenType.TMinus)
            Position_ = Position_ + 1
        ElseIf first = "*" Then
            Set result = Objects.NewToken("*", TokenType.TStar)
            Position_ = Position_ + 1
        ElseIf first = "/" Then
            Set result = Objects.NewToken("/", TokenType.TSlash)
            Position_ = Position_ + 1
        End If
    End If
    SkipWhitespaces
    Set NextToken = result
End Function

Public Sub Initialize(ByVal Source As String)
    Source_ = Source
    Position_ = 1
    Set LA_Token_ = Nothing
End Sub

