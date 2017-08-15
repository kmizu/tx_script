VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Evaluator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Environment_ As Environment

Function Evaluate(ByVal Node As Variant) As ITxValue
    Select Case Node.Tag
        Case ASTTag.TAG_BINARY_EXPRESSION
            Dim bNode As BinaryExpression
            Set bNode = Node
            Dim l, r As ITxValue
            Dim ltag, rtag As Integer
            Select Case bNode.OpType
                Case OpType.OpPlus
                    Set l = Evaluate(bNode.Lhs)
                    Set r = Evaluate(bNode.Rhs)
                    If l.ValueTag = TxValueTag.TAG_INT And r.ValueTag = TxValueTag.TAG_INT Then
                        Dim lInt As TxInt: Set lInt = l
                        Dim rInt As TxInt: Set rInt = r
                        Set Evaluate = TxValues.NewTxInt(lInt.Value + rInt.Value)
                    Else
                        Exception.Raise "incompatible types: operands must be Int"
                    End If
                Case OpType.OpMinus
                    Set l = Evaluate(bNode.Lhs)
                    Set r = Evaluate(bNode.Rhs)
                    If l.ValueTag = TxValueTag.TAG_INT And r.ValueTag = TxValueTag.TAG_INT Then
                        Set lInt = l
                        Set rInt = r
                        Set Evaluate = TxValues.NewTxInt(lInt.Value - rInt.Value)
                    Else
                        Exception.Raise "incompatible types: operands must be Int"
                    End If
                Case OpType.OpMultiply
                    Set l = Evaluate(bNode.Lhs)
                    Set r = Evaluate(bNode.Rhs)
                    If l.ValueTag = TxValueTag.TAG_INT And r.ValueTag = TxValueTag.TAG_INT Then
                        Set lInt = l
                        Set rInt = r
                        Set Evaluate = TxValues.NewTxInt(lInt.Value * rInt.Value)
                    Else
                        Exception.Raise "incompatible types: operands must be Int"
                    End If
                Case OpType.OpDivide
                    Set l = Evaluate(bNode.Lhs)
                    Set r = Evaluate(bNode.Rhs)
                    Evaluate = l / r
                    If l.ValueTag = TxValueTag.TAG_INT And r.ValueTag = TxValueTag.TAG_INT Then
                        Set lInt = l
                        Set rInt = r
                        Set Evaluate = TxValues.NewTxInt(lInt.Value / rInt.Value)
                    Else
                        Exception.Raise "incompatible types: operands must be Int"
                    End If
            End Select
        Case ASTTag.TAG_ASSIGNMENT
            Dim aNode As Assignment, result As ITxValue
            Set aNode = Node
            Set result = Evaluate(aNode.Expression)
            Environment_.Update aNode.Name, result
            Set Evaluate = result
        Case ASTTag.TAG_FUNCTION_CALL
            Dim fNode As FunctionCall, param As ITxValue
            Set fNode = Node
            Set param = Evaluate(fNode.Arg)
            Assert.IsEqualTo "println", fNode.Name 'Currently, only println function is supported
            Set Evaluate = TxValues.NewTxInt(1)
        Case ASTTag.TAG_NAME
            Dim nNode As Name
            Set nNode = Node
            Set Evaluate = Environment_.Lookup(nNode.Value)
        Case ASTTag.TAG_STRING_NODE
            Set Evaluate = TxValues.NewTxString(Node.Value)
        Case ASTTag.TAG_INT_NODE
            Set Evaluate = TxValues.NewTxInt(Node.Value)
    End Select
End Function

Sub Initialize()
    Set Environment_ = Objects.NewEnvironment()
End Sub

