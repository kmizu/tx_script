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

Function Evaluate(ByVal Node As Variant) As Variant
    Select Case Node.Tag
        Case ASTTag.TAG_BINARY_EXPRESSION
            Dim bNode As BinaryExpression
            Set bNode = Node
            Dim l, r As Variant
            Select Case bNode.OpType
                Case OpType.OpPlus
                    l = Evaluate(bNode.Lhs)
                    r = Evaluate(bNode.Rhs)
                    Evaluate = l + r
                Case OpType.OpMinus
                    l = Evaluate(bNode.Lhs)
                    r = Evaluate(bNode.Rhs)
                    Evaluate = l - r
                Case OpType.OpMultiply
                    l = Evaluate(bNode.Lhs)
                    r = Evaluate(bNode.Rhs)
                    Evaluate = l * r
                Case OpType.OpDivide
                    l = Evaluate(bNode.Lhs)
                    r = Evaluate(bNode.Rhs)
                    Evaluate = l / r
            End Select
        Case ASTTag.TAG_ASSIGNMENT
            Dim aNode As Assignment, result As Variant
            Set aNode = Node
            result = Evaluate(aNode.Expression)
            Environment_.Update aNode.name, result
            Debug.Print "Assign " & result & " To " & aNode.name
            Evaluate = result
        Case ASTTag.TAG_NAME
            Debug.Print "Name"
            Dim nNode As name, result2 As Variant
            Set nNode = Node
            result2 = Environment_.Lookup(nNode.Value)
            Debug.Print "Refer" & nNode.Value
            Evaluate = result2
        Case ASTTag.TAG_INT_NODE
            Evaluate = Node.Value
    End Select
End Function

Sub Initialize()
    Set Environment_ = Objects.NewEnvironment()
End Sub

