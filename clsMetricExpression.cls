VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMetricExpression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Text As String
Private m_Pos As Long
Private m_Engine As clsMeasureEngine
Private m_Date As String
Private m_CheckOnly As Boolean
Private m_References As Object
Private m_Depth As Long

Public Function Calculate(ByVal expression As String, ByVal engine As clsMeasureEngine, ByVal reportDate As String) As Double
    Set m_Engine = engine
    m_Date = reportDate
    m_CheckOnly = False
    m_Engine.ValidateContext reportDate
    Prepare expression
    Calculate = ParseSum(m_Date)
    Finish
End Function

Public Function References(ByVal expression As String) As Object

    Dim ignored As Double
    m_CheckOnly = True
    Set m_References = CreateObject("Scripting.Dictionary")
    Prepare expression
    ignored = ParseSum("")
    Finish
    Set References = m_References

End Function

Private Sub Prepare(ByVal expression As String)
    m_Text = UCase$(expression)
    If Len(m_Text) = 0 Or Len(m_Text) > 4096 Then Invalid "Empty or oversized expression."
    m_Pos = 1
    m_Depth = 0
End Sub

Private Sub Finish()
    SkipSpaces
    If m_Pos <= Len(m_Text) Then Invalid "Unexpected token."
End Sub

Private Sub SkipSpaces()
    Dim ch As String
    Do While m_Pos <= Len(m_Text)
        ch = Mid$(m_Text, m_Pos, 1)
        If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then Exit Do
        m_Pos = m_Pos + 1
    Loop
End Sub

Private Function Peek() As String
    SkipSpaces
    Peek = Mid$(m_Text, m_Pos, 1)
End Function

Private Function ParseSum(ByVal reportDate As String) As Double
    Dim result As Double, operand As Double, op As String
    result = ParseProduct(reportDate)
    Do
        op = Peek()
        If op <> "+" And op <> "-" Then Exit Do
        m_Pos = m_Pos + 1
        operand = ParseProduct(reportDate)
        If op = "+" Then
            result = result + operand
        Else
            result = result - operand
        End If
    Loop
    ParseSum = result
End Function

Private Function ParseProduct(ByVal reportDate As String) As Double
    Dim result As Double, operand As Double, op As String
    result = ParseAtom(reportDate)
    Do
        op = Peek()
        If op <> "*" And op <> "/" Then Exit Do
        m_Pos = m_Pos + 1
        operand = ParseAtom(reportDate)
        If Not m_CheckOnly Then
            If op = "*" Then
                result = result * operand
            Else
                If Abs(operand) < 0.0000001 Then Err.Raise 11, "clsMetricExpression", "Zero denominator."
                result = result / operand
            End If
        End If
    Loop
    ParseProduct = result
End Function

Private Function ParseAtom(ByVal reportDate As String) As Double
    Dim ch As String, token As String, start As Long, result As Double
    Dim argumentStart As Long, argumentEnd As Long, priorValue As Double
    m_Depth = m_Depth + 1
    If m_Depth > 100 Then Invalid "Expression nesting exceeds 100."
    ch = Peek()
    Select Case ch
        Case "+", "-"
            m_Pos = m_Pos + 1
            result = ParseAtom(reportDate)
            If ch = "-" Then result = -result
        Case "("
            m_Pos = m_Pos + 1
            result = ParseSum(reportDate)
            RequireClose
        Case "0" To "9", "."
            start = m_Pos
            Do While Mid$(m_Text, m_Pos, 1) Like "[0-9]"
                m_Pos = m_Pos + 1
            Loop
            If Mid$(m_Text, m_Pos, 1) = "." Then
                m_Pos = m_Pos + 1
                Do While Mid$(m_Text, m_Pos, 1) Like "[0-9]"
                    m_Pos = m_Pos + 1
                Loop
            End If
            token = Mid$(m_Text, start, m_Pos - start)
            If token = "." Then Invalid "Invalid number."
            result = val(token)
        Case "A" To "Z", "_"
            start = m_Pos
            Do While Mid$(m_Text, m_Pos, 1) Like "[A-Z0-9_]"
                m_Pos = m_Pos + 1
            Loop
            token = Mid$(m_Text, start, m_Pos - start)
            If Peek() = "(" Then
                If token <> "AVG" And token <> "ABS" Then Invalid "Unsupported function: " & token
                m_Pos = m_Pos + 1
                argumentStart = m_Pos
                result = ParseSum(reportDate)
                argumentEnd = m_Pos
                RequireClose
                If token = "ABS" Then
                    result = Abs(result)
                ElseIf Not m_CheckOnly Then
                    priorValue = ComparableArgument(reportDate, argumentStart, argumentEnd)
                    result = (result + priorValue) / 2#
                End If
            ElseIf m_CheckOnly Then
                m_References(token) = True
                result = 1#
            Else
                result = m_Engine.ResolveValue(token, reportDate)
            End If
        Case Else
            Invalid "Expected number, identifier or parenthesis."
    End Select
    m_Depth = m_Depth - 1
    ParseAtom = result
End Function

Private Function ComparableArgument(ByVal reportDate As String, ByVal argumentStart As Long, ByVal argumentEnd As Long) As Double
    Dim priorDate As String, result As Double, pushed As Boolean
    Dim savedPos As Long, savedDepth As Long, n As Long, description As String, source As String
    savedPos = m_Pos
    savedDepth = m_Depth
    On Error GoTo Failed
    priorDate = Format$(DateAdd("yyyy", -1, CDate(reportDate)), "yyyy-mm-dd")
    m_Engine.PushComparable priorDate
    pushed = True
    m_Engine.ValidateContext priorDate
    m_Pos = argumentStart
    result = ParseSum(priorDate)
    If m_Pos <> argumentEnd Then Invalid "Invalid AVG argument."
    RequireClose
    m_Engine.PopContext
    pushed = False
    m_Pos = savedPos
    m_Depth = savedDepth
    ComparableArgument = result
    Exit Function
Failed:
    n = Err.number
    description = Err.description
    source = Err.source
    If pushed Then m_Engine.PopContext
    m_Pos = savedPos
    m_Depth = savedDepth
    Err.Raise n, source, description
End Function

Private Sub RequireClose()
    If Peek() <> ")" Then Invalid "Expected closing parenthesis."
    m_Pos = m_Pos + 1
End Sub

Private Sub Invalid(ByVal description As String)
    Err.Raise vbObjectError + 2300, "clsMetricExpression", description & " Position " & CStr(m_Pos) & ": " & m_Text
End Sub
