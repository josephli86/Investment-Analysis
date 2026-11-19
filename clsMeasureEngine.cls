VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMeasureEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Definitions As Object
Private m_Parameters As Object
Private m_Active As Object
Private m_Ctx As clsEvalContext
Private m_LastError As String

Private m_Audit As Boolean
Private m_AuditRows As Collection
Private m_Parent As Long
Private m_Level As Long

Private Sub Class_Initialize()
    Set m_Definitions = CreateObject("Scripting.Dictionary")
    Set m_Parameters = CreateObject("Scripting.Dictionary")
End Sub

Public Property Get LastError() As String
    LastError = m_LastError
End Property

Public Property Get DefinitionCount() As Long
    DefinitionCount = m_Definitions.count
End Property

Private Function textValue(ByVal Value As Variant) As String
    If Not IsNull(Value) Then textValue = Trim$(CStr(Value))
End Function

Private Function ValidCode(ByVal code As String) As Boolean
    Dim i As Long
    If Len(code) = 0 Or Len(code) > 100 Then Exit Function
    If Not Left$(code, 1) Like "[A-Z_]" Then Exit Function
    For i = 1 To Len(code)
        If Not Mid$(code, i, 1) Like "[A-Z0-9_]" Then Exit Function
    Next i
    ValidCode = True
End Function

Public Sub LoadFile(ByVal filePath As String, ByVal parameterPath As String)
    Dim conn As Object, rs As Object, rows As Variant
    Dim n As Long, description As String
    On Error GoTo Failed
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Mode=Read;Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
    Set rs = conn.Execute("SELECT CALC_ORDER,METRIC_CODE,FORMULA_TYPE,NUMERATOR,DENOMINATOR,ENABLED FROM [MeasuresDef$]")
    If rs.EOF Then Err.Raise 5, , "MeasuresDef contains no definitions."
    rows = rs.GetRows()
    rs.Close
    conn.Close
    LoadRows rows
    LoadParameters parameterPath
    ValidateDefinitions
    Exit Sub
Failed:
    n = Err.number
    description = Err.description
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
    End If
    If Not conn Is Nothing Then
        If conn.State <> 0 Then conn.Close
    End If
    On Error GoTo 0
    m_LastError = description
    Err.Raise n, "clsMeasureEngine.LoadFile", description
End Sub

Public Sub LoadRows(ByVal rows As Variant)
    Dim i As Long, code As String, kind As String, numerator As String, denominator As String
    Dim enabled As String, order As Variant
    Set m_Definitions = CreateObject("Scripting.Dictionary")
    Set m_Parameters = CreateObject("Scripting.Dictionary")
    For i = LBound(rows, 2) To UBound(rows, 2)
        code = UCase$(textValue(rows(1, i)))
        If Len(code) > 0 Then
            If Not ValidCode(code) Then Err.Raise 5, , "Invalid metric code: " & code
            If code = "TAX_RATE" Then Err.Raise 5, , "TAX_RATE is a built-in version account, not a metric definition."
            If m_Definitions.exists(code) Then Err.Raise 5, , "Duplicate metric code: " & code
            order = rows(0, i)
            If Not IsNumeric(order) Then Err.Raise 5, , "CALC_ORDER must be numeric: " & code
            enabled = UCase$(textValue(rows(5, i)))
            If enabled <> "Y" And enabled <> "N" Then Err.Raise 5, , "ENABLED must be Y or N: " & code
            kind = UCase$(textValue(rows(2, i)))
            numerator = textValue(rows(3, i))
            denominator = textValue(rows(4, i))
            m_Definitions.Add code, Array(kind, numerator, denominator, enabled = "Y", CDbl(order))
        ElseIf Len(textValue(rows(2, i))) > 0 Or Len(textValue(rows(3, i))) > 0 Then
            Err.Raise 5, , "Metric code is required at data row " & CStr(i + 1)
        End If
    Next i
End Sub

Public Sub SetParameter(ByVal code As String, ByVal Value As Double)
    code = UCase$(Trim$(code))
    If Not ValidCode(code) Then Err.Raise 5, , "Invalid parameter: " & code
    If code = "TAX_RATE" Then Err.Raise 5, , "TAX_RATE is a built-in version account, not a parameter."
    If m_Definitions.exists(code) Then Err.Raise 5, , "Parameter conflicts with metric: " & code
    m_Parameters(code) = Value
End Sub

Private Sub LoadParameters(ByVal filePath As String)
    Dim fso As Object, stream As Object, line As String, parts As Variant, code As String
    Dim rx As Object, n As Long, description As String
    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then Exit Sub
    Set stream = fso.OpenTextFile(filePath, 1, False)
    Set rx = CreateObject("VBScript.RegExp")
    rx.Pattern = "^[+-]?([0-9]+(\.[0-9]*)?|\.[0-9]+)$"
    Do While Not stream.AtEndOfStream
        line = Trim$(stream.ReadLine)
        If Len(line) > 0 And Left$(line, 1) <> ";" And Left$(line, 1) <> "#" Then
            parts = Split(line, "=", 2)
            code = UCase$(Trim$(parts(0)))
            If code <> "TAX_RATE" Then
                If UBound(parts) <> 1 Then Err.Raise 5, , "Expected parameter=value: " & line
                If m_Parameters.exists(code) Then Err.Raise 5, , "Duplicate parameter: " & code
                If Not rx.Test(Trim$(parts(1))) Then Err.Raise 5, , "Invalid parameter value: " & line
                SetParameter code, val(Trim$(parts(1)))
            End If
        End If
    Loop
    stream.Close
    Exit Sub
Failed:
    n = Err.number
    description = Err.description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Err.Raise n, "clsMeasureEngine.LoadParameters", description
End Sub

Public Sub ValidateDefinitions()
    Dim states As Object, code As Variant
    Set states = CreateObject("Scripting.Dictionary")
    For Each code In m_Definitions.Keys
        ValidateOne CStr(code), states, 0
    Next code
End Sub

Private Sub ValidateOne(ByVal code As String, ByVal states As Object, ByVal depth As Long)
    Dim d As Variant, refs As Object, parser As clsMetricExpression
    If depth > 100 Then Err.Raise 5, , "Metric dependency depth exceeds 100: " & code
    If states.exists(code) Then
        If states(code) = 1 Then Err.Raise 5, , "Circular metric dependency: " & code
        Exit Sub
    End If
    d = m_Definitions(code)
    If Not d(3) Then Exit Sub
    states(code) = 1
    Select Case d(0)
        Case "DEFINED", "GROWTH"
            If Len(d(2)) > 0 Then Err.Raise 5, , "Unexpected denominator: " & code
        Case "RATIO"
            If Len(d(2)) = 0 Then Err.Raise 5, , "RATIO requires denominator: " & code
        Case Else
            Err.Raise 5, , "Unsupported FORMULA_TYPE: " & code
    End Select
    Set parser = New clsMetricExpression
    Set refs = parser.References(CStr(d(1)))
    ValidateReferences refs, states, depth, code
    If d(0) = "RATIO" Then
        Set refs = parser.References(CStr(d(2)))
        ValidateReferences refs, states, depth, code
    End If
    states(code) = 2
End Sub

Private Sub ValidateReferences(ByVal refs As Object, ByVal states As Object, ByVal depth As Long, ByVal owner As String)
    Dim ref As Variant, d As Variant
    For Each ref In refs.Keys
        If CStr(ref) = "TAX_RATE" Then
        ElseIf m_Definitions.exists(ref) Then
            d = m_Definitions(ref)
            If Not d(3) Then Err.Raise 5, , owner & " references disabled metric: " & CStr(ref)
            ValidateOne CStr(ref), states, depth + 1
        End If
    Next ref
End Sub

Public Function Evaluate(ByVal stmt As clsFinancialStatement, ByVal rptDate As String, ByVal metricCode As String, Optional ByVal ifAdjusted As Integer = 2) As Variant
    Dim ctx As clsEvalContext, provider As clsStatementProvider
    Set provider = New clsStatementProvider
    provider.Initialize stmt
    Set ctx = New clsEvalContext
    ctx.RegisterProvider provider
    Evaluate = EvaluateContext(ctx, rptDate, metricCode, ifAdjusted)
End Function

Public Function EvaluateContext(ByVal ctx As clsEvalContext, ByVal rptDate As String, ByVal metricCode As String, Optional ByVal ifAdjusted As Integer = 2) As Variant
    Dim n As Long, period As String
    On Error GoTo Failed
    If ctx Is Nothing Then Err.Raise 5, "clsMeasureEngine.EvaluateContext", "Evaluation context is required."
    ClearContext
    Set m_Ctx = ctx
    Set m_Active = CreateObject("Scripting.Dictionary")
    m_LastError = ""
    period = FinancialPeriodEnd(rptDate)
    m_Ctx.BeginPeriod period, CLng(ifAdjusted)
    EvaluateContext = ResolveValue(UCase$(Trim$(metricCode)), period)
    ClearContext
    Exit Function
Failed:
    n = Err.number
    m_LastError = Err.description
    ClearContext
    Select Case n
        Case vbObjectError + 2102: EvaluateContext = CVErr(xlErrNA)
        Case 11: EvaluateContext = CVErr(xlErrDiv0)
        Case Else: EvaluateContext = CVErr(xlErrValue)
    End Select
End Function

Private Sub ClearContext()
    Set m_Ctx = Nothing
    Set m_Active = Nothing
    m_Parent = 0
    m_Level = 0
End Sub

Public Sub ValidateContext(ByVal reportDate As String)
    If m_Ctx Is Nothing Then Err.Raise 5, "clsMeasureEngine", "No active evaluation context."
    m_Ctx.Validate reportDate
End Sub

Public Sub PushComparable(ByVal reportDate As String)
    If m_Ctx Is Nothing Then Err.Raise 5, "clsMeasureEngine", "No active evaluation context."
    m_Ctx.PushComparable reportDate
End Sub

Public Sub PopContext()
    If m_Ctx Is Nothing Then Err.Raise 5, "clsMeasureEngine", "No active evaluation context."
    m_Ctx.PopContext
End Sub

Private Function ComparablePart(ByVal expression As String, ByVal reportDate As String) As Double
    Dim pushed As Boolean, result As Double, n As Long, description As String, source As String
    On Error GoTo Failed
    PushComparable reportDate
    pushed = True
    result = CalculatePart(expression, reportDate, "PRIOR GROWTH BASE")
    PopContext
    pushed = False
    ComparablePart = result
    Exit Function
Failed:
    n = Err.number
    description = Err.description
    source = Err.source
    If pushed Then PopContext
    Err.Raise n, source, description
End Function

Public Function ResolveValue(ByVal code As String, ByVal reportDate As String) As Double
    Dim key As String, d As Variant, result As Double, denominator As Double
    Dim leafStatus As String, leafSource As String, resolved As Boolean
    Dim node As Object, parent As Long, level As Long, activated As Boolean
    Dim n As Long, description As String
    parent = m_Parent
    level = m_Level
    On Error GoTo Failed
    code = UCase$(Trim$(code))
    If m_Audit Then Set node = BeginAudit(code, reportDate, "", "")
    ValidateContext reportDate
    key = m_Ctx.ContextKey() & "|" & code
    If Not m_Audit Then
        If m_Ctx.TryGetMemo(key, result) Then
            ResolveValue = result
            Exit Function
        End If
    End If
    If m_Active.exists(key) Then Err.Raise 5, , "Circular evaluation: " & key
    If m_Active.count >= 100 Then Err.Raise 5, , "Evaluation depth exceeds 100."
    m_Active.Add key, True
    activated = True
    
    If m_Parameters.exists(code) Then
        SetAuditType node, "PARAMETER"
        result = m_Parameters(code)
        SetAuditRaw node, result, "OK"
    ElseIf m_Definitions.exists(code) Then
        d = m_Definitions(code)
        SetAuditType node, CStr(d(0))
        If Not d(3) Then Err.Raise vbObjectError + 2102, , "Disabled metric: " & code
        result = CalculatePart(CStr(d(1)), reportDate, "NUMERATOR / CURRENT")
        Select Case d(0)
            Case "RATIO"
                denominator = CalculatePart(CStr(d(2)), reportDate, "DENOMINATOR")
                If Abs(denominator) < 0.0000001 Then Err.Raise 11, , "Zero denominator: " & code
                result = result / denominator
            Case "GROWTH"
                denominator = ComparablePart(CStr(d(1)), Format$(DateAdd("yyyy", -1, CDate(reportDate)), "yyyy-mm-dd"))
                If Abs(denominator) < 0.0000001 Then Err.Raise 11, , "Zero growth base: " & code
                result = (result - denominator) / Abs(denominator)
            Case "DEFINED"
            Case Else
                Err.Raise 5, , "Unsupported formula type: " & code
        End Select
    Else
        SetAuditType node, "RAW", code
        resolved = m_Ctx.ResolveLeaf(code, reportDate, result, leafStatus, leafSource)
        If Not resolved Then
            Err.Raise 5, , "Unresolved symbol or unknown account: " & code
        End If
        SetAuditRaw node, result, leafStatus
    End If
    
    m_Active.Remove key
    activated = False
    m_Ctx.SetMemo key, result
    If m_Audit Then EndAudit node, result, "OK", ""
    m_Parent = parent
    m_Level = level
    ResolveValue = result
    Exit Function
Failed:
    n = Err.number
    description = Err.description
    If activated Then m_Active.Remove key
    If Not node Is Nothing Then EndAudit node, Empty, "ERROR", description
    m_Parent = parent
    m_Level = level
    Err.Raise n, "clsMeasureEngine.ResolveValue", description
End Function

Private Function CalculatePart(ByVal expression As String, ByVal reportDate As String, ByVal role As String) As Double
    Dim parser As clsMetricExpression, node As Object, result As Double
    Dim parent As Long, level As Long, n As Long, description As String
    parent = m_Parent
    level = m_Level
    On Error GoTo Failed
    If m_Audit Then Set node = BeginAudit(role, reportDate, "EXPRESSION", expression)
    ValidateContext reportDate
    Set parser = New clsMetricExpression
    result = parser.Calculate(expression, Me, reportDate)
    If m_Audit Then EndAudit node, result, "OK", ""
    m_Parent = parent
    m_Level = level
    CalculatePart = result
    Exit Function
Failed:
    n = Err.number
    description = Err.description
    If Not node Is Nothing Then EndAudit node, Empty, "ERROR", description
    m_Parent = parent
    m_Level = level
    Err.Raise n, "clsMeasureEngine.CalculatePart", description
End Function

Private Function BeginAudit(ByVal code As String, ByVal reportDate As String, ByVal kind As String, ByVal formula As String) As Object
    Dim node As Object, data As Variant, id As Long
    If m_AuditRows.count >= 20000 Then Err.Raise 5, , "Audit exceeds 20000 rows."
    id = m_AuditRows.count + 1
    data = Array(id, m_Parent, m_Level, kind, code, reportDate, "", formula, Empty, Empty, "", "", "OK", "")
    If Not m_Ctx Is Nothing Then
        data(10) = m_Ctx.CurrentPublicationDate
        data(11) = m_Ctx.currentVersionKey
    End If
    Set node = CreateObject("Scripting.Dictionary")
    node.Add "Data", data
    m_AuditRows.Add node
    m_Parent = id
    m_Level = m_Level + 1
    Set BeginAudit = node
End Function

Private Sub SetAuditType(ByVal node As Object, ByVal kind As String, Optional ByVal accountCode As String = "")
    Dim data As Variant
    If node Is Nothing Then Exit Sub
    data = node("Data")
    data(3) = kind
    data(6) = accountCode
    node("Data") = data
End Sub

Private Sub SetAuditRaw(ByVal node As Object, ByVal amount As Variant, ByVal status As String)
    Dim data As Variant
    If node Is Nothing Then Exit Sub
    data = node("Data")
    data(8) = amount
    data(12) = status
    If status = "DEFAULT_ZERO" Then data(13) = "Missing or NULL account treated as zero in the selected report."
    If status = "DEFAULT_TAX_RATE" Then data(13) = "Missing or NULL TAX_RATE in the context version; default rate 0.25 applied."
    If status = "AS_OF_FALLBACK" Then data(13) = "Exact date market observation absent; latest prior observation used."
    node("Data") = data
End Sub

Private Sub EndAudit(ByVal node As Object, ByVal Value As Variant, ByVal status As String, ByVal diagnostic As String)
    Dim data As Variant
    data = node("Data")
    data(9) = Value
    If status = "ERROR" Then data(12) = status
    If Len(diagnostic) > 0 Then data(13) = diagnostic
    node("Data") = data
End Sub

Public Function AuditMetric(ByVal stmt As clsFinancialStatement, ByVal rptDate As String, ByVal metricCode As String, Optional ByVal ifAdjusted As Integer = 2) As Variant
    Dim ctx As clsEvalContext, provider As clsStatementProvider
    Set provider = New clsStatementProvider
    provider.Initialize stmt
    Set ctx = New clsEvalContext
    ctx.RegisterProvider provider
    AuditMetric = AuditContextMetric(ctx, rptDate, metricCode, ifAdjusted)
End Function

Public Function AuditContextMetric(ByVal ctx As clsEvalContext, _
     ByVal rptDate As String, _
        ByVal metricCode As String, _
             Optional ByVal ifAdjusted As Integer = 2) As Variant
    Dim Value As Variant, result() As Variant, data As Variant, node As Object
    Dim i As Long, j As Long, n As Long, description As String
    On Error GoTo Failed
    m_LastError = ""
    Set m_AuditRows = New Collection
    m_Parent = 0
    m_Level = 0
    m_Audit = True
    Value = EvaluateContext(ctx, FinancialPeriodEnd(rptDate), metricCode, ifAdjusted)
    m_Audit = False
    If m_AuditRows.count = 0 Then Err.Raise 5, , "No audit rows: " & m_LastError
    ReDim result(1 To m_AuditRows.count, 1 To 14)
    For i = 1 To m_AuditRows.count
        Set node = m_AuditRows(i)
        data = node("Data")
        For j = 0 To 13
            result(i, j + 1) = data(j)
        Next j
    Next i
    If IsError(Value) Then result(1, 10) = Value
    AuditContextMetric = result
    Set m_AuditRows = Nothing
    m_Parent = 0
    m_Level = 0
    Exit Function
Failed:
    n = Err.number
    description = Err.description
    m_LastError = description
    m_Audit = False
    Set m_AuditRows = Nothing
    ClearContext
    Err.Raise n, "clsMeasureEngine.AuditContextMetric", description
End Function
