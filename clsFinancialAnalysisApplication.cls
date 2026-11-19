VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFinancialAnalysisApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function RunMetric(ByVal statement As clsFinancialStatement, ByVal context As clsAnalysisContext, ByVal reportPeriod As String, ByVal metricCode As String, Optional ByVal ifAdjusted As Long = 2) As clsAnalysisResult
    Dim engine As clsMeasureEngine
    Set engine = MeasureMgr()
    Set RunMetric = RunMetricCore(statement, context, reportPeriod, metricCode, ifAdjusted, engine)
End Function

Public Function RunMetricWithEngine(ByVal statement As clsFinancialStatement, ByVal context As clsAnalysisContext, ByVal reportPeriod As String, ByVal metricCode As String, ByVal ifAdjusted As Long, ByVal engine As clsMeasureEngine) As clsAnalysisResult
    Set RunMetricWithEngine = RunMetricCore(statement, context, reportPeriod, metricCode, ifAdjusted, engine)
End Function

Public Function RunAuditMetric(ByVal statement As clsFinancialStatement, ByVal context As clsAnalysisContext, ByVal reportPeriod As String, ByVal metricCode As String, Optional ByVal ifAdjusted As Long = 2) As clsAnalysisResult
    Dim engine As clsMeasureEngine
    Set engine = MeasureMgr()
    Set RunAuditMetric = RunAuditMetricWithEngine(statement, context, reportPeriod, metricCode, ifAdjusted, engine)
End Function

Public Function RunAuditMetricWithEngine(ByVal statement As clsFinancialStatement, ByVal context As clsAnalysisContext, ByVal reportPeriod As String, ByVal metricCode As String, ByVal ifAdjusted As Long, ByVal engine As clsMeasureEngine) As clsAnalysisResult
    Dim request As clsAnalysisRequest, auditRows As Variant, tableData() As Variant
    Dim headers As Variant, r As Long, c As Long, rowCount As Long
    Dim table As clsDataTable, result As clsAnalysisResult, nodeStatus As String, diagText As String
    Dim periodEnd As String
    periodEnd = ValidateStatementContext(statement, context, reportPeriod, ifAdjusted, engine)
    Set request = New clsAnalysisRequest
    request.Initialize "FINANCIAL_METRIC_AUDIT", context
    request.SetParameter "COMPANY_CODE", statement.companyCode
    request.SetParameter "REPORT_PERIOD", periodEnd
    request.SetParameter "METRIC_CODE", UCase$(Trim$(metricCode))
    request.SetParameter "IF_ADJUSTED", ifAdjusted
    request.SetParameter "IF_MERGED", statement.ifMerged

    auditRows = engine.AuditMetric(statement, periodEnd, metricCode, CInt(ifAdjusted))
    rowCount = UBound(auditRows, 1)
    headers = Array("NODE_ID", "PARENT_ID", "LEVEL", "NODE_TYPE", "IDENTIFIER", "REPORT_PERIOD", "ACCOUNT_CODE", "EXPRESSION", "RAW_VALUE", "CALCULATED_VALUE", "SELECTED_PUBLICATION", "SELECTED_VERSION_KEY", "STATUS", "DIAGNOSTIC")
    ReDim tableData(1 To rowCount + 1, 1 To 14)
    For c = 1 To 14
        tableData(1, c) = headers(c - 1)
    Next c
    For r = 1 To rowCount
        For c = 1 To 14
            tableData(r + 1, c) = auditRows(r, c)
        Next c
    Next r

    Set table = New clsDataTable
    table.Load tableData
    Set result = New clsAnalysisResult
    result.Initialize request
    result.AddTable "METRIC_AUDIT", table
    AttachStatementSources statement, result

    For r = 1 To rowCount
        nodeStatus = CStr(auditRows(r, 13))
        diagText = CStr(auditRows(r, 14))
        If Len(diagText) = 0 Then diagText = "Node " & CStr(auditRows(r, 1)) & " (" & CStr(auditRows(r, 5)) & ")"
        Select Case nodeStatus
            Case "ERROR"
                result.AddDiagnostic "AUDIT_NODE_ERROR", CStr(auditRows(r, 5)), "ANALYSIS", diagText, "ERROR"
            Case "DEFAULT_ZERO", "DEFAULT_TAX_RATE", "AS_OF_FALLBACK"
                result.AddDiagnostic nodeStatus, CStr(auditRows(r, 5)), "ANALYSIS", diagText, "WARNING"
        End Select
    Next r

    result.Freeze
    Set RunAuditMetricWithEngine = result
End Function

Public Function RunSelectedSeries(ByVal statement As clsFinancialStatement, ByVal context As clsAnalysisContext, Optional ByVal ifAdjusted As Long = 2) As clsAnalysisResult
    Dim request As clsAnalysisRequest, facts As Variant, tableData() As Variant
    Dim headers As Variant, r As Long, c As Long, rowCount As Long
    Dim table As clsDataTable, result As clsAnalysisResult
    If statement Is Nothing Or context Is Nothing Then Err.Raise 5, "clsFinancialAnalysisApplication", "Statement and context are required."
    If ifAdjusted <> 1 And ifAdjusted <> 2 Then Err.Raise 5, , "IfAdjusted must be 1 or 2."
    If statement.asOfDate <> Format$(context.KnowledgeAsOf, "yyyy-mm-dd") Then Err.Raise 5, , "Statement as-of date must match context."

    Set request = New clsAnalysisRequest
    request.Initialize "FINANCIAL_SELECTED_SERIES", context
    request.SetParameter "COMPANY_CODE", statement.companyCode
    request.SetParameter "IF_ADJUSTED", ifAdjusted
    request.SetParameter "IF_MERGED", statement.ifMerged

    facts = statement.ExportSelectedLongTable(ifAdjusted)
    headers = Array("COMPANY_CODE", "END_DATE", "IF_MERGED", "IF_ADJUSTED", "INFO_PUBL_DATE", "ACCOUNT_CODE", "AMOUNT")
    If IsEmpty(facts) Then
        rowCount = 0
    Else
        rowCount = UBound(facts, 1)
    End If
    ReDim tableData(1 To rowCount + 1, 1 To 7)
    For c = 1 To 7
        tableData(1, c) = headers(c - 1)
    Next c
    For r = 1 To rowCount
        For c = 1 To 7
            tableData(r + 1, c) = facts(r, c)
        Next c
    Next r

    Set table = New clsDataTable
    table.Load tableData
    Set result = New clsAnalysisResult
    result.Initialize request
    result.AddTable "SELECTED_SERIES", table
    AttachStatementSources statement, result
    If rowCount = 0 Then
        result.AddDiagnostic "NO_DATA", statement.companyCode, "ANALYSIS", "No selected financial report rows found.", "WARNING"
    End If
    result.Freeze
    Set RunSelectedSeries = result
End Function

Public Function RunTemplateSnapshot(ByVal statement As clsFinancialStatement, _
        ByVal context As clsAnalysisContext, _
        ByVal sourceAlias As String, _
        ByVal startYear As Long, _
        ByVal endYear As Long, _
        ByVal codes As Variant, _
        ByVal definitions As Object, _
        ByVal rawCodes As Object, _
        Optional ByVal ifAdjusted As Long = 2, _
        Optional ByVal engine As clsMeasureEngine = Nothing) As clsAnalysisResult
    Dim activeEngine As clsMeasureEngine, request As clsAnalysisRequest, result As clsAnalysisResult
    Dim valuesTable() As Variant, notesTable() As Variant
    Dim valuesData As clsDataTable, notesData As clsDataTable
    Dim count As Long, r As Long, col As Long, fiscalYear As Long
    Dim code As String, period As String, version As String, diagnostic As String
    Dim evalVal As Variant, amount As Variant, asOfText As String
    If engine Is Nothing Then
        Set activeEngine = MeasureMgr()
    Else
        Set activeEngine = engine
    End If
    If startYear < 1901 Or endYear > 9999 Or startYear > endYear Then Err.Raise 5, , "Years must satisfy 1901 <= start <= end <= 9999."
    Call ValidateStatementContext(statement, context, CStr(endYear) & "-12-31", ifAdjusted, activeEngine)

    Set request = New clsAnalysisRequest
    request.Initialize "FINANCIAL_TEMPLATE_SNAPSHOT", context
    request.SetParameter "COMPANY_CODE", statement.companyCode
    request.SetParameter "START_YEAR", startYear
    request.SetParameter "END_YEAR", endYear
    request.SetParameter "IF_MERGED", statement.ifMerged
    request.SetParameter "IF_ADJUSTED", ifAdjusted

    Set result = New clsAnalysisResult
    result.Initialize request
    AttachStatementSources statement, result

    count = endYear - startYear + 1
    asOfText = statement.asOfDate
    ReDim valuesTable(1 To 46, 1 To count + 1)
    ReDim notesTable(1 To 46, 1 To count + 1)

    valuesTable(1, 1) = "TEMPLATE_LABEL"
    notesTable(1, 1) = "TEMPLATE_LABEL"
    For col = 2 To count + 1
        valuesTable(1, col) = "Y" & CStr(startYear + col - 2)
        notesTable(1, col) = "Y" & CStr(startYear + col - 2)
    Next col

    For r = 1 To 45
        valuesTable(r + 1, 1) = codes(r, 1)
        notesTable(r + 1, 1) = ""
    Next r
    valuesTable(2, 1) = "Company " & statement.companyCode & " | annual report snapshot"
    valuesTable(3, 1) = "Source: " & sourceAlias
    valuesTable(4, 1) = "Cutoff: " & asOfText
    valuesTable(5, 1) = "Merged=" & CStr(statement.ifMerged) & "; adjusted=" & CStr(ifAdjusted)
    valuesTable(6, 1) = "Units: source values; missing raw: DEFAULT_ZERO"

    For fiscalYear = startYear To endYear
        col = fiscalYear - startYear + 2
        period = CStr(fiscalYear) & "-12-31"
        version = statement.SelectedVersionKey(period, ifAdjusted)
        valuesTable(2, col) = "Company " & statement.companyCode
        valuesTable(3, col) = "Source: " & sourceAlias
        valuesTable(4, col) = "Published: " & statement.SelectedPublicationDate(period, ifAdjusted)
        valuesTable(5, col) = "Version: " & version
        valuesTable(6, col) = "Cutoff: " & asOfText
        valuesTable(7, col) = fiscalYear
        If version = "" Then
            diagnostic = "Missing selected report: company=" & statement.companyCode & "; period=" & period & "; merged=" & CStr(statement.ifMerged) & "; adjusted=" & CStr(ifAdjusted) & "; cutoff=" & asOfText
            valuesTable(4, col) = "MISSING SELECTED REPORT"
            result.AddDiagnostic "MISSING_REPORT", period, "ANALYSIS", diagnostic, "WARNING"
        End If
        For r = 8 To 45
            code = UCase$(Trim$(CStr(codes(r, 1))))
            If code <> "" Then
                If version = "" Then
                    valuesTable(r + 1, col) = CVErr(xlErrNA)
                    notesTable(r + 1, col) = diagnostic
                Else
                    evalVal = activeEngine.Evaluate(statement, period, code, CInt(ifAdjusted))
                    valuesTable(r + 1, col) = evalVal
                    If IsError(evalVal) Then
                        notesTable(r + 1, col) = activeEngine.LastError
                        If notesTable(r + 1, col) = "" Then notesTable(r + 1, col) = "Evaluation failed: " & code & " / " & version
                        result.AddDiagnostic "TEMPLATE_METRIC_ERROR", code & "@" & period, "ANALYSIS", notesTable(r + 1, col), "ERROR"
                    ElseIf code = "TAX_RATE" Then
                        If Not statement.TryGetVersionRawValue(version, "TAX_RATE", amount) Then
                            notesTable(r + 1, col) = "DEFAULT_TAX_RATE: missing or NULL TAX_RATE; rate 0.25 applied / " & version
                            result.AddDiagnostic "DEFAULT_TAX_RATE", code & "@" & period, "ANALYSIS", notesTable(r + 1, col), "WARNING"
                        End If
                    ElseIf Not definitions.exists(code) Then
                        If Not statement.TryGetVersionRawValue(version, CStr(rawCodes(code)), amount) Then
                            notesTable(r + 1, col) = "DEFAULT_ZERO: missing or NULL raw account " & code & " / " & version
                            result.AddDiagnostic "DEFAULT_ZERO", code & "@" & period, "ANALYSIS", notesTable(r + 1, col), "WARNING"
                        End If
                    End If
                End If
            End If
        Next r
    Next fiscalYear

    Set valuesData = New clsDataTable: valuesData.Load valuesTable
    Set notesData = New clsDataTable: notesData.Load notesTable
    result.AddTable "TEMPLATE_VALUES", valuesData
    result.AddTable "TEMPLATE_NOTES", notesData
    result.Freeze
    Set RunTemplateSnapshot = result
End Function

Private Function RunMetricCore(ByVal statement As clsFinancialStatement, ByVal context As clsAnalysisContext, ByVal reportPeriod As String, ByVal metricCode As String, ByVal ifAdjusted As Long, ByVal engine As clsMeasureEngine) As clsAnalysisResult
    Dim request As clsAnalysisRequest, Value As Variant
    Dim table As clsDataTable, values() As Variant, result As clsAnalysisResult, status As String
    Dim periodEnd As String
    periodEnd = ValidateStatementContext(statement, context, reportPeriod, ifAdjusted, engine)
    Set request = New clsAnalysisRequest: request.Initialize "FINANCIAL_METRIC", context
    request.SetParameter "COMPANY_CODE", statement.companyCode
    request.SetParameter "REPORT_PERIOD", periodEnd
    request.SetParameter "METRIC_CODE", UCase$(Trim$(metricCode))
    request.SetParameter "IF_ADJUSTED", ifAdjusted
    request.SetParameter "IF_MERGED", statement.ifMerged
    Value = engine.Evaluate(statement, periodEnd, metricCode, CInt(ifAdjusted))
    If IsError(Value) Then status = "CALCULATION_ERROR" Else status = "OK"
    ReDim values(1 To 2, 1 To 4)
    values(1, 1) = "REPORT_PERIOD": values(1, 2) = "METRIC_CODE": values(1, 3) = "VALUE": values(1, 4) = "STATUS"
    values(2, 1) = periodEnd: values(2, 2) = UCase$(Trim$(metricCode))
    If IsError(Value) Then values(2, 3) = Null Else values(2, 3) = Value
    values(2, 4) = status
    Set table = New clsDataTable: table.Load values
    Set result = New clsAnalysisResult: result.Initialize request
    result.AddTable "FINANCIAL_METRIC", table
    AttachStatementSources statement, result
    If IsError(Value) Then result.AddDiagnostic "FINANCIAL_METRIC_ERROR", metricCode, "ANALYSIS", engine.LastError, "ERROR"
    result.Freeze: Set RunMetricCore = result
End Function

Private Function ValidateStatementContext(ByVal statement As clsFinancialStatement, _
        ByVal context As clsAnalysisContext, _
        ByVal reportPeriod As String, _
        ByVal ifAdjusted As Long, _
        ByVal engine As clsMeasureEngine) As String
    Dim periodEnd As String
    If statement Is Nothing Or context Is Nothing Or engine Is Nothing Then
        Err.Raise 5, "clsFinancialAnalysisApplication", "Statement, context and metric engine are required."
    End If
    If ifAdjusted <> 1 And ifAdjusted <> 2 Then Err.Raise 5, , "IfAdjusted must be 1 or 2."
    If statement.asOfDate <> Format$(context.KnowledgeAsOf, "yyyy-mm-dd") Then
        Err.Raise 5, , "Statement as-of date must match context."
    End If
    periodEnd = FinancialPeriodEnd(reportPeriod)
    If CDate(periodEnd) > dateValue(context.valuationDate) Then
        Err.Raise 5, "clsFinancialAnalysisApplication", "Report period exceeds context valuation date."
    End If
    ValidateStatementContext = periodEnd
End Function

Private Sub AttachStatementSources(ByVal statement As clsFinancialStatement, ByVal result As clsAnalysisResult)
    Dim sources As Collection, dataset As clsDataSet
    Set sources = statement.SourceDataSets
    For Each dataset In sources
        result.AddSource dataset
    Next dataset
    If sources.count = 0 Then
        result.AddDiagnostic "MISSING_LINEAGE", statement.companyCode, "LOAD", "Statement has no source metadata."
    End If
End Sub

' Single-company multi-period common-size analysis
Public Function RunCommonSizeSingleCompany(ByVal statement As clsFinancialStatement, _
        ByVal context As clsAnalysisContext, _
        ByVal periods As Variant, _
        Optional ByVal stmtType As String = "IS", _
        Optional ByVal displayMode As String = "BOTH", _
        Optional ByVal ifAdjusted As Long = 2, _
        Optional ByVal engine As clsMeasureEngine = Nothing) As clsAnalysisResult
    Dim activeEngine As clsMeasureEngine, analyzer As clsStatementCommonSizeAnalyzer
    Dim layoutTbl As clsDataTable, longTbl As clsDataTable, wideTbl As clsDataTable, notesTbl As clsDataTable
    Dim request As clsAnalysisRequest, result As clsAnalysisResult, r As Long, st As String

    If statement Is Nothing Or context Is Nothing Then
        Err.Raise 5, "clsFinancialAnalysisApplication.RunCommonSizeSingleCompany", "Statement and context are required."
    End If
    If engine Is Nothing Then Set activeEngine = MeasureMgr() Else Set activeEngine = engine

    Set analyzer = New clsStatementCommonSizeAnalyzer
    Set layoutTbl = analyzer.LoadLayout(stmtType)
    Set longTbl = analyzer.BuildSingleCompanyLong(statement, layoutTbl, periods, ifAdjusted, activeEngine)
    Set wideTbl = analyzer.PivotSingleCompanyMultiPeriod(longTbl, layoutTbl, periods, displayMode, notesTbl)

    Set request = New clsAnalysisRequest
    request.Initialize "FINANCIAL_COMMON_SIZE_MULTI_PERIOD", context
    request.SetParameter "COMPANY_CODE", statement.companyCode
    request.SetParameter "STMT_TYPE", UCase$(Trim$(stmtType))
    request.SetParameter "DISPLAY_MODE", UCase$(Trim$(displayMode))
    request.SetParameter "IF_MERGED", statement.ifMerged
    request.SetParameter "IF_ADJUSTED", ifAdjusted

    Set result = New clsAnalysisResult
    result.Initialize request
    result.AddTable "MULTI_PERIOD_COMPARISON", wideTbl
    result.AddTable "MULTI_PERIOD_NOTES", notesTbl
    result.AddTable "DETAIL_LONG", longTbl
    AttachStatementSources statement, result

    For r = 1 To longTbl.rowCount
        st = CStr(longTbl.ValueAt(r, "STATUS"))
        If st = "ERROR" Then
            result.AddDiagnostic "COMMON_SIZE_ERROR", CStr(longTbl.ValueAt(r, "ACCT_CODE")) & "@" & CStr(longTbl.ValueAt(r, "REPORT_PERIOD")), "ANALYSIS", CStr(longTbl.ValueAt(r, "DIAGNOSTIC")), "ERROR"
        ElseIf st = "MISSING_REPORT" Or st = "DEFAULT_TAX_RATE" Then
            result.AddDiagnostic st, CStr(longTbl.ValueAt(r, "ACCT_CODE")) & "@" & CStr(longTbl.ValueAt(r, "REPORT_PERIOD")), "ANALYSIS", CStr(longTbl.ValueAt(r, "DIAGNOSTIC")), "WARNING"
        End If
    Next r

    result.Freeze
    Set RunCommonSizeSingleCompany = result
End Function

' Multi-company same-period peer common-size comparison
Public Function RunCommonSizePeerComparison(ByVal statements As Collection, _
        ByVal context As clsAnalysisContext, _
        ByVal reportPeriod As String, _
        Optional ByVal stmtType As String = "IS", _
        Optional ByVal sliceField As String = "COMMON_SIZE", _
        Optional ByVal ifAdjusted As Long = 2, _
        Optional ByVal engine As clsMeasureEngine = Nothing, _
        Optional ByVal companyLabels As Object = Nothing) As clsAnalysisResult
    Dim activeEngine As clsMeasureEngine, analyzer As clsStatementCommonSizeAnalyzer
    Dim layoutTbl As clsDataTable, peerTbl As clsDataTable, longTbl As clsDataTable, notesTbl As clsDataTable
    Dim request As clsAnalysisRequest, result As clsAnalysisResult, stmt As clsFinancialStatement
    Dim periodEnd As String, r As Long, st As String

    If statements Is Nothing Or statements.count = 0 Or context Is Nothing Then
        Err.Raise 5, "clsFinancialAnalysisApplication.RunCommonSizePeerComparison", "Statements collection and context are required."
    End If
    If engine Is Nothing Then Set activeEngine = MeasureMgr() Else Set activeEngine = engine

    periodEnd = FinancialPeriodEnd(reportPeriod)
    Set analyzer = New clsStatementCommonSizeAnalyzer
    Set layoutTbl = analyzer.LoadLayout(stmtType)
    Set peerTbl = analyzer.BuildPeerSamePeriodComparison(statements, layoutTbl, periodEnd, sliceField, ifAdjusted, activeEngine, longTbl, notesTbl, companyLabels)

    Set request = New clsAnalysisRequest
    request.Initialize "FINANCIAL_COMMON_SIZE_PEER", context
    request.SetParameter "REPORT_PERIOD", periodEnd
    request.SetParameter "STMT_TYPE", UCase$(Trim$(stmtType))
    request.SetParameter "SLICE_FIELD", UCase$(Trim$(sliceField))
    request.SetParameter "PEER_COUNT", statements.count
    request.SetParameter "IF_ADJUSTED", ifAdjusted

    Set result = New clsAnalysisResult
    result.Initialize request
    result.AddTable "PEER_COMPARISON", peerTbl
    result.AddTable "PEER_NOTES", notesTbl
    result.AddTable "DETAIL_LONG", longTbl
    For Each stmt In statements
        AttachStatementSources stmt, result
    Next stmt

    For r = 1 To longTbl.rowCount
        st = CStr(longTbl.ValueAt(r, "STATUS"))
        If st = "ERROR" Then
            result.AddDiagnostic "PEER_EVAL_ERROR", CStr(longTbl.ValueAt(r, "COMPANY_CODE")) & "|" & CStr(longTbl.ValueAt(r, "ACCT_CODE")), "ANALYSIS", CStr(longTbl.ValueAt(r, "DIAGNOSTIC")), "ERROR"
        ElseIf st = "MISSING_REPORT" Then
            result.AddDiagnostic st, CStr(longTbl.ValueAt(r, "COMPANY_CODE")) & "@" & periodEnd, "ANALYSIS", CStr(longTbl.ValueAt(r, "DIAGNOSTIC")), "WARNING"
        End If
    Next r

    result.Freeze
    Set RunCommonSizePeerComparison = result
End Function
