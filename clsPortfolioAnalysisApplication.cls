VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPortfolioAnalysisApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function RunPanelTable(ByVal panel As clsPortfolioPanel, ByVal context As clsAnalysisContext, ByVal tableName As String, ByVal values As Variant) As clsAnalysisResult
    Set RunPanelTable = BuildResult(panel, context, "PORTFOLIO_" & UCase$(tableName), UCase$(tableName), values)
End Function

Public Function RunIndustryAnalysis(ByVal panel As clsPortfolioPanel, ByVal context As clsAnalysisContext) As clsAnalysisResult
    Dim analyzer As clsPortfolioAnalyzer
    Set analyzer = New clsPortfolioAnalyzer: analyzer.Initialize panel
    Set RunIndustryAnalysis = BuildResult(panel, context, "PORTFOLIO_INDUSTRY_ANALYSIS", "INDUSTRY_ANALYSIS", analyzer.IndustryAnalysis)
End Function

Public Function RunAbsoluteAttribution(ByVal panel As clsPortfolioPanel, ByVal context As clsAnalysisContext) As clsAnalysisResult
    Dim analyzer As clsPortfolioAnalyzer
    Set analyzer = New clsPortfolioAnalyzer: analyzer.Initialize panel
    Set RunAbsoluteAttribution = BuildResult(panel, context, "PORTFOLIO_ABSOLUTE_ATTRIBUTION", "ABSOLUTE_ATTRIBUTION", analyzer.AbsoluteAttribution)
End Function

Public Function RunHoldingSummary(ByVal panel As clsPortfolioPanel, ByVal context As clsAnalysisContext, Optional ByVal amountsComparableConfirmed As Boolean = False) As clsAnalysisResult
    Dim analyzer As clsPortfolioAnalyzer, values(1 To 2, 1 To 2) As Variant
    Set analyzer = New clsPortfolioAnalyzer: analyzer.Initialize panel, amountsComparableConfirmed
    values(1, 1) = "METRIC": values(1, 2) = "VALUE"
    values(2, 1) = "TOTAL_HOLDING_VALUE": values(2, 2) = analyzer.TotalHoldingValue
    Set RunHoldingSummary = BuildResult(panel, context, "PORTFOLIO_HOLDING_SUMMARY", "HOLDING_SUMMARY", values, amountsComparableConfirmed)
End Function

Private Function BuildResult(ByVal panel As clsPortfolioPanel, ByVal context As clsAnalysisContext, ByVal taskType As String, ByVal tableName As String, ByVal values As Variant, Optional ByVal amountsComparableConfirmed As Boolean = False) As clsAnalysisResult
    Dim request As clsAnalysisRequest, result As clsAnalysisResult, data As clsDataTable, source As clsDataSet
    If panel Is Nothing Or context Is Nothing Then Err.Raise 5, "clsPortfolioAnalysisApplication", "Panel and context are required."
    If panel.endDate > context.valuationDate Or panel.endDate > context.KnowledgeAsOf Then Err.Raise 5, , "Panel includes observations beyond the analysis cutoff."
    Set request = New clsAnalysisRequest: request.Initialize taskType, context
    request.SetParameter "PORTFOLIO_CODES", panel.PortfolioCodes
    request.SetParameter "START_DATE", panel.startDate
    request.SetParameter "END_DATE", panel.endDate
    If taskType = "PORTFOLIO_HOLDING_SUMMARY" Then request.SetParameter "AMOUNTS_COMPARABLE_CONFIRMED", amountsComparableConfirmed
    Set data = New clsDataTable: data.Load values
    Set result = New clsAnalysisResult: result.Initialize request: result.AddTable tableName, data
    For Each source In panel.SourceDataSets: result.AddSource source: Next source
    If panel.SourceDataSets.count = 0 Then result.AddDiagnostic "MISSING_LINEAGE", panel.PortfolioCodes, "LOAD", "Panel has no source metadata."
    If data.rowCount = 0 Then result.AddDiagnostic "NO_DATA", panel.PortfolioCodes, "ANALYSIS", "No result rows were produced."
    result.Freeze: Set BuildResult = result
End Function

Public Function RunReturnAnalysis(ByVal panel As clsPortfolioPanel, ByVal context As clsAnalysisContext, Optional ByVal metricKeys As Variant) As clsAnalysisResult
    Dim request As clsAnalysisRequest, analyzer As clsPortfolioAnalyzer, values As Variant
    Dim table As clsDataTable, result As clsAnalysisResult, dataset As clsDataSet, r As Long
    If panel Is Nothing Or context Is Nothing Then Err.Raise 5, "clsPortfolioAnalysisApplication", "Panel and context are required."
    If panel.endDate > context.valuationDate Or panel.endDate > context.KnowledgeAsOf Then Err.Raise 5, , "Panel includes observations beyond the analysis cutoff."
    Set request = New clsAnalysisRequest: request.Initialize "PORTFOLIO_RETURN_ANALYSIS", context
    request.SetParameter "PORTFOLIO_CODES", panel.PortfolioCodes
    request.SetParameter "START_DATE", panel.startDate
    request.SetParameter "END_DATE", panel.endDate
    Set analyzer = New clsPortfolioAnalyzer: analyzer.Initialize panel
    If IsMissing(metricKeys) Then
        values = analyzer.ReturnAnalysis()
    Else
        values = analyzer.ReturnAnalysis(metricKeys)
    End If
    Set table = New clsDataTable: table.Load values
    Set result = New clsAnalysisResult: result.Initialize request
    result.AddTable "RETURNS_ANALYSIS", table
    For Each dataset In panel.SourceDataSets: result.AddSource dataset: Next dataset
    If panel.SourceDataSets.count = 0 Then result.AddDiagnostic "MISSING_LINEAGE", panel.PortfolioCodes, "LOAD", "Panel has no source metadata."
    If table.rowCount = 0 Then result.AddDiagnostic "INSUFFICIENT_DATA", panel.PortfolioCodes, "ANALYSIS", "No return observations."
    For r = 1 To table.rowCount
        If CStr(table.ValueAt(r, "STATUS")) <> "OK" Then result.AddDiagnostic CStr(table.ValueAt(r, "STATUS")), CStr(table.ValueAt(r, "PORT_CODE")), "ANALYSIS", "Review return result row " & CStr(r)
    Next r
    result.Freeze: Set RunReturnAnalysis = result
End Function
