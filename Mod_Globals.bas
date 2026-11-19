Attribute VB_Name = "Mod_Globals"
Option Explicit

' 统一交付模式常量
Public Const DELIV_DATA_ONLY As Long = 0
Public Const DELIV_UDF_SCALAR As Long = 1
Public Const DELIV_CURRENT_WORKBOOK As Long = 2
Public Const DELIV_NEW_WORKBOOK As Long = 3

Private g_ConnManager As clsConnectionManager
Private g_MemoryCache As clsMemoryCache
Private g_Measures As clsMeasureEngine
Private g_MetricError As String

' ==============================================================================
' 统一业务领域门面 (Domain Facades) —— 支持 Equity.xxx, Financial.xxx, Portfolio.xxx
' ==============================================================================

Public Property Get equity() As clsEquityFacade
    Static s_Equity As clsEquityFacade
    If s_Equity Is Nothing Then Set s_Equity = New clsEquityFacade
    Set equity = s_Equity
End Property

Public Property Get Financial() As clsFinancialFacade
    Static s_Financial As clsFinancialFacade
    If s_Financial Is Nothing Then Set s_Financial = New clsFinancialFacade
    Set Financial = s_Financial
End Property

Public Property Get Portfolio() As clsPortfolioFacade
    Static s_Portfolio As clsPortfolioFacade
    If s_Portfolio Is Nothing Then Set s_Portfolio = New clsPortfolioFacade
    Set Portfolio = s_Portfolio
End Property


' Fiscal year strings map to annual December reports; full dates remain supported.
Public Function FinancialPeriodEnd(ByVal fiscalPeriod As String) As String
    fiscalPeriod = Trim$(fiscalPeriod)
    If Len(fiscalPeriod) = 4 And fiscalPeriod Like "####" Then
        If CLng(fiscalPeriod) < 1900 Or CLng(fiscalPeriod) > 9999 Then Err.Raise 5, , "Invalid fiscal year."
        FinancialPeriodEnd = fiscalPeriod & "-12-31"
    Else
        FinancialPeriodEnd = Format$(CDate(fiscalPeriod), "yyyy-mm-dd")
    End If
End Function

Public Function MeasureMgr() As clsMeasureEngine
    Dim candidate As clsMeasureEngine
    If g_Measures Is Nothing Then
        Set candidate = New clsMeasureEngine
        candidate.LoadFile ThisWorkbook.Path & "\Config\MEASURES_DEF.xlsx", ThisWorkbook.Path & "\Config\measure_parameters.ini"
        Set g_Measures = candidate
    End If
    Set MeasureMgr = g_Measures
End Function

Public Function FIN_METRIC_LAST_ERROR() As String
    FIN_METRIC_LAST_ERROR = g_MetricError
End Function

Public Function ConnMgr() As clsConnectionManager
    If g_ConnManager Is Nothing Then Set g_ConnManager = New clsConnectionManager
    Set ConnMgr = g_ConnManager
End Function

Public Function CacheMgr() As clsMemoryCache
    If g_MemoryCache Is Nothing Then Set g_MemoryCache = New clsMemoryCache
    Set CacheMgr = g_MemoryCache
End Function

Public Sub ResetGlobals()
    Set g_Measures = Nothing
    g_MetricError = ""
    If Not g_ConnManager Is Nothing Then
        g_ConnManager.CloseAll
        Set g_ConnManager = Nothing
    End If
    If Not g_MemoryCache Is Nothing Then
        g_MemoryCache.Clear
        Set g_MemoryCache = Nothing
    End If
End Sub

Public Function GET_FIN_METRIC(ByVal companyCode As String, _
        ByVal rptDate As String, _
        ByVal metricCode As String, _
        Optional ByVal ifMerged As Long = 1, _
        Optional ByVal ifAdjusted As Integer = 2, _
        Optional ByVal asOfDate As String = "", _
        Optional ByVal publishedDate As String = "") As Variant

    Dim reportDay As Date, priorDay As Date
    Dim srv As New clsFinancialService, stmt As clsFinancialStatement
    On Error GoTo ErrH
    g_MetricError = ""
    Application.Volatile False
    If ifAdjusted <> 1 And ifAdjusted <> 2 Then Err.Raise 5, , "IfAdjusted must be 1 or 2."
    reportDay = CDate(FinancialPeriodEnd(rptDate))
    ' Load 3 years of prior history to support multi-stage nested AVG / GROWTH formulas.
    priorDay = DateAdd("yyyy", -3, reportDay)
    rptDate = Format$(reportDay, "yyyy-mm-dd")
    Set stmt = srv.GetStatement(companyCode, Format$(priorDay, "yyyy-mm-dd"), rptDate, ifMerged, asOfDate, publishedDate, ifAdjusted)
    GET_FIN_METRIC = EvaluateMetric(stmt, rptDate, metricCode, ifAdjusted)
    Exit Function
ErrH:
    g_MetricError = Err.description
    GET_FIN_METRIC = CVErr(xlErrValue)
End Function

Public Function EvaluateMetric(ByVal stmt As clsFinancialStatement, ByVal rptDate As String, ByVal metricCode As String, Optional ByVal ifAdjusted As Integer = 2) As Variant
    Dim engine As clsMeasureEngine
    On Error GoTo Failed
    g_MetricError = ""
    rptDate = FinancialPeriodEnd(rptDate)
    Set engine = MeasureMgr()
    EvaluateMetric = engine.Evaluate(stmt, rptDate, metricCode, ifAdjusted)
    g_MetricError = engine.LastError
    Exit Function
Failed:
    g_MetricError = Err.description
    EvaluateMetric = CVErr(xlErrValue)
End Function

Public Function EvaluateContextMetric(ByVal ctx As clsEvalContext, ByVal rptDate As String, ByVal metricCode As String, Optional ByVal ifAdjusted As Integer = 2) As Variant
    Dim engine As clsMeasureEngine
    On Error GoTo Failed
    g_MetricError = ""
    rptDate = FinancialPeriodEnd(rptDate)
    Set engine = MeasureMgr()
    EvaluateContextMetric = engine.EvaluateContext(ctx, rptDate, metricCode, ifAdjusted)
    g_MetricError = engine.LastError
    Exit Function
Failed:
    g_MetricError = Err.description
    EvaluateContextMetric = CVErr(xlErrValue)
End Function
