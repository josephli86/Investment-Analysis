VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEquityFacade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' clsEquityFacade: 股票与行情分析统一领域门面 (Domain Facade)
' 统一收口与暴露所有股票分析功能，支持代码自动补全 (Equity.xxx)
' 涵盖行情、技术指标、区间风险度量、相对估值与报表一键交付
' ==============================================================================

Private m_EquityService As clsEquityService
Private m_ValuationAnalyzer As clsEquityValuationAnalyzer
Private m_Renderer As clsResultRenderer

Private Sub Class_Initialize()
    Set m_EquityService = New clsEquityService
    Set m_ValuationAnalyzer = New clsEquityValuationAnalyzer
    Set m_Renderer = New clsResultRenderer
End Sub

Public Function Quotes(ByVal secCode As String, _
                       Optional ByVal beginDate As Variant, _
                       Optional ByVal endDate As Variant, _
                       Optional ByVal includeWarmup As Boolean = False) As clsDataTable
    Dim panel As clsEquityPanel, ts As clsTimeSeriesFrame
    Dim dStart As Date, dEnd As Date
    ResolveDates beginDate, endDate, dStart, dEnd

    Set panel = m_EquityService.GetQuotePanel(Array(secCode), dStart, dEnd)
    Set ts = panel.QuoteFrameForStock(secCode, includeWarmup)
    If ts Is Nothing Then
        Set Quotes = New clsDataTable
    Else
        Set Quotes = New clsDataTable
        Quotes.Load ts.ToTable()
    End If
End Function

Public Function SecInfo(ByVal secCode As String) As clsDataTable
    Dim panel As clsEquityPanel
    Dim dEnd As Date: dEnd = Date
    Dim dStart As Date: dStart = DateAdd("d", -10, dEnd)
    Set panel = m_EquityService.GetQuotePanel(Array(secCode), dStart, dEnd, 0)
    Set SecInfo = panel.MasterTable
End Function

Public Function LatestPrice(ByVal secCode As String) As Variant
    Dim tbl As clsDataTable, rCount As Long
    Set tbl = Quotes(secCode, DateAdd("d", -30, Date), Date, False)
    rCount = tbl.rowCount
    If rCount > 0 Then
        LatestPrice = tbl.ValueAt(rCount, "CLOSE_PRICE")
    Else
        LatestPrice = Null
    End If
End Function

' ------------------------------------------------------------------------------
' 2. 收益与风险度量 (Returns & Risk Measures)
' ------------------------------------------------------------------------------

Public Function ReturnRate(ByVal secCode As String, ByVal beginDate As Variant, ByVal endDate As Variant) As Variant
    Dim ts As clsTimeSeriesFrame
    Set ts = GetStockFrame(secCode, beginDate, endDate, False)
    If ts Is Nothing Then ReturnRate = Null: Exit Function
    ReturnRate = TS_CumulativeReturn(ts, "DAILY_RTN")
End Function

Public Function AnnualizedReturn(ByVal secCode As String, ByVal beginDate As Variant, ByVal endDate As Variant) As Variant
    Dim ts As clsTimeSeriesFrame
    Set ts = GetStockFrame(secCode, beginDate, endDate, False)
    If ts Is Nothing Then AnnualizedReturn = Null: Exit Function
    AnnualizedReturn = TS_AnnualizedReturn(ts, "DAILY_RTN")
End Function

Public Function Volatility(ByVal secCode As String, ByVal beginDate As Variant, ByVal endDate As Variant) As Variant
    Dim ts As clsTimeSeriesFrame
    Set ts = GetStockFrame(secCode, beginDate, endDate, False)
    If ts Is Nothing Then Volatility = Null: Exit Function
    Volatility = TS_Volatility(ts, "DAILY_RTN")
End Function

Public Function SharpeRatio(ByVal secCode As String, ByVal beginDate As Variant, ByVal endDate As Variant, _
                            Optional ByVal riskFreeRate As Double = 0.02) As Variant
    Dim ts As clsTimeSeriesFrame
    Set ts = GetStockFrame(secCode, beginDate, endDate, False)
    If ts Is Nothing Then SharpeRatio = Null: Exit Function
    SharpeRatio = TS_SharpeRatio(ts, "DAILY_RTN", "", 252#) - (riskFreeRate * Sqr(252#) / 1#)
End Function

Public Function maxDrawdown(ByVal secCode As String, ByVal beginDate As Variant, ByVal endDate As Variant) As Variant
    Dim ts As clsTimeSeriesFrame
    Set ts = GetStockFrame(secCode, beginDate, endDate, False)
    If ts Is Nothing Then maxDrawdown = Null: Exit Function
    maxDrawdown = TS_MaxDrawdown(ts, "DAILY_RTN")
End Function

Public Function SMA(ByVal secCode As String, ByVal n As Long, _
                    Optional ByVal beginDate As Variant, Optional ByVal endDate As Variant) As Variant
    Dim ts As clsTimeSeriesFrame, res As Variant
    Set ts = GetStockFrame(secCode, beginDate, endDate, True)
    If ts Is Nothing Then SMA = Null: Exit Function
    res = TA_SMA(ts, "ADJ_CLOSE", n)
    If IsArray(res) Then
        SMA = res(UBound(res))
    Else
        SMA = Null
    End If
End Function

Public Function ema(ByVal secCode As String, ByVal n As Long, _
                    Optional ByVal beginDate As Variant, Optional ByVal endDate As Variant) As Variant
    Dim ts As clsTimeSeriesFrame, res As Variant
    Set ts = GetStockFrame(secCode, beginDate, endDate, True)
    If ts Is Nothing Then ema = Null: Exit Function
    res = TA_EMA(ts, "ADJ_CLOSE", n)
    If IsArray(res) Then
        ema = res(UBound(res))
    Else
        ema = Null
    End If
End Function

Public Function macd(ByVal secCode As String, _
                     Optional ByVal fastPeriod As Long = 12, _
                     Optional ByVal slowPeriod As Long = 26, _
                     Optional ByVal signalPeriod As Long = 9, _
                     Optional ByVal beginDate As Variant, _
                     Optional ByVal endDate As Variant) As clsDataTable
    Dim ts As clsTimeSeriesFrame, mat As Variant, n As Long, i As Long
    Dim outArr() As Variant, dt As clsDataTable
    Set ts = GetStockFrame(secCode, beginDate, endDate, True)
    If ts Is Nothing Then Set macd = New clsDataTable: Exit Function
    mat = TA_MACD(ts, "ADJ_CLOSE", fastPeriod, slowPeriod, signalPeriod)
    If Not IsArray(mat) Then Set macd = New clsDataTable: Exit Function

    n = UBound(mat, 1)
    ReDim outArr(1 To n + 1, 1 To 4)
    outArr(1, 1) = "T_DATE": outArr(1, 2) = "MACD_DIF": outArr(1, 3) = "MACD_DEA": outArr(1, 4) = "MACD_HIST"
    For i = 1 To n
        outArr(i + 1, 1) = ts.RowDate(i)
        outArr(i + 1, 2) = mat(i, 1)
        outArr(i + 1, 3) = mat(i, 2)
        outArr(i + 1, 4) = mat(i, 3)
    Next i
    Set dt = New clsDataTable
    dt.Load outArr
    Set macd = dt
End Function

Public Function TechnicalIndicatorTable(ByVal secCode As String, _
                                       Optional ByVal beginDate As Variant, _
                                       Optional ByVal endDate As Variant) As clsDataTable
    Dim tsFull As clsTimeSeriesFrame, dStart As Date, dEnd As Date
    ResolveDates beginDate, endDate, dStart, dEnd
    Set tsFull = GetStockFrame(secCode, dStart, dEnd, True)
    If tsFull Is Nothing Then Set TechnicalIndicatorTable = New clsDataTable: Exit Function
    Set TechnicalIndicatorTable = TA_BuildIndicatorTable(tsFull, dStart, "ADJ_CLOSE")
End Function

Public Function ValuationSummary(ByVal secCode As String, _
                                 Optional ByVal asOfDate As Variant) As clsDataTable
    Dim dEnd As Date, dStart As Date, panel As clsEquityPanel
    Dim seriesList As Collection, series As clsDataTable
    If IsMissing(asOfDate) Or IsEmpty(asOfDate) Then dEnd = Date Else dEnd = CDate(asOfDate)
    dStart = DateAdd("yyyy", -3, dEnd)
    Set panel = m_EquityService.GetQuotePanel(Array(secCode), dStart, dEnd)
    Set seriesList = New Collection
    Set series = m_ValuationAnalyzer.BuildDailyLfyRatioSeries(panel, secCode)
    seriesList.Add series
    Set ValuationSummary = m_ValuationAnalyzer.BuildValuationSummary(seriesList, "ADJ_PE")

Public Function ValuationPE(ByVal secCode As String, Optional ByVal asOfDate As Variant) As Variant
    Dim sumTbl As clsDataTable, r As Long
    Set sumTbl = ValuationSummary(secCode, asOfDate)
    For r = 1 To sumTbl.rowCount
        If sumTbl.ValueAt(r, "METRIC_NAME") = "PE" Then
            ValuationPE = sumTbl.ValueAt(r, "LATEST_RATIO")
            Exit Function
        End If
    Next r
    ValuationPE = Null
End Function

Public Function ValuationPB(ByVal secCode As String, Optional ByVal asOfDate As Variant) As Variant
    Dim sumTbl As clsDataTable, r As Long
    Set sumTbl = ValuationSummary(secCode, asOfDate)
    For r = 1 To sumTbl.rowCount
        If sumTbl.ValueAt(r, "METRIC_NAME") = "PB" Then
            ValuationPB = sumTbl.ValueAt(r, "LATEST_RATIO")
            Exit Function
        End If
    Next r
    ValuationPB = Null
End Function

' ------------------------------------------------------------------------------
' 5. 综合报表一键交付 (Export Report)
' ------------------------------------------------------------------------------

Public Function ExportReport(ByVal secCode As String, _
                            Optional ByVal deliveryMode As Long = 2, _
                            Optional ByVal beginDate As Variant, _
                            Optional ByVal endDate As Variant) As Variant
    Dim dStart As Date, dEnd As Date
    Dim req As New clsAnalysisRequest, ctx As New clsAnalysisContext
    Dim res As New clsAnalysisResult
    Dim qTbl As clsDataTable, techTbl As clsDataTable, valTbl As clsDataTable

    ResolveDates beginDate, endDate, dStart, dEnd
    ctx.Initialize dEnd, dEnd, "CNY", "EQUITY_REPORT_V1"
    req.Initialize "EQUITY_REPORT", ctx
    req.SetParameter "SEC_CODE", secCode
    req.SetParameter "START_DATE", Format$(dStart, "yyyy-mm-dd")
    req.SetParameter "END_DATE", Format$(dEnd, "yyyy-mm-dd")

    res.Initialize req

    ' 装载三张核心报表
    Set qTbl = Quotes(secCode, dStart, dEnd, False)
    Set techTbl = TechnicalIndicatorTable(secCode, dStart, dEnd)
    Set valTbl = ValuationSummary(secCode, dEnd)

    res.AddTable "QUOTES", qTbl
    res.AddTable "INDICATORS", techTbl
    res.AddTable "VALUATION", valTbl

    Set ExportReport = m_Renderer.Render(res, deliveryMode, "EQ_" & CleanCode(secCode))
End Function

' ------------------------------------------------------------------------------
' 私有辅助函数
' ------------------------------------------------------------------------------

Private Function GetStockFrame(ByVal secCode As String, ByVal beginDate As Variant, ByVal endDate As Variant, _
                               ByVal includeWarmup As Boolean) As clsTimeSeriesFrame
    Dim panel As clsEquityPanel, dStart As Date, dEnd As Date
    ResolveDates beginDate, endDate, dStart, dEnd
    Set panel = m_EquityService.GetQuotePanel(Array(secCode), dStart, dEnd)
    Set GetStockFrame = panel.QuoteFrameForStock(secCode, includeWarmup)
End Function

Private Sub ResolveDates(ByVal beginDate As Variant, ByVal endDate As Variant, _
                         ByRef outStart As Date, ByRef outEnd As Date)
    If IsMissing(endDate) Or IsEmpty(endDate) Then outEnd = Date Else outEnd = CDate(endDate)
    If IsMissing(beginDate) Or IsEmpty(beginDate) Then outStart = DateAdd("yyyy", -1, outEnd) Else outStart = CDate(beginDate)
    If outStart > outEnd Then outStart = DateAdd("yyyy", -1, outEnd)
End Sub

Private Function CleanCode(ByVal code As String) As String
    Dim s As String
    s = Replace(code, ".", "_")
    s = Replace(s, ":", "_")
    CleanCode = s
End Function
