VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPortfolioFacade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' clsPortfolioFacade: 投资组合与时序分析统一领域门面 (Domain Facade)
' 统一收口与暴露所有组合分析功能，支持代码自动补全 (Portfolio.xxx)
' 涵盖持仓切片、行业穿透汇总、绩效收益时序、绝对归因、风险度量与组合报告导出
' ==============================================================================

Private m_PortService As clsPortfolioService
Private m_App As clsPortfolioAnalysisApplication
Private m_Renderer As clsResultRenderer

Private Sub Class_Initialize()
    Set m_PortService = New clsPortfolioService
    Set m_App = New clsPortfolioAnalysisApplication
    Set m_Renderer = New clsResultRenderer
End Sub

Public Function holdings(ByVal portfolioId As String, Optional ByVal asOfDate As Variant) As clsDataTable
    Dim panel As clsPortfolioPanel, dEnd As Date, dStart As Date, result As clsDataTable
    ResolveDates asOfDate, asOfDate, dStart, dEnd
    Set panel = m_PortService.GetPanel(Array(portfolioId), dStart, dEnd)
    Set result = New clsDataTable
    result.Load panel.holdingsFrame.ToTable()
    Set holdings = result
End Function

Public Function TotalValue(ByVal portfolioId As String, Optional ByVal asOfDate As Variant) As Variant
    Dim panel As clsPortfolioPanel, analyzer As clsPortfolioAnalyzer
    Dim dEnd As Date, dStart As Date
    ResolveDates asOfDate, asOfDate, dStart, dEnd
    Set panel = m_PortService.GetPanel(Array(portfolioId), dStart, dEnd)
    Set analyzer = New clsPortfolioAnalyzer
    analyzer.Initialize panel, True
    TotalValue = analyzer.TotalHoldingValue(dEnd)
End Function

Public Function AssetAllocation(ByVal portfolioId As String, Optional ByVal asOfDate As Variant) As clsDataTable
    Dim panel As clsPortfolioPanel, analyzer As clsPortfolioAnalyzer
    Dim dEnd As Date, dStart As Date, dict As Object, k As Variant
    Dim arr() As Variant, r As Long, dt As clsDataTable

    ResolveDates asOfDate, asOfDate, dStart, dEnd
    Set panel = m_PortService.GetPanel(Array(portfolioId), dStart, dEnd)
    Set analyzer = New clsPortfolioAnalyzer
    analyzer.Initialize panel, True

    Set dict = analyzer.HoldingValuesByAssetType()
    ReDim arr(1 To dict.count + 1, 1 To 2)
    arr(1, 1) = "ASSET_TYPE": arr(1, 2) = "HOLDING_VALUE"
    r = 1
    For Each k In dict.Keys
        r = r + 1
        arr(r, 1) = k
        arr(r, 2) = dict(k)
    Next k
    Set dt = New clsDataTable
    dt.Load arr
    Set AssetAllocation = dt
End Function

Public Function IndustrySummary(ByVal portfolioId As String, Optional ByVal asOfDate As Variant) As clsDataTable
    Dim panel As clsPortfolioPanel, ctx As clsAnalysisContext, res As clsAnalysisResult
    Dim dEnd As Date, dStart As Date
    ResolveDates asOfDate, asOfDate, dStart, dEnd

    Set panel = m_PortService.GetPanel(Array(portfolioId), dStart, dEnd)
    Set ctx = New clsAnalysisContext
    ctx.Initialize dEnd, dEnd, "CNY", "PORT_IND_V1"

    Set res = m_App.RunIndustryAnalysis(panel, ctx)
    Set IndustrySummary = res.table("INDUSTRY_ANALYSIS")
End Function

Public Function returns(ByVal portfolioId As String, ByVal beginDate As Variant, ByVal endDate As Variant) As clsDataTable
    Dim panel As clsPortfolioPanel, dStart As Date, dEnd As Date, result As clsDataTable
    ResolveDates beginDate, endDate, dStart, dEnd
    Set panel = m_PortService.GetPanel(Array(portfolioId), dStart, dEnd)
    Set result = New clsDataTable
    result.Load panel.performanceFrame.ToTable()
    Set returns = result
End Function

Public Function AttributionAbsolute(ByVal portfolioId As String, ByVal beginDate As Variant, ByVal endDate As Variant) As clsDataTable
    Dim panel As clsPortfolioPanel, ctx As clsAnalysisContext, res As clsAnalysisResult
    Dim dStart As Date, dEnd As Date
    ResolveDates beginDate, endDate, dStart, dEnd

    Set panel = m_PortService.GetPanel(Array(portfolioId), dStart, dEnd)
    Set ctx = New clsAnalysisContext
    ctx.Initialize dEnd, dEnd, "CNY", "PORT_ATTR_V1"

    Set res = m_App.RunAbsoluteAttribution(panel, ctx)
    Set AttributionAbsolute = res.table("ABSOLUTE_ATTRIBUTION")
End Function

Public Function SharpeRatio(ByVal portfolioId As String, ByVal beginDate As Variant, ByVal endDate As Variant, _
                            Optional ByVal riskFreeRate As Double = 0.02) As Variant
    Dim panel As clsPortfolioPanel, dStart As Date, dEnd As Date
    ResolveDates beginDate, endDate, dStart, dEnd
    Set panel = m_PortService.GetPanel(Array(portfolioId), dStart, dEnd)
    If panel.performanceFrame.rowCount = 0 Then SharpeRatio = Null: Exit Function
    SharpeRatio = TS_SharpeRatio(panel.performanceFrame, "RTN", "", 252#) - (riskFreeRate * Sqr(252#))
End Function

Public Function maxDrawdown(ByVal portfolioId As String, ByVal beginDate As Variant, ByVal endDate As Variant) As Variant
    Dim panel As clsPortfolioPanel, dStart As Date, dEnd As Date
    ResolveDates beginDate, endDate, dStart, dEnd
    Set panel = m_PortService.GetPanel(Array(portfolioId), dStart, dEnd)
    If panel.performanceFrame.rowCount = 0 Then maxDrawdown = Null: Exit Function
    maxDrawdown = TS_MaxDrawdown(panel.performanceFrame, "RTN")
End Function

Public Function Volatility(ByVal portfolioId As String, ByVal beginDate As Variant, ByVal endDate As Variant) As Variant
    Dim panel As clsPortfolioPanel, dStart As Date, dEnd As Date
    ResolveDates beginDate, endDate, dStart, dEnd
    Set panel = m_PortService.GetPanel(Array(portfolioId), dStart, dEnd)
    If panel.performanceFrame.rowCount = 0 Then Volatility = Null: Exit Function
    Volatility = TS_Volatility(panel.performanceFrame, "RTN")
End Function

Public Function ExportReport(ByVal portfolioId As String, _
                            Optional ByVal deliveryMode As Long = 2, _
                            Optional ByVal asOfDate As Variant) As Variant
    Dim dEnd As Date, dStart As Date, panel As clsPortfolioPanel
    Dim ctx As New clsAnalysisContext, req As New clsAnalysisRequest, res As New clsAnalysisResult
    Dim indTbl As clsDataTable, attrTbl As clsDataTable, holdingsTbl As clsDataTable

    ResolveDates asOfDate, asOfDate, dStart, dEnd
    dStart = DateAdd("m", -1, dEnd)

    ctx.Initialize dEnd, dEnd, "CNY", "PORT_REPORT_V1"
    req.Initialize "PORTFOLIO_REPORT", ctx
    req.SetParameter "PORTFOLIO_ID", portfolioId
    req.SetParameter "AS_OF_DATE", Format$(dEnd, "yyyy-mm-dd")

    res.Initialize req

    Set panel = m_PortService.GetPanel(Array(portfolioId), dStart, dEnd)
    Set indTbl = IndustrySummary(portfolioId, dEnd)
    Set attrTbl = AttributionAbsolute(portfolioId, dStart, dEnd)
    Set holdingsTbl = New clsDataTable
    holdingsTbl.Load panel.holdingsFrame.ToTable()

    res.AddTable "HOLDINGS", holdingsTbl
    res.AddTable "INDUSTRY", indTbl
    res.AddTable "ATTRIBUTION", attrTbl

    Set ExportReport = m_Renderer.Render(res, deliveryMode, "PORT_" & portfolioId)
End Function

' ------------------------------------------------------------------------------
' 私有日期辅助函数
' ------------------------------------------------------------------------------

Private Sub ResolveDates(ByVal beginDate As Variant, ByVal endDate As Variant, _
                         ByRef outStart As Date, ByRef outEnd As Date)
    If IsMissing(endDate) Or IsEmpty(endDate) Then outEnd = Date Else outEnd = CDate(endDate)
    If IsMissing(beginDate) Or IsEmpty(beginDate) Then outStart = DateAdd("m", -1, outEnd) Else outStart = CDate(beginDate)
    If outStart > outEnd Then outStart = DateAdd("m", -1, outEnd)
End Sub
