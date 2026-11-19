Attribute VB_Name = "Mod_PortfolioTests"
Option Explicit

Public Function RunPortfolioOfflineTests() As String
    Dim h As Variant, p As Variant, profit As Variant, orders As Variant, industries As Variant
    Dim panel As clsPortfolioPanel, analyzer As clsPortfolioAnalyzer
    Dim report As Workbook, returnReport As Workbook, count As Long, attribution As Variant, returns As Variant, industryOutput As Variant
    Dim seriesByPort As Object, portfolioSeries As clsTimeSeriesFrame
    Dim context As clsAnalysisContext, app As clsPortfolioAnalysisApplication, analysisResult As clsAnalysisResult
    On Error GoTo Failed
    h = HoldingsFixture()
    p = PerformanceFixture()
    profit = ProfitFixture()
    orders = EmptyFixture(Array("PORT_CODE", "T_DATE", "ORDER_ID"))
    industries = IndustryFixture()
    Set panel = BuildPanel(h, p, profit, orders, industries)
    RunDataTableTests count
    RunPanelContractTests panel, count
    RunEmptyPanelTests count
    RunTimeSeriesFrameTests count
    Check panel.HoldingRowCount = 1, count: Check panel.PerformanceRowCount = 6, count
    Check panel.PortfolioCodes = "'001','002'", count: Check panel.startDate = DateSerial(2026, 11, 10), count
    Check panel.endDate = DateSerial(2026, 11, 11), count
    Set analyzer = New clsPortfolioAnalyzer: analyzer.Initialize panel
    Check panel.performanceFrame.rowCount = panel.PerformanceRowCount, count
    Check panel.holdingsFrame.rowCount = panel.HoldingRowCount, count
    Check panel.profitFrame.rowCount = panel.ProfitRowCount, count
    Check panel.ordersData.rowCount = panel.OrderRowCount, count
    Set seriesByPort = analyzer.BuildPortfolioReturnFrames()
    Set portfolioSeries = seriesByPort("001")
    Check portfolioSeries.rowCount = 2, count
    Check CStr(portfolioSeries.ValueAt(1, "PORT_CODE")) = "001", count
    Check CDbl(portfolioSeries.ValueAt(2, "P_TOTALNAV")) = 110#, count
    Check Abs(CDbl(portfolioSeries.EvaluateMetric("RETURN_RATE", "RTN", "RTN_BM")) - 0.32) < 0.000001, count
    attribution = analyzer.AbsoluteAttribution
    Check Abs(FindAttribution(attribution, "001", "SECURITY", "000001") - 0.34) < 0.000001, count
    Check Abs(FindAttribution(attribution, "002", "SECURITY", "000001") - 0.115) < 0.000001, count
    Check Abs(FindAttribution(attribution, "001", "INDUSTRY", "S1") - 0.12) < 0.000001, count
    Check Abs(FindAttribution(attribution, "001", "INDUSTRY", "S2") - 0.22) < 0.000001, count
    Check Abs(FindAttribution(attribution, "002", "INDUSTRY", "S1") + 0.105) < 0.000001, count
    Check Abs(FindAttribution(attribution, "002", "INDUSTRY", "S2") - 0.22) < 0.000001, count
    industryOutput = analyzer.IndustryAnalysis
    Check UBound(industryOutput, 1) = 2 And industryOutput(2, 3) = "S1", count
    Check Abs(CDbl(industryOutput(2, 5)) - 25#) < 0.000001, count
    Check Abs(CDbl(industryOutput(2, 7)) - 9#) < 0.000001, count
    Check industryOutput(2, 11) = "OK", count
    returns = analyzer.ReturnAnalysis
    Check Abs(FindReturn(returns, "001", "2026", 6) - 0.32) < 0.000001, count
    Set context = New clsAnalysisContext: context.Initialize panel.endDate, panel.endDate, "CNY", "PORTFOLIO_RETURN_V1"
    Set app = New clsPortfolioAnalysisApplication
    Set analysisResult = app.RunReturnAnalysis(panel, context)
    ' Portfolio "001" has strictly positive returns (+0.1, +0.2), so Sortino downside RMS is 0 -> UNDEFINED_RISK_RATIO warning.
    Check analysisResult.status = "PARTIAL", count
    Check analysisResult.table("RETURNS_ANALYSIS").rowCount = UBound(returns, 1) - 1, count
    Check analysisResult.table("LINEAGE").rowCount = panel.SourceDataSets.count, count
    Check analysisResult.table("DIAGNOSTICS").rowCount = 1, count
    Set analysisResult = app.RunIndustryAnalysis(panel, context)
    Check analysisResult.status = "SUCCESS", count
    Check analysisResult.table("INDUSTRY_ANALYSIS").rowCount = UBound(industryOutput, 1) - 1, count
    Check analysisResult.table("DIAGNOSTICS").rowCount = 0, count
    Set analysisResult = app.RunAbsoluteAttribution(panel, context)
    Check analysisResult.status = "SUCCESS", count
    Check analysisResult.table("ABSOLUTE_ATTRIBUTION").rowCount = UBound(attribution, 1) - 1, count
    Set analysisResult = app.RunHoldingSummary(panel, context, True)
    Check analysisResult.table("HOLDING_SUMMARY").rowCount = 1, count
    Check analysisResult.table("HOLDING_SUMMARY").ValueAt(1, "METRIC") = "TOTAL_HOLDING_VALUE", count
    Check analysisResult.table("HOLDING_SUMMARY").ValueAt(1, "VALUE") = 25, count
    Check Abs(FindReturn(returns, "001", "2026", 7) - 0.0098) < 0.000001, count
    Check Abs(FindReturn(returns, "001", "2026", 8) - 0.3102) < 0.000001, count
    Check Abs(FindReturn(returns, "001", "2026", 9) - (Sqr(0.005) * Sqr(252#))) < 0.001, count
    Check Abs(FindReturn(returns, "001", "2026", 10) - ((0.15 / Sqr(0.005)) * Sqr(252#))) < 0.001, count
    analyzer.Initialize panel, True: Check Abs(analyzer.HoldingWeight(1) - 0.25) < 0.000001, count
    Check analyzer.TotalHoldingValue = 25, count
    Set report = RenderPortfolioReport(panel, True)
    Check report.Worksheets.count = 8, count
    Check report.Worksheets("AbsoluteAttribution").Cells(1, 5).Value2 = "ABSOLUTE_CONTRIBUTION", count
    Check report.Worksheets("ReturnsAnalysis").Cells(1, 9).Value2 = "ANNUALIZED_VOLATILITY", count
    report.Close False: Set report = Nothing
    Set returnReport = RenderPortfolioReturnReport(panel)
    Check returnReport.Worksheets.count = 2, count
    Check returnReport.Worksheets("ReturnsAnalysis").Cells(1, 6).Value2 = "RETURN_RATE", count
    returnReport.Close False
    RunPortfolioOfflineTests = "PASS: " & count & " assertions; no database connections."
    Exit Function
Failed:
    If Not report Is Nothing Then report.Close False
    If Not returnReport Is Nothing Then returnReport.Close False
    Err.Raise Err.number, "Mod_PortfolioTests", "Offline assertion " & CStr(count + 1) & " failed: " & Err.description
End Function

Private Sub RunDataTableTests(ByRef count As Long)
    Dim t(1 To 4, 1 To 3) As Variant, bad(1 To 1, 1 To 2) As Variant
    Dim threeDimensional(1 To 1, 1 To 1, 1 To 1) As Variant
    Dim data As clsDataTable, selected As clsDataTable, groups As Object, output As Variant, errorNumber As Long
    Dim frame As clsTimeSeriesFrame, emptyRows As New Collection
    t(1, 1) = " T_DATE ": t(1, 2) = "PORT_CODE": t(1, 3) = "VALUE"
    t(2, 1) = DateSerial(2026, 11, 10): t(2, 2) = "001": t(2, 3) = Null
    t(3, 1) = DateSerial(2026, 11, 11): t(3, 2) = "001": t(3, 3) = 0#
    t(4, 1) = DateSerial(2026, 11, 11): t(4, 2) = "002": t(4, 3) = Empty
    Set data = New clsDataTable: data.Load t
    Check data.IsLoaded And data.rowCount = 3 And data.fieldCount = 3, count
    Check data.FieldIndex("t_date") = 1 And data.HasField("port_code"), count
    Check Not data.HasField("MISSING"), count
    Check IsNull(data.ValueAt(1, "VALUE")), count
    Check data.ValueAt(2, "VALUE") = 0# And IsEmpty(data.ValueAt(3, "VALUE")), count
    output = data.ToTable(): output(2, 2) = "CHANGED": t(2, 2) = "CHANGED"
    Check data.ValueAt(1, "PORT_CODE") = "001", count
    Set selected = data.FilterField("PORT_CODE", "001")
    Check selected.rowCount = 2, count
    Set selected = data.SelectRows(emptyRows)
    Check selected.rowCount = 0 And selected.fieldCount = 3, count
    Set groups = data.GroupRows(Array("PORT_CODE"))
    Check groups.count = 2, count
    Set groups = data.UniqueRowIndex(Array("T_DATE", "PORT_CODE"))
    Check groups.count = 3, count
    Set frame = New clsTimeSeriesFrame: frame.LoadData data, "T_DATE"
    Check frame.rowCount = 3 And frame.FieldIndex("VALUE") = 3, count
    Set selected = New clsDataTable: selected.Load IndustryFixture()
    Check Not selected.HasField("T_DATE") And selected.rowCount = 2, count
    bad(1, 1) = "NAME": bad(1, 2) = "name"
    On Error Resume Next
    Err.Clear
    Set selected = New clsDataTable: selected.Load bad
    errorNumber = Err.number
    On Error GoTo 0
    Check errorNumber = 5, count
    On Error Resume Next
    Err.Clear
    data.Load t
    errorNumber = Err.number
    On Error GoTo 0
    Check errorNumber = 5, count
    On Error Resume Next
    Err.Clear
    Set selected = New clsDataTable: selected.Load Array(1, 2)
    errorNumber = Err.number
    On Error GoTo 0
    Check errorNumber = 5, count
    On Error Resume Next
    Err.Clear
    Set selected = New clsDataTable: selected.Load threeDimensional
    errorNumber = Err.number
    On Error GoTo 0
    Check errorNumber = 5, count
End Sub

Private Sub RunEmptyPanelTests(ByRef count As Long)
    Dim panel As clsPortfolioPanel, analyzer As clsPortfolioAnalyzer, output As Variant
    Set panel = BuildPanel(EmptyFixture(Array("PORT_CODE", "T_DATE", "I_CODE", "H_EVAL", "H_AI")), _
        EmptyFixture(Array("PORT_CODE", "T_DATE", "P_TOTALNAV", "RTN", "RTN_BM")), _
        EmptyFixture(Array("PORT_CODE", "T_DATE", "I_CODE", "DAILYPROFIT")), _
        EmptyFixture(Array("PORT_CODE", "T_DATE", "ORDER_ID")), _
        EmptyFixture(Array("SEC_CODE", "START_DATE", "EXPIRATION_DATE", "SW_SECTOR_CODE", "SW_SECTOR")))
    Set analyzer = New clsPortfolioAnalyzer
    analyzer.Initialize panel
    output = analyzer.IndustryAnalysis
    Check UBound(output, 1) = 1 And UBound(output, 2) = 11, count
    output = analyzer.AbsoluteAttribution
    Check UBound(output, 1) = 1 And UBound(output, 2) = 6, count
    output = analyzer.ReturnAnalysis
    Check UBound(output, 1) = 1 And UBound(output, 2) = 12, count
End Sub

Private Sub RunPanelContractTests(ByVal panel As clsPortfolioPanel, ByRef count As Long)
    Dim other As clsPortfolioPanel, emptyFrame As clsTimeSeriesFrame, emptyTable As clsDataTable
    Dim h As clsTimeSeriesFrame, p As clsTimeSeriesFrame, profit As clsTimeSeriesFrame
    Dim orders As clsDataTable, industries As clsDataTable, errorNumber As Long, table As Variant
    Set h = panel.holdingsFrame: Set p = panel.performanceFrame: Set profit = panel.profitFrame
    Set orders = panel.ordersData: Set industries = panel.industryData
    Check orders.rowCount = 0 And orders.fieldCount = 3 And industries.rowCount = 2, count
    table = panel.HoldingsTable: table(2, 1) = "CHANGED"
    Check h.ValueAt(1, "PORT_CODE") = "001", count
    On Error Resume Next
    Err.Clear
    panel.Load "x", panel.startDate, panel.endDate, "OFFLINE", h, p, profit, orders, industries
    errorNumber = Err.number
    On Error GoTo 0
    Check errorNumber = 5, count
    Set other = New clsPortfolioPanel: Set emptyFrame = New clsTimeSeriesFrame
    Set emptyTable = New clsDataTable
    On Error Resume Next
    Err.Clear
    other.Load "x", panel.startDate, panel.endDate, "OFFLINE", emptyFrame, p, profit, orders, industries
    errorNumber = Err.number
    On Error GoTo 0
    Check errorNumber = 5, count
    emptyFrame.Load EmptyFixture(Array("PORT_CODE", "T_DATE", "RTN", "RTN_BM", "P_TOTALNAV")), "T_DATE"
    emptyTable.Load EmptyFixture(Array("SEC_CODE", "START_DATE", "EXPIRATION_DATE", "SW_SECTOR_CODE", "SW_SECTOR"))
    other.Load "x", panel.startDate, panel.endDate, "OFFLINE", emptyFrame, emptyFrame, emptyFrame, orders, emptyTable
    Check other.HoldingRowCount = 0 And other.PerformanceRowCount = 0 And other.IndustryRowCount = 0, count
End Sub

Private Sub RunTimeSeriesFrameTests(ByRef count As Long)
    Dim table() As Variant, frame As clsTimeSeriesFrame, filtered As clsTimeSeriesFrame, sorted As clsTimeSeriesFrame
    Dim returnsTable() As Variant, returnsFrame As clsTimeSeriesFrame, compounded As clsTimeSeriesFrame
    Dim aggregate As clsTimeSeriesFrame, singleSeries As clsTimeSeriesFrame
    Dim groups As Object, uniqueRows As Object, rows As Collection, key As Variant
    Dim r As Long, compounded001 As Double, metricValue As Variant, groupedMetrics As Variant, failureNumber As Long
    Dim found001 As Boolean, found002 As Boolean
    ReDim table(1 To 6, 1 To 3)
    table(1, 1) = "T_DATE": table(1, 2) = "PORT_CODE": table(1, 3) = "VALUE"
    table(2, 1) = DateSerial(2026, 11, 11): table(2, 2) = "001": table(2, 3) = 10
    table(3, 1) = DateSerial(2026, 11, 10): table(3, 2) = "001": table(3, 3) = 20
    table(4, 1) = DateSerial(2026, 11, 11): table(4, 2) = "001": table(4, 3) = 30
    table(5, 1) = DateSerial(2026, 11, 11): table(5, 2) = "002": table(5, 3) = 40
    table(6, 1) = Empty: table(6, 2) = "003": table(6, 3) = 50
    Set frame = New clsTimeSeriesFrame
    frame.Load table, "T_DATE"
    Check frame.rowCount = 5, count
    Check frame.FieldIndex("port_code") = 2, count
    Check Not frame.HasDate(5), count
    Check IsNull(frame.RowDate(5)), count
    Set rows = frame.RowsAtDate(DateSerial(2026, 11, 11))
    Check rows.count = 3, count
    Check CDbl(frame.ValueAt(CLng(rows(1)), "VALUE")) = 10, count
    Set filtered = frame.FilterDateRange(DateSerial(2026, 11, 10), DateSerial(2026, 11, 10))
    Check filtered.rowCount = 1, count
    Check filtered.ValueAt(1, "PORT_CODE") = "001", count
    Set filtered = frame.FilterDateRange(DateSerial(2026, 11, 10), DateSerial(2026, 11, 11))
    Set aggregate = filtered.AggregateByDate(Array("PORT_CODE"), "VALUE", "SUM")
    For r = 1 To aggregate.rowCount
        If aggregate.ValueAt(r, "PORT_CODE") = "001" And CDate(aggregate.RowDate(r)) = DateSerial(2026, 11, 11) Then Check CDbl(aggregate.ValueAt(r, "VALUE")) = 40, count
    Next r
    Check aggregate.rowCount = 3, count
    Set groups = frame.GroupRows(Array("PORT_CODE"))
    Check groups.count = 3, count
    Set rows = Nothing
    For Each key In groups.Keys
        Set rows = groups(key)
        If rows.count = 3 Then Exit For
        Set rows = Nothing
    Next key
    Check Not rows Is Nothing, count
    Set sorted = frame.SortedByDate()
    Check CDate(sorted.RowDate(1)) = DateSerial(2026, 11, 10), count
    Check CDate(sorted.RowDate(2)) = DateSerial(2026, 11, 11), count
    Check CDate(sorted.RowDate(3)) = DateSerial(2026, 11, 11), count
    Check CDate(sorted.RowDate(4)) = DateSerial(2026, 11, 11), count
    Check Not sorted.HasDate(5), count
    On Error Resume Next
    Err.Clear
    Set uniqueRows = frame.UniqueRowIndex(Array("T_DATE", "PORT_CODE", "VALUE"))
    failureNumber = Err.number
    On Error GoTo 0
    Call Check(failureNumber = 0, count)
    Check uniqueRows.count = 5, count
    On Error Resume Next
    Err.Clear
    Set uniqueRows = frame.UniqueRowIndex(Array("PORT_CODE"))
    failureNumber = Err.number
    On Error GoTo 0
    Call Check(failureNumber = 5, count)
    Err.Clear
    ReDim returnsTable(1 To 5, 1 To 3)
    returnsTable(1, 1) = "T_DATE": returnsTable(1, 2) = "PORT_CODE": returnsTable(1, 3) = "RETURN"
    returnsTable(2, 1) = DateSerial(2026, 11, 11): returnsTable(2, 2) = "001": returnsTable(2, 3) = 0.2
    returnsTable(3, 1) = DateSerial(2026, 11, 10): returnsTable(3, 2) = "001": returnsTable(3, 3) = 0.1
    returnsTable(4, 1) = DateSerial(2026, 11, 11): returnsTable(4, 2) = "002": returnsTable(4, 3) = -0.1
    returnsTable(5, 1) = DateSerial(2026, 11, 10): returnsTable(5, 2) = "002": returnsTable(5, 3) = 0#
    Set returnsFrame = New clsTimeSeriesFrame
    returnsFrame.Load returnsTable, "T_DATE"
    Set compounded = returnsFrame.CompoundSeries(Array("PORT_CODE"), "RETURN")
    For r = 1 To compounded.rowCount
        If compounded.ValueAt(r, "PORT_CODE") = "001" Then compounded001 = CDbl(compounded.ValueAt(r, "RETURN_COMPOUNDED"))
    Next r
    Check compounded.rowCount = 2, count
    Check Abs(compounded001 - 0.32) < 0.000001, count
    Set singleSeries = returnsFrame.FilterField("PORT_CODE", "001")
    metricValue = singleSeries.EvaluateMetric("AVERAGE", "RETURN")
    Check Abs(CDbl(metricValue) - 0.15) < 0.000001, count
    metricValue = singleSeries.EvaluateMetric("CUMULATIVE_RETURN", "RETURN")
    Check Abs(CDbl(metricValue) - 0.32) < 0.000001, count
    metricValue = singleSeries.EvaluateMetric("VOLATILITY", "RETURN")
    Check Abs(CDbl(metricValue) - (Sqr(0.005) * Sqr(252#))) < 0.001, count
    groupedMetrics = returnsFrame.EvaluateMetricByGroup("AVERAGE", Array("PORT_CODE"), "RETURN")
    Check UBound(groupedMetrics, 1) = 3, count
    For r = 2 To UBound(groupedMetrics, 1)
        If groupedMetrics(r, 1) = "001" Then found001 = Abs(CDbl(groupedMetrics(r, 3)) - 0.15) < 0.000001
        If groupedMetrics(r, 1) = "002" Then found002 = Abs(CDbl(groupedMetrics(r, 3)) + 0.05) < 0.000001
    Next r
    Check found001, count: Check found002, count
    RunFrameMigrationTests count
End Sub

Private Sub RunFrameMigrationTests(ByRef count As Long)
    Dim table(1 To 4, 1 To 6) As Variant, frame As clsTimeSeriesFrame
    Dim oneYear As clsTimeSeriesFrame, emptyFrame As clsTimeSeriesFrame, slice As clsTimeSeriesFrame
    Dim years As Object, metric As Variant, indices(0 To 1) As Long, unusedIndices() As Long
    Dim failureNumber As Long
    table(1, 1) = "DAY": table(1, 2) = "PORT_CODE": table(1, 3) = "RTN"
    table(1, 4) = "BM": table(1, 5) = "RF": table(1, 6) = "EXPOSURE"
    table(2, 1) = DateSerial(2027, 1, 2): table(2, 2) = "001": table(2, 3) = 0.05
    table(2, 4) = 0.02: table(2, 5) = 0.01: table(2, 6) = 30#
    table(3, 1) = DateSerial(2026, 12, 31): table(3, 2) = "001": table(3, 3) = -0.1
    table(3, 4) = 0#: table(3, 5) = 0.01: table(3, 6) = 10#
    table(4, 1) = DateSerial(2027, 1, 1): table(4, 2) = "001": table(4, 3) = -0.1
    table(4, 4) = 0.01: table(4, 5) = 0.01: table(4, 6) = 20#
    Set frame = New clsTimeSeriesFrame
    frame.Load table, "DAY"
    Check CDate(frame.startDate) = DateSerial(2026, 12, 31), count
    Check CDate(frame.endDate) = DateSerial(2027, 1, 2), count
    Set years = frame.GroupByCalendarYear()
    Check years.count = 2, count
    Set oneYear = years("2027")
    Check oneYear.rowCount = 2, count
    metric = oneYear.EvaluateMetric("RETURN_RATE", "RTN", "BM")
    Check Abs(CDbl(metric) + 0.055) < 0.000001, count
    metric = oneYear.EvaluateMetric("BENCHMARK_RETURN", "RTN", "BM")
    Check Abs(CDbl(metric) - 0.0302) < 0.000001, count
    metric = oneYear.EvaluateMetric("EXCESS_RETURN", "RTN", "BM")
    Check Abs(CDbl(metric) + 0.0852) < 0.000001, count
    metric = oneYear.EvaluateMetric("MAX_DRAWDOWN", "RTN")
    Check Abs(CDbl(metric) + 0.1) < 0.000001, count
    metric = oneYear.EvaluateMetric("TRACKING_ERROR", "RTN", "BM", periodsPerYear:=1#)
    Check Abs(CDbl(metric) - Sqr(0.0098)) < 0.000001, count
    metric = oneYear.EvaluateMetric("INFORMATION_RATIO", "RTN", "BM", periodsPerYear:=1#)
    Check Abs(CDbl(metric) - (-0.04 / Sqr(0.0098))) < 0.000001, count
    metric = oneYear.EvaluateMetric("SHARPE_RATIO", "RTN", riskFreeField:="RF", periodsPerYear:=1#)
    Check Abs(CDbl(metric) - (-0.035 / Sqr(0.01125))) < 0.000001, count
    metric = oneYear.EvaluateMetric("SORTINO_RATIO", "RTN", riskFreeField:="RF", periodsPerYear:=1#)
    Check Abs(CDbl(metric) - (-0.035 / Sqr(0.005))) < 0.000001, count
    metric = oneYear.EvaluateMetric("ANNUALIZED_RETURN", "RTN", periodsPerYear:=2#)
    Check Abs(CDbl(metric) + 0.055) < 0.000001, count
    metric = oneYear.EvaluateMetric("CALMAR_RATIO", "RTN", periodsPerYear:=2#)
    Check Abs(CDbl(metric) + 0.55) < 0.000001, count
    metric = frame.EvaluateMetric("AVERAGE", "EXPOSURE")
    Check CDbl(metric) = 20#, count
    indices(0) = 1: indices(1) = 2
    Set slice = frame.SliceByIndices(indices, 2)
    Check slice.rowCount = 2 And CDate(slice.RowDate(1)) = DateSerial(2026, 12, 31), count
    Check CDbl(slice.ValueAt(1, "EXPOSURE")) = 10#, count
    Set emptyFrame = frame.SliceByIndices(unusedIndices, 0)
    Check emptyFrame.rowCount = 0, count
    Check emptyFrame.DateField = "DAY" And emptyFrame.fieldCount = 6, count
    Check IsNull(emptyFrame.startDate) And IsNull(emptyFrame.endDate), count
    Check IsNull(emptyFrame.EvaluateMetric("RETURN_RATE", "RTN")), count
    table(4, 1) = DateSerial(2027, 1, 2) + TimeSerial(12, 0, 0)
    Set frame = New clsTimeSeriesFrame
    frame.Load table, "DAY"
    On Error Resume Next
    Err.Clear
    frame.ValidateSingleSeries
    failureNumber = Err.number
    On Error GoTo 0
    Call Check(failureNumber = 5, count)
    table(4, 1) = DateSerial(2027, 1, 1): table(4, 3) = Null
    Set frame = New clsTimeSeriesFrame
    frame.Load table, "DAY"
    Check IsNull(frame.EvaluateMetric("CALMAR_RATIO", "RTN")), count
End Sub

Private Function HoldingsFixture() As Variant
    Dim t(1 To 2, 1 To 14) As Variant
    t(1, 1) = "PORT_CODE": t(1, 2) = "T_DATE": t(1, 3) = "I_CODE": t(1, 4) = "M_TYPE": t(1, 5) = "I_NAME": t(1, 6) = "A_TYPE": t(1, 7) = "CURRENCY"
    t(1, 8) = "H_COUNT": t(1, 9) = "H_COST": t(1, 10) = "H_EVAL": t(1, 11) = "H_AI": t(1, 12) = "H_VALUE_WITH_AI": t(1, 13) = "UNIT_PRICE": t(1, 14) = "H_NAV_PCT"
    t(2, 1) = "001": t(2, 2) = DateSerial(2026, 11, 10): t(2, 3) = "000001": t(2, 4) = "M1": t(2, 5) = "=1+1": t(2, 6) = "BOND": t(2, 7) = "CNY"
    t(2, 10) = 25#: t(2, 12) = 25#: t(2, 14) = 0.25
    HoldingsFixture = t
End Function

Private Function PerformanceFixture() As Variant
    Dim t(1 To 7, 1 To 9) As Variant
    t(1, 1) = "PORT_CODE": t(1, 2) = "T_DATE": t(1, 3) = "P_UNITNAV": t(1, 4) = "P_CUMU_NAV": t(1, 5) = "P_TOTALNAV": t(1, 6) = "P_TOTAL_ASSET": t(1, 7) = "P_SIZE": t(1, 8) = "RTN": t(1, 9) = "RTN_BM"
    AddPerformanceRow t, 2, "001", DateSerial(2026, 11, 7), 90#, 0#, 0#
    AddPerformanceRow t, 3, "001", DateSerial(2026, 11, 10), 100#, 0.1, 0.02
    AddPerformanceRow t, 4, "001", DateSerial(2026, 11, 11), 110#, 0.2, -0.01
    AddPerformanceRow t, 5, "002", DateSerial(2026, 11, 7), 180#, 0#, 0#
    AddPerformanceRow t, 6, "002", DateSerial(2026, 11, 10), 200#, -0.1, 0.01
    AddPerformanceRow t, 7, "002", DateSerial(2026, 11, 11), 220#, 0.05, 0.03
    PerformanceFixture = t
End Function

Private Sub AddPerformanceRow(ByRef t() As Variant, ByVal r As Long, ByVal port As String, ByVal day As Date, ByVal nav As Double, ByVal rtn As Double, ByVal bm As Double)
    t(r, 1) = port: t(r, 2) = day: t(r, 5) = nav: t(r, 8) = rtn: t(r, 9) = bm
End Sub

Private Function ProfitFixture() As Variant
    Dim t(1 To 5, 1 To 4) As Variant
    t(1, 1) = "PORT_CODE": t(1, 2) = "T_DATE": t(1, 3) = "I_CODE": t(1, 4) = "DAILYPROFIT"
    AddProfitRow t, 2, "001", DateSerial(2026, 11, 10), "000001", 9#
    AddProfitRow t, 3, "001", DateSerial(2026, 11, 11), "000001", 22#
    AddProfitRow t, 4, "002", DateSerial(2026, 11, 10), "000001", -18#
    AddProfitRow t, 5, "002", DateSerial(2026, 11, 11), "000001", 44#
    ProfitFixture = t
End Function

Private Sub AddProfitRow(ByRef t() As Variant, ByVal r As Long, ByVal port As String, ByVal day As Date, ByVal sec As String, ByVal amount As Double)
    t(r, 1) = port: t(r, 2) = day: t(r, 3) = sec: t(r, 4) = amount
End Sub

Private Function IndustryFixture() As Variant
    Dim t(1 To 3, 1 To 5) As Variant
    t(1, 1) = "SEC_CODE": t(1, 2) = "START_DATE": t(1, 3) = "EXPIRATION_DATE": t(1, 4) = "SW_SECTOR_CODE": t(1, 5) = "SW_SECTOR"
    t(2, 1) = "000001": t(2, 2) = DateSerial(2026, 1, 1): t(2, 3) = DateSerial(2026, 11, 10): t(2, 4) = "S1": t(2, 5) = "Sector 1"
    t(3, 1) = "000001": t(3, 2) = DateSerial(2026, 11, 11): t(3, 3) = DateSerial(2026, 12, 31): t(3, 4) = "S2": t(3, 5) = "Sector 2"
    IndustryFixture = t
End Function

Private Function EmptyFixture(ByVal names As Variant) As Variant
    Dim result() As Variant, i As Long
    ReDim result(1 To 1, 1 To UBound(names) - LBound(names) + 1)
    For i = LBound(names) To UBound(names): result(1, i - LBound(names) + 1) = CStr(names(i)): Next i
    EmptyFixture = result
End Function

Private Function FindAttribution(ByVal table As Variant, ByVal port As String, ByVal level As String, ByVal code As String) As Double
    Dim r As Long
    For r = 2 To UBound(table, 1)
        If CStr(table(r, 1)) = port And CStr(table(r, 2)) = level And CStr(table(r, 3)) = code Then FindAttribution = CDbl(table(r, 5)): Exit Function
    Next r
    Err.Raise 5, , "Missing attribution row: " & port & " / " & level & " / " & code
End Function

Private Function FindReturn(ByVal table As Variant, ByVal port As String, ByVal period As String, ByVal column As Long) As Double
    Dim r As Long
    For r = 2 To UBound(table, 1)
        If CStr(table(r, 1)) = port And CStr(table(r, 2)) = period Then FindReturn = CDbl(table(r, column)): Exit Function
    Next r
    Err.Raise 5, , "Missing return row: " & port & " / " & period
End Function

Private Function BuildPanel(ByVal h As Variant, ByVal p As Variant, ByVal profit As Variant, ByVal orders As Variant, ByVal industries As Variant) As clsPortfolioPanel
    Dim holdingsFrame As clsTimeSeriesFrame, performanceFrame As clsTimeSeriesFrame, profitFrame As clsTimeSeriesFrame
    Dim ordersTable As clsDataTable, industryTable As clsDataTable, panel As clsPortfolioPanel
    Dim sources As New Collection
    Set holdingsFrame = New clsTimeSeriesFrame: holdingsFrame.Load h, "T_DATE"
    Set performanceFrame = New clsTimeSeriesFrame: performanceFrame.Load p, "T_DATE"
    Set profitFrame = New clsTimeSeriesFrame: profitFrame.Load profit, "T_DATE"
    Set ordersTable = New clsDataTable: ordersTable.Load orders
    Set industryTable = New clsDataTable: industryTable.Load industries

    sources.Add MakeFixtureDataSet(h, "PORT_HOLDINGS", "PORT_CODE+T_DATE+I_CODE", "PORT_CODE")
    sources.Add MakeFixtureDataSet(p, "PORT_PERFORMANCE", "PORT_CODE+T_DATE", "PORT_CODE")
    sources.Add MakeFixtureDataSet(profit, "PORT_PROFIT", "PORT_CODE+T_DATE+I_CODE", "PORT_CODE")
    sources.Add MakeFixtureDataSet(orders, "PORT_ORDERS", "PORT_CODE+T_DATE+ORDER_ID", "PORT_CODE")
    sources.Add MakeFixtureDataSet(industries, "PORT_INDUSTRY", "SEC_CODE+START_DATE", "SEC_CODE")

    Set panel = New clsPortfolioPanel
    panel.Load "'001','002'", DateSerial(2026, 11, 10), DateSerial(2026, 11, 11), "OFFLINE", holdingsFrame, performanceFrame, profitFrame, ordersTable, industryTable, sources
    Set BuildPanel = panel
End Function

Private Function MakeFixtureDataSet(ByVal rawTable As Variant, ByVal datasetName As String, ByVal grain As String, ByVal keyField As String) As clsDataSet
    Dim table As clsDataTable, schema As clsDataSchema, dataset As clsDataSet
    Set table = New clsDataTable
    table.Load rawTable
    Set schema = NewAnySchema(datasetName, grain)
    schema.AddField keyField, "TEXT", False
    schema.Freeze
    Set dataset = New clsDataSet
    dataset.Initialize table, schema, "OFFLINE", datasetName, "FIXTURE-1", "KNOWN"
    Set MakeFixtureDataSet = dataset
End Function

Private Sub Check(ByVal Condition As Boolean, ByRef count As Long)
    If Not Condition Then Err.Raise 5, "Mod_PortfolioTests", "Assertion failed."
    count = count + 1
End Sub
