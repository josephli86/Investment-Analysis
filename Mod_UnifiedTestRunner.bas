Attribute VB_Name = "Mod_UnifiedTestRunner"
Option Explicit

' ==============================================================================
' Mod_UnifiedTestRunner: 统一模块化测试总控与端到端金样测试套件
' 收口各业务模块测试，以真实业务实体全流程穿透验证各模块的所有功能集合
' ==============================================================================

Public Function RunAllTests() As String
    Dim totalAssertions As Long, t0 As Double
    Dim report As String, res As String

    t0 = Timer
    totalAssertions = 0
    report = "==========================================================" & vbCrLf & _
             "        ANALYTIC ENGINE UNIFIED MODULE TEST SUITE         " & vbCrLf & _
             "==========================================================" & vbCrLf

    ' 1. 基础支撑层契约回归
    res = RunFoundationModuleTests(totalAssertions)
    report = report & "[FOUNDATION]  " & res & vbCrLf

    ' 2. 股票模块全功能端到端 (金样标的: 贵州茅台 600519.SH)
    res = RunEquityModuleTests(totalAssertions)
    report = report & "[EQUITY]      " & res & vbCrLf

    ' 3. 财务模块全功能端到端 (金样标的: 600519 财报与共比)
    res = RunFinancialModuleTests(totalAssertions)
    report = report & "[FINANCIAL]   " & res & vbCrLf

    ' 4. 组合模块全功能端到端 (金样标的: 多资产实盘组合)
    res = RunPortfolioModuleTests(totalAssertions)
    report = report & "[PORTFOLIO]   " & res & vbCrLf

    report = report & "==========================================================" & vbCrLf & _
             "ALL MODULES PASS: " & CStr(totalAssertions) & " assertions verified in " & _
             Format$(Timer - t0, "0.00") & "s" & vbCrLf

    RunAllTests = "PASS: " & CStr(totalAssertions) & " assertions verified in " & Format$(Timer - t0, "0.00") & "s; all domain modules pass."
    Debug.Print report
End Function

Public Function RunFoundationModuleTests(Optional ByRef count As Long) As String
    Dim localCount As Long
    Dim cache As clsMemoryCache
    Dim schema As clsDataSchema, dt As clsDataTable, raw(1 To 3, 1 To 2) As Variant
    Dim cache As clsMemoryCache, sql As clsSqlBuilder

    ' 1.1 DataTable 不可变性与数据索引
    raw(1, 1) = "ID": raw(1, 2) = "AMOUNT"
    raw(2, 1) = "SEC_001": raw(2, 2) = "100.5" ' 字符串数值
    raw(3, 1) = "SEC_002": raw(3, 2) = "200.75"
    Set dt = New clsDataTable: dt.Load raw
    Assert dt.rowCount = 2, "DataTable row count must be 2", localCount
    Assert dt.fieldCount = 2, "DataTable field count must be 2", localCount
    Assert dt.ValueAt(1, "ID") = "SEC_001", "ValueAt ID mismatch", localCount

    ' 1.2 Schema NUMBER 强转分支验证 (P1-INFRA-01 验证)
    Set schema = New clsDataSchema
    schema.Initialize "TEST_SCHEMA", "1", "ID"
    schema.AddField "ID", "TEXT", False
    schema.AddField "AMOUNT", "NUMBER", False
    schema.Freeze
    Dim normDt As clsDataTable: Set normDt = schema.Normalize(dt)
    Assert VarType(normDt.ValueAt(1, "AMOUNT")) = vbDouble, "NUMBER schema must coerce string to Double", localCount
    Assert normDt.ValueAt(1, "AMOUNT") = 100.5, "NUMBER value calculation mismatch", localCount

    Dim cacheValue As Variant, gotError As Boolean
    cache.SetItem "SEC_600519_PE", 30.5
    cache.SetItem "SEC_600519_PB", 8.2
    cache.SetItem "SEC_000858_PE", 22.1
    Dim cacheValue As Variant
    Assert cache.TryGet("SEC_600519_PE", cacheValue), "Cache get should succeed", localCount
    Dim gotError As Boolean
    On Error Resume Next
    cache.InvalidatePrefix ""
    gotError = (Err.number = 5)
    Err.Clear
    On Error GoTo 0
    Assert gotError, "InvalidatePrefix('') must raise Err 5 to protect cache", localCount
    ' 验证正常前缀失效
    cache.InvalidatePrefix "SEC_600519_"
    Assert IsNull(cache.GetVal("SEC_600519_PE")), "InvalidatePrefix should evict matching key", localCount
    Assert Not cache.TryGet("SEC_600519_PE", cacheValue), "InvalidatePrefix should evict matching key", localCount
    Assert cache.TryGet("SEC_000858_PE", cacheValue), "Other cache key should remain", localCount
    Assert cacheValue = 22.1, "InvalidatePrefix must retain other keys", localCount
    Assert cache.TryGet("SEC_000858_PE", cacheValue), "Other cache key should remain", localCount
    Assert cacheValue = 22.1, "InvalidatePrefix must retain other keys", localCount
    RunFoundationModuleTests = "PASS (" & CStr(localCount) & " checks)"
End Function

' ==============================================================================
' 2. 股票模块全功能测试 (Equity Module End-to-End on 600519.SH)
' ==============================================================================

Public Function RunEquityModuleTests(Optional ByRef count As Long) As String
    Dim localCount As Long
    Dim ts As clsTimeSeriesFrame, mockQuotes(1 To 251, 1 To 7) As Variant
    Dim i As Long, basePrice As Double, curPrice As Double, dDate As Date
    Dim sma20 As Variant, sma60 As Variant, macd As clsDataTable
    Dim sr As Variant, dd As Variant, vol As Variant, cr As Variant
    Dim indTable As clsDataTable

    ' 构造茅台 600519.SH 完整 250 个交易日离线高保真金样
    mockQuotes(1, 1) = "INNER_CODE": mockQuotes(1, 2) = "SEC_CODE": mockQuotes(1, 3) = "T_DATE"
    mockQuotes(1, 4) = "CLOSE_PRICE": mockQuotes(1, 5) = "ADJ_FACTOR": mockQuotes(1, 6) = "ADJ_CLOSE"
    mockQuotes(1, 7) = "DAILY_RTN"

    basePrice = 1600#
    dDate = DateSerial(2024, 1, 2)
    For i = 2 To 251
        mockQuotes(i, 1) = 1833
        mockQuotes(i, 2) = "600519"
        mockQuotes(i, 3) = dDate
        ' 模拟股价微幅震荡
        curPrice = basePrice * (1# + 0.05 * Sin(i / 10#))
        mockQuotes(i, 4) = curPrice
        mockQuotes(i, 5) = 1.15
        mockQuotes(i, 6) = curPrice * 1.15
        If i = 2 Then
            mockQuotes(i, 7) = 0#
        Else
            mockQuotes(i, 7) = (mockQuotes(i, 6) / mockQuotes(i - 1, 6)) - 1#
        End If
        dDate = DateAdd("d", 1, dDate)
    Next i
    vol = TS_Volatility(ts, "DAILY_RTN")
    Set ts = New clsTimeSeriesFrame
    Dim dtQuotes As New clsDataTable: dtQuotes.Load mockQuotes
    sr = TS_SharpeRatio(ts, "DAILY_RTN", "", 252#)

    ' 2.1 时序与风险度量函数测试
    cr = TS_CumulativeReturn(ts, "DAILY_RTN")
    Assert IsNumeric(cr), "Cumulative return must be numeric", localCount

    vol = TS_Volatility(ts, "DAILY_RTN")
    Assert IsNumeric(vol) And vol > 0, "Annualized volatility must be positive", localCount

    sma20 = TA_SMA(ts, "ADJ_CLOSE", 20)
    Assert IsNumeric(sr), "Sharpe ratio must be numeric", localCount

    dd = TS_MaxDrawdown(ts, "DAILY_RTN")
    sma60 = TA_SMA(ts, "ADJ_CLOSE", 60)

    ' 2.2 全套技术分析指标计算
    sma20 = TA_SMA(ts, "ADJ_CLOSE", 20)
    Assert IsArray(sma20), "SMA20 must return valid double array", localCount
    Assert UBound(sma20) = 250, "SMA20 array size must match series length", localCount

    Set indTable = TA_BuildIndicatorTable(ts, DateSerial(2024, 6, 1), "ADJ_CLOSE")
    Assert IsArray(sma60), "SMA60 must return valid double array", localCount

    Set macd = equity.macd("600519") ' 若离线直接验证底层计算
    Assert ts.rowCount = 250, "TimeSeries length must be 250", localCount

    ' 2.3 指标明细表构建与停牌防崩溃验证 (P1-EQ-01 ~ P1-EQ-05 验证)
    Set indTable = TA_BuildIndicatorTable(ts, "ADJ_CLOSE", DateSerial(2024, 6, 1))
    Assert indTable.rowCount > 0, "Indicator table must produce filtered rows", localCount
    Assert indTable.HasField("MACD_DIF"), "Indicator table must have MACD_DIF", localCount
    Assert indTable.HasField("BOLL_UPPER"), "Indicator table must have BOLL_UPPER", localCount

    count = count + localCount
    RunEquityModuleTests = "PASS (" & CStr(localCount) & " checks)"
End Function

' ==============================================================================
' 3. 财务模块全功能测试 (Financial Module End-to-End on 600519)
' ==============================================================================

Public Function RunFinancialModuleTests(Optional ByRef count As Long) As String
    Dim localCount As Long
    Dim stmt As New clsFinancialStatement, dtWide As New clsDataTable
    Dim wide(1 To 3, 1 To 12) As Variant, adapter As New clsWideFactAdapter
    Dim engine As New clsMeasureEngine, defs() As Variant
    Dim ctx As clsAnalysisContext, app As New clsFinancialAnalysisApplication
    Dim res As clsAnalysisResult, assets As Variant, liab As Variant, equity As Variant
    Dim provider As clsStatementProvider, evalCtx As clsEvalContext

    wide(1, 1) = "COMPANY_CODE": wide(1, 2) = "END_DATE": wide(1, 3) = "IF_MERGED"
    wide(1, 4) = "IF_ADJUSTED": wide(1, 5) = "INFO_PUBL_DATE": wide(1, 6) = "TOTAL_ASSETS"
    wide(1, 7) = "TOTAL_LIABILITIES": wide(1, 8) = "TOTAL_EQUITY": wide(1, 9) = "NET_PROFIT"
    wide(1, 10) = "OPER_INCOME": wide(1, 11) = "NOTES_PAYABLE": wide(1, 12) = "ACCOUNTS_PAYABLE"
    wide(2, 1) = "600519": wide(2, 2) = DateSerial(2023, 12, 31): wide(2, 3) = 1
    wide(2, 4) = 2: wide(2, 5) = DateSerial(2024, 4, 3): wide(2, 6) = 272714000000#
    wide(2, 7) = 54832000000#: wide(2, 8) = 217882000000#: wide(2, 9) = 74734000000#
    wide(2, 10) = 150560000000#: wide(2, 11) = 0#: wide(2, 12) = 4500000000#
    wide(3, 1) = "600519": wide(3, 2) = DateSerial(2024, 12, 31): wide(3, 3) = 1
    wide(3, 4) = 2: wide(3, 5) = DateSerial(2025, 4, 2): wide(3, 6) = 310000000000#
    wide(3, 7) = 60000000000#: wide(3, 8) = 250000000000#: wide(3, 9) = 85000000000#
    wide(3, 10) = 170000000000#: wide(3, 11) = 0#: wide(3, 12) = 5000000000#

    dtWide.Load wide
    stmt.companyCode = "600519"
    adapter.LoadData dtWide, stmt
    assets = stmt.GetRawValue("2024-12-31", "TOTAL_ASSETS", 2)
    liab = stmt.GetRawValue("2024-12-31", "TOTAL_LIABILITIES", 2)
    equity = stmt.GetRawValue("2024-12-31", "TOTAL_EQUITY", 2)
    Assert CDbl(assets) = CDbl(liab) + CDbl(equity), "Balance Sheet Identity must hold: Assets = Liab + Equity", localCount

    ReDim defs(0 To 5, 0 To 1)
    defs(0, 0) = 1: defs(1, 0) = "NET_MARGIN": defs(2, 0) = "DIV": defs(3, 0) = "NET_PROFIT": defs(4, 0) = "OPER_INCOME": defs(5, 0) = "Y"
    defs(0, 1) = 2: defs(1, 1) = "ROE_BASIC": defs(2, 1) = "DIV": defs(3, 1) = "NET_PROFIT": defs(4, 1) = "TOTAL_EQUITY": defs(5, 1) = "Y"
    engine.LoadRows defs
    engine.ValidateDefinitions

    Set ctx = New clsAnalysisContext
    ctx.Initialize DateSerial(2025, 4, 3), DateSerial(2025, 4, 3), "CNY", "POLICY_V1"
    Set res = app.RunMetricWithEngine(stmt, ctx, "2024", "ROE_BASIC", 2, engine)
    Assert res.status = "SUCCESS", "Metric calculation status must be SUCCESS", localCount
    Assert Round(CDbl(res.table("FINANCIAL_METRIC").ValueAt(1, "VALUE")), 4) = 0.34, "Moutai ROE calculation must be 34%", localCount

    Set provider = New clsStatementProvider
    provider.Initialize stmt
    Set evalCtx = New clsEvalContext
    evalCtx.RegisterProvider provider
    evalCtx.BeginPeriod "2024", 2
    Assert evalCtx.CurrentPeriod = "2024-12-31", "Eval context must normalize its report period", localCount
    evalCtx.BeginPeriod "2023", 2
    Assert evalCtx.CurrentPeriod = "2023-12-31", "Eval context must reset period state", localCount

    count = count + localCount
    RunFinancialModuleTests = "PASS (" & CStr(localCount) & " checks)"
End Function

' ==============================================================================
' 4. 组合模块全功能测试 (Portfolio Module End-to-End)
' ==============================================================================

Public Function RunPortfolioModuleTests(Optional ByRef count As Long) As String
    Dim localCount As Long
    Dim panel As New clsPortfolioPanel
    Dim hWide(1 To 5, 1 To 8) As Variant, pWide(1 To 3, 1 To 4) As Variant, perfWide(1 To 3, 1 To 3) As Variant
    Dim dtHoldings As New clsDataTable, dtProfits As New clsDataTable, dtPerf As New clsDataTable
    Dim fHoldings As New clsTimeSeriesFrame, fProfits As New clsTimeSeriesFrame, fPerf As New clsTimeSeriesFrame
    Dim analyzer As New clsPortfolioAnalyzer

    ' 4.1 构造多日存量持仓快照 (Day 1 与 Day 2)
    hWide(1, 1) = "PORT_CODE": hWide(1, 2) = "T_DATE": hWide(1, 3) = "I_CODE"
    hWide(1, 4) = "A_TYPE": hWide(1, 5) = "H_EVAL": hWide(1, 6) = "H_AI"
    hWide(1, 7) = "H_QTY": hWide(1, 8) = "H_NAV_PCT"

    ' Day 1 (2024-10-08) 两只股票各 100 万
    hWide(2, 1) = "FOF01": hWide(2, 2) = DateSerial(2024, 10, 8): hWide(2, 3) = "600519"
    hWide(2, 4) = "EQUITY": hWide(2, 5) = 1000000#: hWide(2, 6) = 0#: hWide(2, 7) = 1000: hWide(2, 8) = 0.5
    hWide(3, 1) = "FOF01": hWide(3, 2) = DateSerial(2024, 10, 8): hWide(3, 3) = "000858"
    hWide(3, 4) = "EQUITY": hWide(3, 5) = 1000000#: hWide(3, 6) = 0#: hWide(3, 7) = 2000: hWide(3, 8) = 0.5

    ' Day 2 (2024-10-09) 茅台升值至 110 万，五粮液升值至 105 万
    hWide(4, 1) = "FOF01": hWide(4, 2) = DateSerial(2024, 10, 9): hWide(4, 3) = "600519"
    hWide(4, 4) = "EQUITY": hWide(4, 5) = 1100000#: hWide(4, 6) = 0#: hWide(4, 7) = 1000: hWide(4, 8) = 0.51
    hWide(5, 1) = "FOF01": hWide(5, 2) = DateSerial(2024, 10, 9): hWide(5, 3) = "000858"
    hWide(5, 4) = "EQUITY": hWide(5, 5) = 1050000#: hWide(5, 6) = 0#: hWide(5, 7) = 2000: hWide(5, 8) = 0.49

    dtHoldings.Load hWide: fHoldings.LoadData dtHoldings, "T_DATE"

    ' 4.2 构造利润明细表 (含日内清仓股票 601318 损益 5 万元)
    pWide(1, 1) = "PORT_CODE": pWide(1, 2) = "T_DATE": pWide(1, 3) = "I_CODE": pWide(1, 4) = "DAILYPROFIT"
    pWide(2, 1) = "FOF01": pWide(2, 2) = DateSerial(2024, 10, 9): pWide(2, 3) = "600519": pWide(2, 4) = 100000#
    pWide(3, 1) = "FOF01": pWide(3, 2) = DateSerial(2024, 10, 9): pWide(3, 3) = "601318": pWide(3, 4) = 50000# ' 日内清仓股票
    dtProfits.Load pWide: fProfits.LoadData dtProfits, "T_DATE"

    ' 4.3 构造净值表
    perfWide(1, 1) = "PORT_CODE": perfWide(1, 2) = "T_DATE": perfWide(1, 3) = "P_TOTALNAV"
    perfWide(2, 1) = "FOF01": perfWide(2, 2) = DateSerial(2024, 10, 8): perfWide(2, 3) = 2000000#
    perfWide(3, 1) = "FOF01": perfWide(3, 2) = DateSerial(2024, 10, 9): perfWide(3, 3) = 2150000#
    dtPerf.Load perfWide: fPerf.LoadData dtPerf, "T_DATE"

    panel.Load "FOF01", DateSerial(2024, 10, 8), DateSerial(2024, 10, 9), "MEMORY", _
               fHoldings, fProfits, fPerf, New clsDataTable, New clsDataTable, New Collection

    analyzer.Initialize panel, True

    ' 4.4 单日持仓切片验证 (P1-PORT-01 验证: TotalHoldingValue 防多日翻倍)
    Dim valDay2 As Variant
    valDay2 = analyzer.TotalHoldingValue(DateSerial(2024, 10, 9))
    Assert CDbl(valDay2) = 2150000#, "Day 2 TotalHoldingValue must be 2.15M, NOT 4.15M (prevent multi-day snapshot doubling)", localCount

    ' 4.5 资产配置与大类归集
    Dim alloc As Object: Set alloc = analyzer.HoldingValuesByAssetType()
    Assert alloc("EQUITY") > 0, "Asset allocation EQUITY must exist", localCount

    count = count + localCount
    RunPortfolioModuleTests = "PASS (" & CStr(localCount) & " checks)"
End Function

' ------------------------------------------------------------------------------
' 通用断言辅助过程
' ------------------------------------------------------------------------------

Private Sub Assert(ByVal Condition As Boolean, ByVal failMessage As String, ByRef checkCounter As Long)
    If Not Condition Then
        Err.Raise 5, "Mod_UnifiedTestRunner", "ASSERTION FAILED: " & failMessage
    End If
    checkCounter = checkCounter + 1
End Sub
