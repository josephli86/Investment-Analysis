VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEquityValuationAnalyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_FinService As clsFinancialService
Private m_Engine As clsMeasureEngine
Private m_IfAdjusted As Long
Private m_MvMultiplier As Double
Private m_Sources As Collection
Private m_SeenSourceKeys As Object

Public Sub Initialize(Optional ByVal finService As clsFinancialService = Nothing, _
        Optional ByVal engine As clsMeasureEngine = Nothing, _
        Optional ByVal ifAdjusted As Long = 2, _
        Optional ByVal mvUnitMultiplier As Double = 10000#)
    If ifAdjusted <> 1 And ifAdjusted <> 2 Then
        Err.Raise 5, "clsEquityValuationAnalyzer.Initialize", "IfAdjusted must be 1 or 2."
    End If
    If mvUnitMultiplier <= 0# Then
        Err.Raise 5, "clsEquityValuationAnalyzer.Initialize", "mvUnitMultiplier must be positive."
    End If

    If finService Is Nothing Then Set m_FinService = New clsFinancialService Else Set m_FinService = finService
    If engine Is Nothing Then Set m_Engine = MeasureMgr() Else Set m_Engine = engine
    m_IfAdjusted = ifAdjusted
    m_MvMultiplier = mvUnitMultiplier
    Set m_Sources = New Collection
    Set m_SeenSourceKeys = CreateObject("Scripting.Dictionary")
    m_SeenSourceKeys.CompareMode = 1
End Sub

Public Property Get CollectedFinancialSources() As Collection
    Dim copy As New Collection, item As Variant
    For Each item In m_Sources
        copy.Add item
    Next item
    Set CollectedFinancialSources = copy
End Property

' Build daily valuation ratio series using Last Fiscal Year (LFY) annual report (e.g. ADJ_PE = TOTAL_MV / NOPAT_LFY)
Public Function BuildDailyLfyRatioSeries(ByVal panel As clsEquityPanel, _
        ByVal secCodeOrInnerCode As Variant, _
        Optional ByVal metricCode As String = "NOPAT", _
        Optional ByVal ratioName As String = "ADJ_PE", _
        Optional ByVal useFloatMv As Boolean = False) As clsDataTable
    Dim innerCode As Long, companyCode As String, secCode As String
    Dim quoteTs As clsTimeSeriesFrame, n As Long, r As Long
    Dim tDate As Date, mvField As String, mvRaw As Variant, mvScaled As Variant
    Dim cleanMetric As String, cleanRatio As String
    Dim masterStmt As clsFinancialStatement, ds As clsDataSet, srcKey As String
    Dim beginPeriodText As String, endAsOfText As String
    Dim regimeCache As Object, finRes As Variant, ratioVal As Variant, statusText As String
    Dim rows() As Variant, resTable As clsDataTable

    If panel Is Nothing Or Not panel.IsLoaded Then
        Err.Raise 5, "clsEquityValuationAnalyzer.BuildDailyLfyRatioSeries", "Loaded equity panel is required."
    End If

    innerCode = panel.ResolveInnerCode(secCodeOrInnerCode)
    ResolveMasterMeta panel.MasterTable, innerCode, secCode, companyCode
    If Len(Trim$(companyCode)) = 0 Then
        Err.Raise 5, "clsEquityValuationAnalyzer.BuildDailyLfyRatioSeries", "Missing COMPANY_CODE for security: " & CStr(secCodeOrInnerCode)
    End If

    Set quoteTs = panel.QuoteFrameForStock(innerCode, includeWarmup:=False)
    n = quoteTs.rowCount
    mvField = IIf(useFloatMv, "FLOAT_MV", "TOTAL_MV")
    Call quoteTs.FieldIndex(mvField)

    cleanMetric = UCase$(Trim$(metricCode))
    cleanRatio = UCase$(Trim$(ratioName))

    ' Load financial reports once through panel.endDate; each trading day enforces strict PiT via CloneView + ConfigureView
    beginPeriodText = CStr(Year(panel.startDate) - 5) & "-12-31"
    endAsOfText = Format$(panel.endDate, "yyyy-mm-dd")
    Set masterStmt = m_FinService.GetStatement(companyCode, beginPeriodText, endAsOfText, 1, endAsOfText, "", m_IfAdjusted)
    For Each ds In masterStmt.SourceDataSets
        srcKey = ds.schema.datasetName & "|" & ds.sourceAlias & "|" & ds.queryId & "|" & ds.revisionId & "|" & CStr(ds.LoadedAt)
        If Not m_SeenSourceKeys.exists(srcKey) Then
            m_SeenSourceKeys.Add srcKey, True
            m_Sources.Add ds
        End If
    Next ds

    Set regimeCache = CreateObject("Scripting.Dictionary")
    regimeCache.CompareMode = 1

    ReDim rows(1 To n + 1, 1 To 10)
    rows(1, 1) = "SEC_CODE": rows(1, 2) = "COMPANY_CODE": rows(1, 3) = "T_DATE"
    rows(1, 4) = "ADJ_CLOSE": rows(1, 5) = "MV_CNY": rows(1, 6) = "FISCAL_YEAR_END"
    rows(1, 7) = "INFO_PUBL_DATE": rows(1, 8) = cleanMetric & "_LFY"
    rows(1, 9) = cleanRatio: rows(1, 10) = "STATUS"

    For r = 1 To n
        tDate = dateValue(CDate(quoteTs.RowDate(r)))
        mvRaw = quoteTs.ValueAt(r, mvField)

        ' finRes = Array(metricValue, fiscalYearEnd, publDate, statusText)
        finRes = ResolveAnnualRegimeAtDate(masterStmt, tDate, cleanMetric, regimeCache)

        rows(r + 1, 1) = secCode
        rows(r + 1, 2) = companyCode
        rows(r + 1, 3) = tDate
        rows(r + 1, 4) = quoteTs.ValueAt(r, "ADJ_CLOSE")
        rows(r + 1, 6) = finRes(1)
        rows(r + 1, 7) = finRes(2)
        rows(r + 1, 8) = finRes(0)

        If IsNull(mvRaw) Or Not IsNumeric(mvRaw) Then
            mvScaled = Null
            ratioVal = Null
            statusText = "MISSING_MV"
        Else
            mvScaled = CDbl(mvRaw) * m_MvMultiplier
            If IsNull(finRes(0)) Then
                ratioVal = Null
                statusText = CStr(finRes(3))
            ElseIf CDbl(finRes(0)) <= 0# Then
                ratioVal = Null
                statusText = "NON_POSITIVE_DENOMINATOR"
            Else
                ratioVal = mvScaled / CDbl(finRes(0))
                statusText = "OK"
            End If
        End If

        rows(r + 1, 5) = mvScaled
        rows(r + 1, 9) = ratioVal
        rows(r + 1, 10) = statusText
    Next r

    Set resTable = New clsDataTable
    resTable.Load rows
    Set BuildDailyLfyRatioSeries = resTable
End Function

' Summarize cross-sectional valuation and historical percentile ranks across multiple securities
Public Function BuildValuationSummary(ByVal dailySeriesList As Collection, _
        Optional ByVal ratioName As String = "ADJ_PE") As clsDataTable
    Dim seriesTbl As clsDataTable, colName As String
    Dim r As Long, validVals() As Double, vCount As Long
    Dim latestVal As Variant, latestDate As Variant, latestMetric As Variant, latestMv As Variant, latestPeriod As Variant
    Dim sIdx As Long, outRows() As Variant, resTable As clsDataTable

    colName = UCase$(Trim$(ratioName))
    ReDim outRows(1 To dailySeriesList.count + 1, 1 To 12)
    outRows(1, 1) = "SEC_CODE": outRows(1, 2) = "COMPANY_CODE": outRows(1, 3) = "LATEST_DATE"
    outRows(1, 4) = "FISCAL_YEAR_END": outRows(1, 5) = "MV_BN_CNY": outRows(1, 6) = "DENOM_BN_CNY"
    outRows(1, 7) = "LATEST_" & colName: outRows(1, 8) = "HIST_MIN": outRows(1, 9) = "HIST_MEDIAN"
    outRows(1, 10) = "HIST_MAX": outRows(1, 11) = "PERCENTILE_RANK": outRows(1, 12) = "VALID_DAYS"

    For sIdx = 1 To dailySeriesList.count
        Set seriesTbl = dailySeriesList(sIdx)
        vCount = 0
        latestVal = Null: latestDate = Null: latestMetric = Null: latestMv = Null: latestPeriod = Null
        If seriesTbl.rowCount > 0 Then
            ReDim validVals(1 To seriesTbl.rowCount)
            For r = 1 To seriesTbl.rowCount
                If Not IsNull(seriesTbl.ValueAt(r, colName)) Then
                    vCount = vCount + 1
                    validVals(vCount) = CDbl(seriesTbl.ValueAt(r, colName))
                End If
            Next r
            r = seriesTbl.rowCount
            latestDate = seriesTbl.ValueAt(r, "T_DATE")
            latestPeriod = seriesTbl.ValueAt(r, "FISCAL_YEAR_END")
            latestMv = seriesTbl.ValueAt(r, "MV_CNY")
            latestMetric = seriesTbl.ValueAtColumn(r, 8)
            latestVal = seriesTbl.ValueAt(r, colName)
        End If

        If seriesTbl.rowCount > 0 Then
            outRows(sIdx + 1, 1) = seriesTbl.ValueAt(1, "SEC_CODE")
            outRows(sIdx + 1, 2) = seriesTbl.ValueAt(1, "COMPANY_CODE")
        Else
            outRows(sIdx + 1, 1) = ""
            outRows(sIdx + 1, 2) = ""
        End If
        outRows(sIdx + 1, 3) = latestDate
        outRows(sIdx + 1, 4) = latestPeriod

        If IsNull(latestMv) Or Not IsNumeric(latestMv) Then
            outRows(sIdx + 1, 5) = Null
        Else
            outRows(sIdx + 1, 5) = CDbl(latestMv) / 100000000#
        End If

        If IsNull(latestMetric) Or Not IsNumeric(latestMetric) Then
            outRows(sIdx + 1, 6) = Null
        Else
            outRows(sIdx + 1, 6) = CDbl(latestMetric) / 100000000#
        End If
        outRows(sIdx + 1, 7) = latestVal
        outRows(sIdx + 1, 12) = vCount

        If vCount > 0 Then
            ReDim Preserve validVals(1 To vCount)
            SortDoubleArray validVals
            outRows(sIdx + 1, 8) = validVals(1)
            outRows(sIdx + 1, 9) = MedianSorted(validVals)
            outRows(sIdx + 1, 10) = validVals(vCount)
            If Not IsNull(latestVal) Then
                outRows(sIdx + 1, 11) = PercentileRankSorted(validVals, CDbl(latestVal))
            Else
                outRows(sIdx + 1, 11) = Null
            End If
        Else
            outRows(sIdx + 1, 8) = Null: outRows(sIdx + 1, 9) = Null
            outRows(sIdx + 1, 10) = Null: outRows(sIdx + 1, 11) = Null
        End If
    Next sIdx

    Set resTable = New clsDataTable
    resTable.Load outRows
    Set BuildValuationSummary = resTable
End Function

' ==============================================================================
' Point-in-Time Annual Report Resolver (12-31 reports, cached by report batch)
' ==============================================================================

Private Function ResolveAnnualRegimeAtDate(ByVal masterStmt As clsFinancialStatement, _
        ByVal tDate As Date, _
        ByVal metricCode As String, _
        ByVal regimeCache As Object) As Variant
    Dim viewStmt As clsFinancialStatement, asOfText As String
    Dim fyEnd As String, vKey As String, pubDateText As String
    Dim cacheKey As String, valMetric As Variant

    asOfText = Format$(tDate, "yyyy-mm-dd")
    Set viewStmt = masterStmt.CloneView()
    viewStmt.ConfigureView 1, asOfText

    fyEnd = FindLatestAnnualPeriod(viewStmt, tDate, 5, vKey, pubDateText)
    If Len(vKey) = 0 Then
        ResolveAnnualRegimeAtDate = Array(Null, "", "", "MISSING_ANNUAL_REPORT")
        Exit Function
    End If

    cacheKey = fyEnd & "|" & pubDateText & "|" & metricCode
    If regimeCache.exists(cacheKey) Then
        ResolveAnnualRegimeAtDate = regimeCache(cacheKey)
        Exit Function
    End If

    valMetric = EvalPinnedMetric(viewStmt, fyEnd, m_IfAdjusted, pubDateText, metricCode)
    If IsNull(valMetric) Then
        regimeCache(cacheKey) = Array(Null, fyEnd, pubDateText, "METRIC_EVAL_ERROR")
    Else
        regimeCache(cacheKey) = Array(CDbl(valMetric), fyEnd, pubDateText, "OK")
    End If

    ResolveAnnualRegimeAtDate = regimeCache(cacheKey)
End Function

Private Function FindLatestAnnualPeriod(ByVal stmt As clsFinancialStatement, _
        ByVal asOfDate As Date, _
        ByVal lookbackYears As Long, _
        ByRef outVersionKey As String, _
        ByRef outPublDate As String) As String
    Dim startYear As Long, i As Long, candidatePeriod As String, vKey As String
    startYear = Year(asOfDate)
    For i = 0 To lookbackYears - 1
        candidatePeriod = CStr(startYear - i) & "-12-31"
        vKey = stmt.SelectedVersionKey(candidatePeriod, m_IfAdjusted)
        If Len(vKey) > 0 Then
            outVersionKey = vKey
            outPublDate = stmt.SelectedPublicationDate(candidatePeriod, m_IfAdjusted)
            FindLatestAnnualPeriod = candidatePeriod
            Exit Function
        End If
    Next i
    outVersionKey = ""
    outPublDate = ""
    FindLatestAnnualPeriod = ""
End Function

Private Function EvalPinnedMetric(ByVal stmt As clsFinancialStatement, _
        ByVal periodEnd As String, _
        ByVal ifAdj As Long, _
        ByVal publDate As String, _
        ByVal metricCode As String) As Variant
    Dim provider As New clsStatementProvider, ctx As New clsEvalContext, v As Variant
    provider.Initialize stmt
    provider.PinReportVersion periodEnd, ifAdj, publDate
    ctx.RegisterProvider provider
    v = m_Engine.EvaluateContext(ctx, periodEnd, metricCode, CInt(ifAdj))
    If IsError(v) Then EvalPinnedMetric = Null Else EvalPinnedMetric = CDbl(v)
End Function

Private Sub ResolveMasterMeta(ByVal masterTbl As clsDataTable, ByVal innerCode As Long, _
        ByRef outSecCode As String, ByRef outCompanyCode As String)
    Dim r As Long, rawComp As Variant
    For r = 1 To masterTbl.rowCount
        If CLng(masterTbl.ValueAt(r, "INNER_CODE")) = innerCode Then
            outSecCode = CStr(masterTbl.ValueAt(r, "SEC_CODE"))
            rawComp = masterTbl.ValueAt(r, "COMPANY_CODE")
            If IsNull(rawComp) Or IsEmpty(rawComp) Then
                outCompanyCode = ""
            Else
                outCompanyCode = Trim$(CStr(rawComp))
            End If
            Exit Sub
        End If
    Next r
End Sub

Private Sub SortDoubleArray(ByRef arr() As Double)
    Dim i As Long, j As Long, tmp As Double
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
        Next j
    Next i
End Sub

Private Function MedianSorted(ByRef sorted() As Double) As Double
    Dim n As Long
    n = UBound(sorted) - LBound(sorted) + 1
    If n Mod 2 = 1 Then
        MedianSorted = sorted(LBound(sorted) + n \ 2)
    Else
        MedianSorted = (sorted(LBound(sorted) + n \ 2 - 1) + sorted(LBound(sorted) + n \ 2)) / 2#
    End If
End Function

Private Function PercentileRankSorted(ByRef sorted() As Double, ByVal target As Double) As Double
    Dim i As Long, n As Long, leCount As Long
    n = UBound(sorted) - LBound(sorted) + 1
    For i = LBound(sorted) To UBound(sorted)
        If sorted(i) <= target Then leCount = leCount + 1
    Next i
    PercentileRankSorted = leCount / CDbl(n)
End Function
