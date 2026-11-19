VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPortfolioAnalyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Panel As clsPortfolioPanel
Private m_AmountsConfirmed As Boolean

Public Sub Initialize(ByVal panel As clsPortfolioPanel, Optional ByVal amountsComparableConfirmed As Boolean = False)
    If panel Is Nothing Then Err.Raise 5, , "Portfolio panel is required."
    Set m_Panel = panel
    m_AmountsConfirmed = amountsComparableConfirmed
End Sub

Public Property Get HoldingWeight(ByVal rowNumber As Long) As Variant
    Dim Value As Variant
    HoldingWeight = Null
    If Not m_AmountsConfirmed Then Exit Property
    Value = HoldingValue(rowNumber, "H_NAV_PCT")
    If IsNull(Value) Then Exit Property
    If Not IsNumeric(Value) Then Err.Raise 13, , "Non-numeric holding weight."
    HoldingWeight = CDbl(Value)
End Property

Public Function TotalHoldingValue(Optional ByVal asOfDate As Variant) As Variant
    Dim r As Long, Value As Variant, total As Double, targetDate As Date, hasDateFilter As Boolean
    TotalHoldingValue = Null
    If Not m_AmountsConfirmed Or m_Panel.HoldingRowCount = 0 Then Exit Function
    If Not IsMissing(asOfDate) And Not IsEmpty(asOfDate) Then
        targetDate = CDate(asOfDate)
        hasDateFilter = True
    End If
    For r = 1 To m_Panel.HoldingRowCount
        If hasDateFilter Then
            If dateValue(CDate(HoldingValue(r, "T_DATE"))) <> dateValue(targetDate) Then GoTo NextRow
        End If
        Value = HoldingValue(r, "H_EVAL")
        If IsNull(Value) Then Exit Function
        If Not IsNumeric(Value) Then Err.Raise 13, , "Non-numeric holding value."
        total = total + CDbl(Value)
NextRow:
    Next r
    TotalHoldingValue = total
End Function

Public Function HoldingValuesByAssetType() As Object
    Dim result As Object, r As Long, Value As Variant, assetType As String
    If Not m_AmountsConfirmed Then Err.Raise 5, , "Confirm amount currency and units before aggregation."
    Set result = CreateObject("Scripting.Dictionary")
    For r = 1 To m_Panel.HoldingRowCount
        assetType = CStr(HoldingValue(r, "A_TYPE"))
        Value = HoldingValue(r, "H_EVAL")
        If Not result.exists(assetType) Then result.Add assetType, 0#
        If IsNull(Value) Then
            result(assetType) = Null
        ElseIf Not IsNumeric(Value) Then
            Err.Raise 13, , "Non-numeric holding value."
        ElseIf Not IsNull(result(assetType)) Then
            result(assetType) = CDbl(result(assetType)) + CDbl(Value)
        End If
    Next r
    Set HoldingValuesByAssetType = result
End Function

Private Function HoldingValue(ByVal rowNumber As Long, ByVal fieldName As String) As Variant
    HoldingValue = m_Panel.holdingsFrame.ValueAt(rowNumber, fieldName)
End Function

Public Function IndustryAnalysis() As Variant
    Dim holdingsFrame As clsTimeSeriesFrame, profitFrame As clsTimeSeriesFrame, performanceFrame As clsTimeSeriesFrame
    Dim industryIndex As Object, profitIndex As Object, navIndex As Object, groups As Object, seenProfit As Object
    Dim r As Long, outputRow As Long, holdDate As Date, DateKey As String
    Dim portCode As String, secCode As String, sectorCode As String, sectorName As String
    Dim amount As Variant, nav As Variant, matchCount As Long, status As String
    Dim key As String, profitKey As String, navKey As String, item As Variant, values As Variant, output() As Variant
    Dim industryRows As Collection, industryItem As Variant, bucket As Collection
    Dim unclassifiedLabel As String

    unclassifiedLabel = ChrW(&H672A) & ChrW(&H5206) & ChrW(&H7C7B)
    Set holdingsFrame = m_Panel.holdingsFrame
    Set profitFrame = m_Panel.profitFrame
    Set performanceFrame = m_Panel.performanceFrame
    ValidateFrameFields holdingsFrame, Array("PORT_CODE", "T_DATE", "I_CODE", "H_EVAL", "H_AI"), "portfolio"
    ValidateFrameFields profitFrame, Array("PORT_CODE", "I_CODE", "T_DATE", "DAILYPROFIT"), "profit"
    ValidateFrameFields performanceFrame, Array("PORT_CODE", "T_DATE", "P_TOTALNAV"), "performance"
    ValidateDataTableFields m_Panel.industryData, Array("SEC_CODE", "START_DATE", "EXPIRATION_DATE", "SW_SECTOR_CODE", "SW_SECTOR"), "industry"
    Set industryIndex = CreateObject("Scripting.Dictionary"): Set profitIndex = CreateObject("Scripting.Dictionary")
    Set navIndex = CreateObject("Scripting.Dictionary"): Set groups = CreateObject("Scripting.Dictionary")
    Set seenProfit = CreateObject("Scripting.Dictionary")
    industryIndex.CompareMode = 1: profitIndex.CompareMode = 1: navIndex.CompareMode = 1: seenProfit.CompareMode = 1

    For r = 1 To m_Panel.industryData.rowCount
        If Not IsNull(m_Panel.industryData.ValueAt(r, "SEC_CODE")) Then
            secCode = UCase$(Trim$(CStr(m_Panel.industryData.ValueAt(r, "SEC_CODE"))))
            If Len(secCode) > 0 And IsDate(m_Panel.industryData.ValueAt(r, "START_DATE")) And IsDate(m_Panel.industryData.ValueAt(r, "EXPIRATION_DATE")) Then
                If Not industryIndex.exists(secCode) Then Set bucket = New Collection: industryIndex.Add secCode, bucket
                Set bucket = industryIndex(secCode)
                bucket.Add Array(CDbl(dateValue(CDate(m_Panel.industryData.ValueAt(r, "START_DATE")))), CDbl(dateValue(CDate(m_Panel.industryData.ValueAt(r, "EXPIRATION_DATE")))), m_Panel.industryData.ValueAt(r, "SW_SECTOR_CODE"), m_Panel.industryData.ValueAt(r, "SW_SECTOR"))
            End If
        End If
    Next r

    For r = 1 To profitFrame.rowCount
        If Not IsNull(profitFrame.ValueAt(r, "PORT_CODE")) And Not IsNull(profitFrame.ValueAt(r, "I_CODE")) And IsDate(profitFrame.ValueAt(r, "T_DATE")) Then
            portCode = UCase$(Trim$(CStr(profitFrame.ValueAt(r, "PORT_CODE"))))
            secCode = UCase$(Trim$(CStr(profitFrame.ValueAt(r, "I_CODE"))))
            profitKey = portCode & "|" & Format$(dateValue(CDate(profitFrame.ValueAt(r, "T_DATE"))), "yyyymmdd") & "|" & secCode
            If Not profitIndex.exists(profitKey) Then profitIndex.Add profitKey, 0#
            If Not IsNull(profitFrame.ValueAt(r, "DAILYPROFIT")) Then
                If IsNumeric(profitFrame.ValueAt(r, "DAILYPROFIT")) Then profitIndex(profitKey) = CDbl(profitIndex(profitKey)) + CDbl(profitFrame.ValueAt(r, "DAILYPROFIT"))
            End If
        End If
    Next r

    For r = 1 To performanceFrame.rowCount
        If Not IsNull(performanceFrame.ValueAt(r, "PORT_CODE")) And IsDate(performanceFrame.ValueAt(r, "T_DATE")) Then
            portCode = UCase$(Trim$(CStr(performanceFrame.ValueAt(r, "PORT_CODE"))))
            navKey = portCode & "|" & Format$(dateValue(CDate(performanceFrame.ValueAt(r, "T_DATE"))), "yyyymmdd")
            If Not navIndex.exists(navKey) Then navIndex.Add navKey, performanceFrame.ValueAt(r, "P_TOTALNAV")
        End If
    Next r

    For r = 1 To holdingsFrame.rowCount
        If Not IsDate(holdingsFrame.ValueAt(r, "T_DATE")) Or IsNull(holdingsFrame.ValueAt(r, "PORT_CODE")) Then GoTo NextHolding
        matchCount = 0
        status = "OK"
        sectorCode = ""
        sectorName = ""

        portCode = UCase$(Trim$(CStr(holdingsFrame.ValueAt(r, "PORT_CODE"))))
        holdDate = dateValue(CDate(holdingsFrame.ValueAt(r, "T_DATE"))): DateKey = Format$(holdDate, "yyyymmdd")
        If IsNull(holdingsFrame.ValueAt(r, "I_CODE")) Then secCode = "" Else secCode = UCase$(Trim$(CStr(holdingsFrame.ValueAt(r, "I_CODE"))))

        If IsNull(holdingsFrame.ValueAt(r, "H_EVAL")) Then
            amount = Null
        ElseIf Not IsNumeric(holdingsFrame.ValueAt(r, "H_EVAL")) Then
            amount = holdingsFrame.ValueAt(r, "H_EVAL")
        ElseIf IsNull(holdingsFrame.ValueAt(r, "H_AI")) Then
            amount = CDbl(holdingsFrame.ValueAt(r, "H_EVAL"))
        ElseIf IsNumeric(holdingsFrame.ValueAt(r, "H_AI")) Then
            amount = CDbl(holdingsFrame.ValueAt(r, "H_EVAL")) + CDbl(holdingsFrame.ValueAt(r, "H_AI"))
        Else
            amount = holdingsFrame.ValueAt(r, "H_EVAL")
        End If

        If industryIndex.exists(secCode) Then
            Set industryRows = industryIndex(secCode)
            For Each industryItem In industryRows
                If CDbl(holdDate) >= CDbl(industryItem(0)) And CDbl(holdDate) <= CDbl(industryItem(1)) Then
                    matchCount = matchCount + 1
                    If IsNull(industryItem(2)) Then sectorCode = "" Else sectorCode = CStr(industryItem(2))
                    If IsNull(industryItem(3)) Then sectorName = "" Else sectorName = CStr(industryItem(3))
                End If
            Next industryItem
        End If

        If matchCount = 0 Then
            sectorCode = "UNCLASSIFIED": sectorName = unclassifiedLabel: status = "NO_INDUSTRY_MATCH"
        ElseIf matchCount > 1 Then
            status = "MULTIPLE_INDUSTRY_MATCH"
        End If

        navKey = portCode & "|" & DateKey
        If navIndex.exists(navKey) Then nav = navIndex(navKey) Else nav = Null
        key = portCode & "|" & DateKey & "|" & sectorCode & "|" & sectorName
        If Not groups.exists(key) Then groups.Add key, Array(portCode, holdDate, sectorCode, sectorName, 0#, 0#, nav, status)
        values = groups(key)

        If IsNull(amount) Then
            values(7) = "MISSING_AMOUNT"
        ElseIf IsNumeric(amount) Then
            values(4) = CDbl(values(4)) + CDbl(amount)
        Else
            values(7) = "NON_NUMERIC_AMOUNT"
        End If

        profitKey = portCode & "|" & DateKey & "|" & secCode
        If profitIndex.exists(profitKey) And Not seenProfit.exists(profitKey) Then
            seenProfit.Add profitKey, True
            values(5) = CDbl(values(5)) + CDbl(profitIndex(profitKey))
        End If
        If status <> "OK" Then values(7) = status
        groups(key) = values
NextHolding:
    Next r

    ReDim output(1 To groups.count + 1, 1 To 11)
    output(1, 1) = "PORT_CODE": output(1, 2) = "T_DATE": output(1, 3) = "SW_SECTOR_CODE": output(1, 4) = "SW_SECTOR"
    output(1, 5) = "HOLDING_VALUE": output(1, 6) = "HOLDING_PCT": output(1, 7) = "PROFIT": output(1, 8) = "RETURN_RATE"
    output(1, 9) = "RETURN_CONTRIBUTION": output(1, 10) = "P_TOTALNAV": output(1, 11) = "STATUS"
    outputRow = 1
    For Each item In groups.Keys
        outputRow = outputRow + 1: values = groups(item)
        output(outputRow, 1) = values(0): output(outputRow, 2) = values(1): output(outputRow, 3) = values(2)
        output(outputRow, 4) = values(3): output(outputRow, 5) = values(4): output(outputRow, 7) = values(5)
        output(outputRow, 10) = values(6)
        If Not IsNull(values(6)) Then
            If IsNumeric(values(6)) Then
                If CDbl(values(6)) <> 0 Then
                    output(outputRow, 6) = CDbl(values(4)) / CDbl(values(6))
                    output(outputRow, 9) = CDbl(values(5)) / CDbl(values(6))
                End If
            End If
        End If
        If CDbl(values(4)) <> 0 Then output(outputRow, 8) = CDbl(values(5)) / CDbl(values(4))
        output(outputRow, 11) = values(7)
    Next item
    IndustryAnalysis = output
End Function

Public Function AbsoluteAttribution() As Variant
    Dim perfByPort As Object, profitBuckets As Object, attribution As Object, industryBySecurity As Object
    Dim r As Long, i As Long, n As Long, outputRow As Long
    Dim port As String, sec As String, dayKey As String, bucketKey As String, key As String, level As String
    Dim d As Date, priorNav As Double, contribution As Double, factor As Double, profit As Double
    Dim item As Variant, values As Variant, rowValues As Variant
    Dim industryItem As Variant, portfolioKey As Variant, output() As Variant
    Dim rows As Collection, dayProfitRows As Collection, industryRows As Collection
    Dim performanceFrame As clsTimeSeriesFrame, profitFrame As clsTimeSeriesFrame, portfolioFrame As clsTimeSeriesFrame
    Dim sector As String, sectorName As String, status As String

    Set performanceFrame = m_Panel.performanceFrame
    Set profitFrame = m_Panel.profitFrame
    ValidateFrameFields performanceFrame, Array("T_DATE", "PORT_CODE", "RTN", "P_TOTALNAV"), "performance"
    ValidateFrameFields profitFrame, Array("T_DATE", "PORT_CODE", "I_CODE", "DAILYPROFIT"), "profit"
    ValidateDataTableFields m_Panel.industryData, Array("SEC_CODE", "START_DATE", "EXPIRATION_DATE", "SW_SECTOR_CODE", "SW_SECTOR"), "industry"

    Set perfByPort = CreateObject("Scripting.Dictionary")
    Set profitBuckets = CreateObject("Scripting.Dictionary")
    Set attribution = CreateObject("Scripting.Dictionary")
    Set industryBySecurity = CreateObject("Scripting.Dictionary")
    perfByPort.CompareMode = 1: profitBuckets.CompareMode = 1: attribution.CompareMode = 1: industryBySecurity.CompareMode = 1

    ' 1. Bucket performance rows (including prior-day buffer before m_Panel.startDate)
    For r = 1 To performanceFrame.rowCount
        If IsDate(performanceFrame.ValueAt(r, "T_DATE")) And Not IsNull(performanceFrame.ValueAt(r, "PORT_CODE")) Then
            port = UCase$(Trim$(CStr(performanceFrame.ValueAt(r, "PORT_CODE"))))
            d = dateValue(CDate(performanceFrame.ValueAt(r, "T_DATE")))
            If d <= m_Panel.endDate Then
                If Not perfByPort.exists(port) Then
                    Set rows = New Collection
                    perfByPort.Add port, rows
                End If
                perfByPort(port).Add r
            End If
        End If
    Next r

    ' 2. Bucket daily profits by PORT_CODE|yyyymmdd for O(1) lookup
    For r = 1 To profitFrame.rowCount
        If Not IsNull(profitFrame.ValueAt(r, "PORT_CODE")) And IsDate(profitFrame.ValueAt(r, "T_DATE")) And Not IsNull(profitFrame.ValueAt(r, "I_CODE")) Then
            port = UCase$(Trim$(CStr(profitFrame.ValueAt(r, "PORT_CODE"))))
            d = dateValue(CDate(profitFrame.ValueAt(r, "T_DATE")))
            If d >= m_Panel.startDate And d <= m_Panel.endDate Then
                bucketKey = port & "|" & Format$(d, "yyyymmdd")
                If Not profitBuckets.exists(bucketKey) Then
                    Set dayProfitRows = New Collection
                    profitBuckets.Add bucketKey, dayProfitRows
                End If
                If Not IsNumeric(profitFrame.ValueAt(r, "DAILYPROFIT")) Then Err.Raise 13, , "Missing/non-numeric DAILYPROFIT: " & port & " " & Format$(d, "yyyymmdd")
                profitBuckets(bucketKey).Add Array(UCase$(Trim$(CStr(profitFrame.ValueAt(r, "I_CODE")))), CDbl(profitFrame.ValueAt(r, "DAILYPROFIT")))
            End If
        End If
    Next r

    ' 3. Index industry intervals
    For r = 1 To m_Panel.industryData.rowCount
        If Not IsNull(m_Panel.industryData.ValueAt(r, "SEC_CODE")) And IsDate(m_Panel.industryData.ValueAt(r, "START_DATE")) And IsDate(m_Panel.industryData.ValueAt(r, "EXPIRATION_DATE")) Then
            sec = UCase$(Trim$(CStr(m_Panel.industryData.ValueAt(r, "SEC_CODE"))))
            If Not industryBySecurity.exists(sec) Then
                Set industryRows = New Collection
                industryBySecurity.Add sec, industryRows
            End If
            industryBySecurity(sec).Add Array(CDbl(dateValue(CDate(m_Panel.industryData.ValueAt(r, "START_DATE")))), CDbl(dateValue(CDate(m_Panel.industryData.ValueAt(r, "EXPIRATION_DATE")))), m_Panel.industryData.ValueAt(r, "SW_SECTOR_CODE"), m_Panel.industryData.ValueAt(r, "SW_SECTOR"))
        End If
    Next r

    ' 4. Backward Frongello linking in O(N)
    For Each portfolioKey In perfByPort.Keys
        port = CStr(portfolioKey)
        Set rows = perfByPort(port)
        Set portfolioFrame = performanceFrame.SelectRows(rows)
        Set portfolioFrame = portfolioFrame.SortedByDate()
        portfolioFrame.ValidateSingleSeries
        n = portfolioFrame.rowCount
        factor = 1#

        For i = n To 1 Step -1
            d = dateValue(CDate(portfolioFrame.RowDate(i)))
            If d >= m_Panel.startDate Then
                If Not IsNumeric(portfolioFrame.ValueAt(i, "RTN")) Then Err.Raise 13, , "Missing/non-numeric RTN: " & port & " " & Format$(d, "yyyy-mm-dd")
                dayKey = Format$(d, "yyyymmdd")
                bucketKey = port & "|" & dayKey

                If profitBuckets.exists(bucketKey) Then
                    If i = 1 Then Err.Raise 5, , "Missing positive prior NAV for portfolio " & port & " on " & dayKey
                    If Not IsNumeric(portfolioFrame.ValueAt(i - 1, "P_TOTALNAV")) Then Err.Raise 13, , "Missing/non-numeric prior NAV for portfolio " & port
                    priorNav = CDbl(portfolioFrame.ValueAt(i - 1, "P_TOTALNAV"))
                    If priorNav <= 0# Then Err.Raise 5, , "Missing positive prior NAV for portfolio " & port & " on " & dayKey

                    Set dayProfitRows = profitBuckets(bucketKey)
                    For Each item In dayProfitRows
                        sec = CStr(item(0))
                        profit = CDbl(item(1))
                        sector = "UNCLASSIFIED"
                        sectorName = "UNCLASSIFIED"
                        status = "NO_INDUSTRY_MATCH"
                        If industryBySecurity.exists(sec) Then
                            Set industryRows = industryBySecurity(sec)
                            For Each industryItem In industryRows
                                If CDbl(d) >= CDbl(industryItem(0)) And CDbl(d) <= CDbl(industryItem(1)) Then
                                    If IsNull(industryItem(2)) Then sector = "UNCLASSIFIED" Else sector = CStr(industryItem(2))
                                    If IsNull(industryItem(3)) Then sectorName = "UNCLASSIFIED" Else sectorName = CStr(industryItem(3))
                                    status = "OK"
                                    Exit For
                                End If
                            Next industryItem
                        End If
                        contribution = profit / priorNav * factor
                        For Each values In Array("SECURITY|" & sec & "|" & sec, "INDUSTRY|" & sector & "|" & sectorName)
                            rowValues = Split(CStr(values), "|")
                            level = CStr(rowValues(0))
                            key = port & "|" & level & "|" & CStr(rowValues(1)) & "|" & CStr(rowValues(2))
                            If Not attribution.exists(key) Then attribution.Add key, Array(port, level, CStr(rowValues(1)), CStr(rowValues(2)), 0#, status)
                            rowValues = attribution(key)
                            rowValues(4) = CDbl(rowValues(4)) + contribution
                            attribution(key) = rowValues
                        Next values
                    Next item
                End If
                factor = factor * (1# + CDbl(portfolioFrame.ValueAt(i, "RTN")))
            End If
        Next i
    Next portfolioKey

    ReDim output(1 To attribution.count + 1, 1 To 6)
    output(1, 1) = "PORT_CODE"
    output(1, 2) = "LEVEL"
    output(1, 3) = "CODE"
    output(1, 4) = "NAME"
    output(1, 5) = "ABSOLUTE_CONTRIBUTION"
    output(1, 6) = "STATUS"
    outputRow = 1
    For Each item In attribution.Keys
        outputRow = outputRow + 1
        values = attribution(item)
        output(outputRow, 1) = values(0)
        output(outputRow, 2) = values(1)
        output(outputRow, 3) = values(2)
        output(outputRow, 4) = values(3)
        output(outputRow, 5) = values(4)
        output(outputRow, 6) = values(5)
    Next item
    AbsoluteAttribution = output
End Function

Public Function ReturnAnalysis(Optional ByVal metricKeys As Variant) As Variant
    Dim activeMetrics As Variant, seriesByPort As Object, portKey As Variant
    Dim ts As clsTimeSeriesFrame, yearGroups As Object, yrKey As Variant, yrTs As clsTimeSeriesFrame
    Dim rows As New Collection, rowItem As Variant, output() As Variant
    Dim mCount As Long, mIdx As Long, r As Long, c As Long, status As String, valOut As Variant

    If IsMissing(metricKeys) Then
        activeMetrics = Array("RETURN_RATE", "BENCHMARK_RETURN", "EXCESS_RETURN", "ANNUALIZED_VOLATILITY", "SHARPE_RATIO", "SORTINO_RATIO")
    Else
        activeMetrics = metricKeys
    End If
    mCount = UBound(activeMetrics) - LBound(activeMetrics) + 1

    Set seriesByPort = BuildPortfolioReturnFrames()
    For Each portKey In seriesByPort.Keys
        Set ts = seriesByPort(portKey)
        Set yearGroups = ts.GroupByCalendarYear()
        For Each yrKey In yearGroups.Keys
            Set yrTs = yearGroups(yrKey)
            ReDim rowItem(1 To 6 + mCount)
            rowItem(1) = CStr(portKey)
            rowItem(2) = CStr(yrKey)
            rowItem(3) = yrTs.startDate
            rowItem(4) = yrTs.endDate
            rowItem(5) = yrTs.rowCount
            status = IIf(yrTs.rowCount < 2, "INSUFFICIENT_OBSERVATIONS", "OK")

            For mIdx = 1 To mCount
                valOut = yrTs.EvaluateMetric(CStr(activeMetrics(LBound(activeMetrics) + mIdx - 1)), "RTN", "RTN_BM")
                If IsNull(valOut) Then
                    rowItem(5 + mIdx) = Empty
                    If status = "OK" Then status = "UNDEFINED_RISK_RATIO"
                Else
                    rowItem(5 + mIdx) = valOut
                End If
            Next mIdx
            rowItem(6 + mCount) = status
            rows.Add rowItem
        Next yrKey
    Next portKey

    ReDim output(1 To rows.count + 1, 1 To 6 + mCount)
    output(1, 1) = "PORT_CODE": output(1, 2) = "CALENDAR_YEAR": output(1, 3) = "START_DATE"
    output(1, 4) = "END_DATE": output(1, 5) = "OBSERVATIONS"
    For mIdx = 1 To mCount
        output(1, 5 + mIdx) = UCase$(Trim$(CStr(activeMetrics(LBound(activeMetrics) + mIdx - 1))))
    Next mIdx
    output(1, 6 + mCount) = "STATUS"

    For r = 1 To rows.count
        rowItem = rows(r)
        For c = 1 To 6 + mCount
            output(r + 1, c) = rowItem(c)
        Next c
    Next r
    ReturnAnalysis = output
End Function

Public Function BuildPortfolioReturnFrames() As Object
    Dim performance As clsTimeSeriesFrame, ranged As clsTimeSeriesFrame, sorted As clsTimeSeriesFrame
    Dim groups As Object, groupKey As Variant, portfolioCode As String, result As Object
    Dim portfolioFrame As clsTimeSeriesFrame, normalized As clsTimeSeriesFrame
    Dim rows As Collection, r As Long, portColumn As Long, table As Variant

    Set performance = m_Panel.performanceFrame
    Set ranged = performance.FilterDateRange(m_Panel.startDate, m_Panel.endDate)
    Set sorted = ranged.SortedByDate()
    portColumn = sorted.FieldIndex("PORT_CODE")
    Call sorted.FieldIndex("RTN")
    Call sorted.FieldIndex("RTN_BM")
    table = sorted.ToTable()
    For r = 1 To sorted.rowCount
        If Not sorted.HasDate(r) Then Err.Raise 5, "clsPortfolioAnalyzer.BuildPortfolioReturnFrames", "Performance row has no date."
        If IsNull(sorted.ValueAt(r, "PORT_CODE")) Or IsEmpty(sorted.ValueAt(r, "PORT_CODE")) Then Err.Raise 5, "clsPortfolioAnalyzer.BuildPortfolioReturnFrames", "Performance row has no portfolio code."
        If IsNull(sorted.ValueAt(r, "RTN")) Or Not IsNumeric(sorted.ValueAt(r, "RTN")) Then Err.Raise 13, "clsPortfolioAnalyzer.BuildPortfolioReturnFrames", "Missing/non-numeric RTN: " & Format$(CDate(sorted.RowDate(r)), "yyyy-mm-dd")
        If IsNull(sorted.ValueAt(r, "RTN_BM")) Or Not IsNumeric(sorted.ValueAt(r, "RTN_BM")) Then Err.Raise 13, "clsPortfolioAnalyzer.BuildPortfolioReturnFrames", "Missing/non-numeric RTN_BM: " & Format$(CDate(sorted.RowDate(r)), "yyyy-mm-dd")
        If CDbl(sorted.ValueAt(r, "RTN")) <= -1# Or CDbl(sorted.ValueAt(r, "RTN_BM")) <= -1# Then Err.Raise 5, "clsPortfolioAnalyzer.BuildPortfolioReturnFrames", "Return must be greater than -100%."
        portfolioCode = UCase$(Trim$(CStr(sorted.ValueAt(r, "PORT_CODE"))))
        If Len(portfolioCode) = 0 Then Err.Raise 5, "clsPortfolioAnalyzer.BuildPortfolioReturnFrames", "Performance row has a blank portfolio code."
        table(r + 1, portColumn) = portfolioCode
    Next r
    Set normalized = New clsTimeSeriesFrame
    normalized.Load table, sorted.DateField
    Set groups = normalized.GroupRows(Array("PORT_CODE"))
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1
    For Each groupKey In groups.Keys
        Set rows = groups(groupKey)
        Set portfolioFrame = normalized.SelectRows(rows)
        portfolioFrame.ValidateSingleSeries
        portfolioCode = CStr(portfolioFrame.ValueAt(1, "PORT_CODE"))
        result.Add portfolioCode, portfolioFrame
    Next groupKey
    Set BuildPortfolioReturnFrames = result
End Function

Private Sub ValidateFrameFields(ByVal frame As clsTimeSeriesFrame, ByVal fieldNames As Variant, ByVal label As String)
    Dim i As Long, fieldName As String
    For i = LBound(fieldNames) To UBound(fieldNames)
        fieldName = CStr(fieldNames(i))
        If Not frame.HasField(fieldName) Then Err.Raise 5, "clsPortfolioAnalyzer", "Missing " & label & " field: " & fieldName
    Next i
End Sub

Private Sub ValidateDataTableFields(ByVal table As clsDataTable, ByVal fieldNames As Variant, ByVal label As String)
    Dim i As Long, fieldName As String
    For i = LBound(fieldNames) To UBound(fieldNames)
        fieldName = CStr(fieldNames(i))
        If Not table.HasField(fieldName) Then Err.Raise 5, "clsPortfolioAnalyzer", "Missing " & label & " field: " & fieldName
    Next i
End Sub
