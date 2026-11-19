Attribute VB_Name = "Mod_TechnicalOps"
Option Explicit

' ==============================================================================
' 1. 独立完备技术指标公式库（输入单只股票的 clsTimeSeriesFrame）
' ==============================================================================

' 简单移动平均线 SMA(N)，返回 1..RowCount 的 Variant 数组（前 N-1 项为 Null）
Public Function TA_SMA(ByVal frame As clsTimeSeriesFrame, _
        ByVal valueField As String, _
        ByVal window As Long) As Variant
    Dim ordered As clsTimeSeriesFrame, values() As Double
    Dim n As Long, i As Long, rollingSum As Double, outArr() As Variant
    If window < 1 Then Err.Raise 5, "Mod_TechnicalOps.TA_SMA", "Window must be at least 1."
    If Not PrepareOrderedFrame(frame, ordered) Then Exit Function
    If Not ExtractNumericSeries(ordered, valueField, values) Then Exit Function

    n = UBound(values)
    ReDim outArr(1 To n)
    For i = 1 To n
        rollingSum = rollingSum + values(i)
        If i > window Then rollingSum = rollingSum - values(i - window)
        If i < window Then
            outArr(i) = Null
        Else
            outArr(i) = rollingSum / window
        End If
    Next i
    TA_SMA = outArr
End Function

' 指数移动平均线 EMA(Span)，alpha = 2 / (Span + 1)
Public Function TA_EMA(ByVal frame As clsTimeSeriesFrame, _
        ByVal valueField As String, _
        ByVal span As Long) As Variant
    Dim ordered As clsTimeSeriesFrame, values() As Double, emaVec() As Double
    Dim n As Long, i As Long, outArr() As Variant
    If span < 1 Then Err.Raise 5, "Mod_TechnicalOps.TA_EMA", "Span must be at least 1."
    If Not PrepareOrderedFrame(frame, ordered) Then Exit Function
    If Not ExtractNumericSeries(ordered, valueField, values) Then Exit Function

    emaVec = ComputeEmaVector(values, span)
    n = UBound(emaVec)
    ReDim outArr(1 To n)
    For i = 1 To n
        outArr(i) = emaVec(i)
    Next i
    TA_EMA = outArr
End Function

' MACD 指标（默认 12, 26, 9），返回二维 Variant(1..N, 1..3)：第1列 DIF，第2列 DEA，第3列 MACD 柱(2*(DIF-DEA))
Public Function TA_MACD(ByVal frame As clsTimeSeriesFrame, _
        ByVal valueField As String, _
        Optional ByVal fastSpan As Long = 12, _
        Optional ByVal slowSpan As Long = 26, _
        Optional ByVal signalSpan As Long = 9) As Variant
    Dim ordered As clsTimeSeriesFrame, values() As Double
    Dim emaFast() As Double, emaSlow() As Double, dif() As Double, dea() As Double
    Dim n As Long, i As Long, outMat() As Variant
    If fastSpan < 1 Or slowSpan <= fastSpan Or signalSpan < 1 Then
        Err.Raise 5, "Mod_TechnicalOps.TA_MACD", "Invalid MACD spans (require 1 <= fast < slow and signal >= 1)."
    End If
    If Not PrepareOrderedFrame(frame, ordered) Then Exit Function
    If Not ExtractNumericSeries(ordered, valueField, values) Then Exit Function

    n = UBound(values)
    emaFast = ComputeEmaVector(values, fastSpan)
    emaSlow = ComputeEmaVector(values, slowSpan)
    ReDim dif(1 To n)
    For i = 1 To n
        dif(i) = emaFast(i) - emaSlow(i)
    Next i
    dea = ComputeEmaVector(dif, signalSpan)

    ReDim outMat(1 To n, 1 To 3)
    For i = 1 To n
        outMat(i, 1) = dif(i)
        outMat(i, 2) = dea(i)
        outMat(i, 3) = 2# * (dif(i) - dea(i))
    Next i
    TA_MACD = outMat
End Function

' Wilder 平滑相对强弱指标 RSI(N)（默认 14），返回 1..N 的 Variant 数组（0..100）
Public Function TA_RSI(ByVal frame As clsTimeSeriesFrame, _
        ByVal valueField As String, _
        Optional ByVal window As Long = 14) As Variant
    Dim ordered As clsTimeSeriesFrame, values() As Double
    Dim n As Long, i As Long, changeVal As Double, gain As Double, loss As Double
    Dim avgGain As Double, avgLoss As Double, rs As Double, outArr() As Variant
    If window < 1 Then Err.Raise 5, "Mod_TechnicalOps.TA_RSI", "Window must be at least 1."
    If Not PrepareOrderedFrame(frame, ordered) Then Exit Function
    If Not ExtractNumericSeries(ordered, valueField, values) Then Exit Function

    n = UBound(values)
    ReDim outArr(1 To n)
    For i = 1 To n: outArr(i) = Null: Next i
    If n <= window Then
        TA_RSI = outArr
        Exit Function
    End If

    For i = 2 To window + 1
        changeVal = values(i) - values(i - 1)
        If changeVal > 0# Then avgGain = avgGain + changeVal Else avgLoss = avgLoss - changeVal
    Next i
    avgGain = avgGain / window
    avgLoss = avgLoss / window
    outArr(window + 1) = RsiFromAverages(avgGain, avgLoss)

    For i = window + 2 To n
        changeVal = values(i) - values(i - 1)
        gain = IIf(changeVal > 0#, changeVal, 0#)
        loss = IIf(changeVal < 0#, -changeVal, 0#)
        avgGain = (avgGain * (window - 1) + gain) / window
        avgLoss = (avgLoss * (window - 1) + loss) / window
        outArr(i) = RsiFromAverages(avgGain, avgLoss)
    Next i
    TA_RSI = outArr
End Function

' 布林带指标（默认 20 日，2 倍样本标准差），返回二维 Variant(1..N, 1..4)：MID, UPPER, LOWER, PCT_B
Public Function TA_BollingerBands(ByVal frame As clsTimeSeriesFrame, _
        ByVal valueField As String, _
        Optional ByVal window As Long = 20, _
        Optional ByVal numStdDev As Double = 2#) As Variant
    Dim ordered As clsTimeSeriesFrame, values() As Double
    Dim n As Long, i As Long, k As Long, meanVal As Double, diff As Double, sqSum As Double, sd As Double
    Dim upperVal As Double, lowerVal As Double, bandWidth As Double, outMat() As Variant
    If window < 2 Then Err.Raise 5, "Mod_TechnicalOps.TA_BollingerBands", "Bollinger window must be at least 2."
    If numStdDev <= 0# Then Err.Raise 5, "Mod_TechnicalOps.TA_BollingerBands", "Standard deviation multiplier must be positive."
    If Not PrepareOrderedFrame(frame, ordered) Then Exit Function
    If Not ExtractNumericSeries(ordered, valueField, values) Then Exit Function

    n = UBound(values)
    ReDim outMat(1 To n, 1 To 4)
    For i = 1 To n
        If i < window Then
            outMat(i, 1) = Null: outMat(i, 2) = Null: outMat(i, 3) = Null: outMat(i, 4) = Null
        Else
            meanVal = 0#
            For k = i - window + 1 To i
                meanVal = meanVal + values(k)
            Next k
            meanVal = meanVal / window

            sqSum = 0#
            For k = i - window + 1 To i
                diff = values(k) - meanVal
                sqSum = sqSum + diff * diff
            Next k
            sd = Sqr(sqSum / (window - 1))

            upperVal = meanVal + numStdDev * sd
            lowerVal = meanVal - numStdDev * sd
            bandWidth = upperVal - lowerVal

            outMat(i, 1) = meanVal
            outMat(i, 2) = upperVal
            outMat(i, 3) = lowerVal
            If bandWidth > 0.0000001 Then
                outMat(i, 4) = (values(i) - lowerVal) / bandWidth
            Else
                outMat(i, 4) = Null
            End If
        End If
    Next i
    TA_BollingerBands = outMat
End Function

' 平均真实波幅 ATR(N)（Wilder 平滑，默认 14 日）
Public Function TA_ATR(ByVal frame As clsTimeSeriesFrame, _
        Optional ByVal window As Long = 14, _
        Optional ByVal highField As String = "ADJ_HIGH", _
        Optional ByVal lowField As String = "ADJ_LOW", _
        Optional ByVal closeField As String = "ADJ_CLOSE") As Variant
    Dim ordered As clsTimeSeriesFrame, highs() As Double, lows() As Double, closes() As Double
    Dim n As Long, i As Long, tr() As Double, hl As Double, hc As Double, lc As Double
    Dim atrVal As Double, outArr() As Variant
    If window < 1 Then Err.Raise 5, "Mod_TechnicalOps.TA_ATR", "ATR window must be at least 1."
    If Not PrepareOrderedFrame(frame, ordered) Then Exit Function
    If Not ExtractNumericSeries(ordered, highField, highs) Then Exit Function
    If Not ExtractNumericSeries(ordered, lowField, lows) Then Exit Function
    If Not ExtractNumericSeries(ordered, closeField, closes) Then Exit Function

    n = UBound(closes)
    ReDim tr(1 To n)
    ReDim outArr(1 To n)
    tr(1) = highs(1) - lows(1)
    For i = 2 To n
        hl = highs(i) - lows(i)
        hc = Abs(highs(i) - closes(i - 1))
        lc = Abs(lows(i) - closes(i - 1))
        tr(i) = hl
        If hc > tr(i) Then tr(i) = hc
        If lc > tr(i) Then tr(i) = lc
    Next i

    For i = 1 To n: outArr(i) = Null: Next i
    If n < window Then
        TA_ATR = outArr
        Exit Function
    End If

    For i = 1 To window
        atrVal = atrVal + tr(i)
    Next i
    atrVal = atrVal / window
    outArr(window) = atrVal

    For i = window + 1 To n
        atrVal = (atrVal * (window - 1) + tr(i)) / window
        outArr(i) = atrVal
    Next i
    TA_ATR = outArr
End Function

' ==============================================================================
' 2. 综合技术指标明细表生成器（支持按 displayStartDate 裁掉预热期）
' ==============================================================================

Public Function TA_BuildIndicatorTable(ByVal frame As clsTimeSeriesFrame, _
        Optional ByVal displayStartDate As Variant, _
        Optional ByVal priceField As String = "ADJ_CLOSE") As clsDataTable
    Dim ordered As clsTimeSeriesFrame, n As Long, i As Long, outCount As Long, r As Long
    Dim cutoffDate As Date, hasCutoff As Boolean, rowDt As Date
    Dim sma20 As Variant, sma60 As Variant, sma250 As Variant
    Dim macdMat As Variant, rsi14 As Variant, bollMat As Variant, atr14 As Variant
    Dim tableData() As Variant, resTable As clsDataTable

    If Not PrepareOrderedFrame(frame, ordered) Then Err.Raise 5, "Mod_TechnicalOps.TA_BuildIndicatorTable", "Non-empty quote frame is required."
    n = ordered.rowCount
    If Not IsMissing(displayStartDate) And Not IsEmpty(displayStartDate) Then
        cutoffDate = dateValue(CDate(displayStartDate))
        hasCutoff = True
    End If

    sma20 = TA_SMA(ordered, priceField, 20)
    sma60 = TA_SMA(ordered, priceField, 60)
    sma250 = TA_SMA(ordered, priceField, 250)
    macdMat = TA_MACD(ordered, priceField, 12, 26, 9)
    rsi14 = TA_RSI(ordered, priceField, 14)
    bollMat = TA_BollingerBands(ordered, priceField, 20, 2#)
    atr14 = TA_ATR(ordered, 14)

    For i = 1 To n
        rowDt = dateValue(CDate(ordered.RowDate(i)))
        If Not hasCutoff Or rowDt >= cutoffDate Then outCount = outCount + 1
    Next i

    ReDim tableData(1 To outCount + 1, 1 To 15)
    tableData(1, 1) = "INNER_CODE": tableData(1, 2) = "T_DATE": tableData(1, 3) = "ADJ_CLOSE"
    tableData(1, 4) = "DAILY_RTN": tableData(1, 5) = "SMA_20": tableData(1, 6) = "SMA_60"
    tableData(1, 7) = "SMA_250": tableData(1, 8) = "MACD_DIF": tableData(1, 9) = "MACD_DEA"
    tableData(1, 10) = "MACD_HIST": tableData(1, 11) = "RSI_14": tableData(1, 12) = "BOLL_MID"
    tableData(1, 13) = "BOLL_UPPER": tableData(1, 14) = "BOLL_LOWER": tableData(1, 15) = "ATR_14"

    Dim hasInner As Boolean, hasRtn As Boolean
    Dim hasSma20 As Boolean, hasSma60 As Boolean, hasSma250 As Boolean
    Dim hasMacd As Boolean, hasRsi As Boolean, hasBoll As Boolean, hasAtr As Boolean

    hasInner = ordered.HasField("INNER_CODE")
    hasRtn = ordered.HasField("DAILY_RTN")
    hasSma20 = IsArray(sma20)
    hasSma60 = IsArray(sma60)
    hasSma250 = IsArray(sma250)
    hasMacd = IsArray(macdMat)
    hasRsi = IsArray(rsi14)
    hasBoll = IsArray(bollMat)
    hasAtr = IsArray(atr14)

    r = 1
    For i = 1 To n
        rowDt = dateValue(CDate(ordered.RowDate(i)))
        If Not hasCutoff Or rowDt >= cutoffDate Then
            r = r + 1
            If hasInner Then tableData(r, 1) = ordered.ValueAt(i, "INNER_CODE") Else tableData(r, 1) = Null
            tableData(r, 2) = rowDt
            tableData(r, 3) = ordered.ValueAt(i, priceField)
            If hasRtn Then tableData(r, 4) = ordered.ValueAt(i, "DAILY_RTN") Else tableData(r, 4) = Null
            If hasSma20 Then tableData(r, 5) = sma20(i) Else tableData(r, 5) = Null
            If hasSma60 Then tableData(r, 6) = sma60(i) Else tableData(r, 6) = Null
            If hasSma250 Then tableData(r, 7) = sma250(i) Else tableData(r, 7) = Null
            If hasMacd Then
                tableData(r, 8) = macdMat(i, 1)
                tableData(r, 9) = macdMat(i, 2)
                tableData(r, 10) = macdMat(i, 3)
            Else
                tableData(r, 8) = Null: tableData(r, 9) = Null: tableData(r, 10) = Null
            End If
            If hasRsi Then tableData(r, 11) = rsi14(i) Else tableData(r, 11) = Null
            If hasBoll Then
                tableData(r, 12) = bollMat(i, 1)
                tableData(r, 13) = bollMat(i, 2)
                tableData(r, 14) = bollMat(i, 3)
            Else
                tableData(r, 12) = Null: tableData(r, 13) = Null: tableData(r, 14) = Null
            End If
            If hasAtr Then tableData(r, 15) = atr14(i) Else tableData(r, 15) = Null
        End If
    Next i

    Set resTable = New clsDataTable
    resTable.Load tableData
    Set TA_BuildIndicatorTable = resTable
End Function

' ==============================================================================
' 3. 内部校验与向量辅助原语
' ==============================================================================

Private Function PrepareOrderedFrame(ByVal frame As clsTimeSeriesFrame, ByRef outOrdered As clsTimeSeriesFrame) As Boolean
    If frame Is Nothing Then Exit Function
    frame.ValidateSingleSeries
    If frame.rowCount = 0 Then Exit Function
    Set outOrdered = frame.SortedByDate()
    PrepareOrderedFrame = True
End Function

Private Function ExtractNumericSeries(ByVal ordered As clsTimeSeriesFrame, _
        ByVal fieldName As String, _
        ByRef outValues() As Double) As Boolean
    Dim n As Long, i As Long, rawVal As Variant
    If Len(Trim$(fieldName)) = 0 Then Err.Raise 5, "Mod_TechnicalOps.ExtractNumericSeries", "Field name is required."
    Call ordered.FieldIndex(fieldName)
    n = ordered.rowCount
    If n = 0 Then Exit Function
    ReDim outValues(1 To n)
    For i = 1 To n
        If Not ordered.HasDate(i) Then Err.Raise 5, "Mod_TechnicalOps.ExtractNumericSeries", "Every observation requires a valid date."
        rawVal = ordered.ValueAt(i, fieldName)
        If IsNull(rawVal) Or IsEmpty(rawVal) Then
            If i > 1 Then
                outValues(i) = outValues(i - 1)
            Else
                Exit Function
            End If
        ElseIf IsError(rawVal) Or Not IsNumeric(rawVal) Then
            Err.Raise 13, "Mod_TechnicalOps.ExtractNumericSeries", "Non-numeric value in " & fieldName
        Else
            outValues(i) = CDbl(rawVal)
        End If
    Next i
    ExtractNumericSeries = True
End Function

Private Function ComputeEmaVector(ByRef values() As Double, ByVal span As Long) As Double()
    Dim n As Long, i As Long, alpha As Double, ema() As Double
    n = UBound(values)
    ReDim ema(1 To n)
    alpha = 2# / (span + 1#)
    ema(1) = values(1)
    For i = 2 To n
        ema(i) = alpha * values(i) + (1# - alpha) * ema(i - 1)
    Next i
    ComputeEmaVector = ema
End Function

Private Function RsiFromAverages(ByVal avgGain As Double, ByVal avgLoss As Double) As Double
    Dim rs As Double
    If avgLoss = 0# Then
        RsiFromAverages = IIf(avgGain = 0#, 50#, 100#)
    Else
        rs = avgGain / avgLoss
        RsiFromAverages = 100# - (100# / (1# + rs))
    End If
End Function
