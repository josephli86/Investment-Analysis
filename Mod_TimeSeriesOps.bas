Attribute VB_Name = "Mod_TimeSeriesOps"
Option Explicit

' ==============================================================================
' 1. 完备指标公式库（每个函数独立完成校验、取数与计算，可直接被业务层调用）
' ==============================================================================

Public Function TS_Observations(ByVal frame As clsTimeSeriesFrame) As Variant
    Dim ordered As clsTimeSeriesFrame
    TS_Observations = Null
    If Not PrepareOrderedFrame(frame, ordered) Then Exit Function
    TS_Observations = ordered.rowCount
End Function

Public Function TS_Sum(ByVal frame As clsTimeSeriesFrame, ByVal valueField As String) As Variant
    Dim ordered As clsTimeSeriesFrame, values() As Double
    Dim i As Long, total As Double
    TS_Sum = Null
    If Not PrepareOrderedFrame(frame, ordered) Then Exit Function
    If Not ExtractNumericSeries(ordered, valueField, "valueField", values) Then Exit Function
    For i = 1 To UBound(values)
        total = total + values(i)
    Next i
    TS_Sum = total
End Function

Public Function TS_Average(ByVal frame As clsTimeSeriesFrame, ByVal valueField As String) As Variant
    Dim ordered As clsTimeSeriesFrame, values() As Double
    Dim i As Long, total As Double
    TS_Average = Null
    If Not PrepareOrderedFrame(frame, ordered) Then Exit Function
    If Not ExtractNumericSeries(ordered, valueField, "valueField", values) Then Exit Function
    For i = 1 To UBound(values)
        total = total + values(i)
    Next i
    TS_Average = total / UBound(values)
End Function

Public Function TS_Min(ByVal frame As clsTimeSeriesFrame, ByVal valueField As String) As Variant
    Dim ordered As clsTimeSeriesFrame, values() As Double
    Dim i As Long, minimum As Double
    TS_Min = Null
    If Not PrepareOrderedFrame(frame, ordered) Then Exit Function
    If Not ExtractNumericSeries(ordered, valueField, "valueField", values) Then Exit Function
    minimum = values(1)
    For i = 2 To UBound(values)
        If values(i) < minimum Then minimum = values(i)
    Next i
    TS_Min = minimum
End Function

Public Function TS_Max(ByVal frame As clsTimeSeriesFrame, ByVal valueField As String) As Variant
    Dim ordered As clsTimeSeriesFrame, values() As Double
    Dim i As Long, maximum As Double
    TS_Max = Null
    If Not PrepareOrderedFrame(frame, ordered) Then Exit Function
    If Not ExtractNumericSeries(ordered, valueField, "valueField", values) Then Exit Function
    maximum = values(1)
    For i = 2 To UBound(values)
        If values(i) > maximum Then maximum = values(i)
    Next i
    TS_Max = maximum
End Function

Public Function TS_StandardDeviation(ByVal frame As clsTimeSeriesFrame, ByVal valueField As String) As Variant
    Dim ordered As clsTimeSeriesFrame, values() As Double
    TS_StandardDeviation = Null
    If Not PrepareOrderedFrame(frame, ordered) Then Exit Function
    If Not ExtractNumericSeries(ordered, valueField, "valueField", values) Then Exit Function
    TS_StandardDeviation = SampleStdDev(values)
End Function

Public Function TS_Volatility(ByVal frame As clsTimeSeriesFrame, _
        ByVal valueField As String, _
        Optional ByVal periodsPerYear As Double = 252#) As Variant
    Dim sd As Variant
    ValidatePeriodsPerYear periodsPerYear, "TS_Volatility"
    TS_Volatility = Null
    sd = TS_StandardDeviation(frame, valueField)
    If IsNull(sd) Then Exit Function
    TS_Volatility = CDbl(sd) * Sqr(periodsPerYear)
End Function

Public Function TS_CumulativeReturn(ByVal frame As clsTimeSeriesFrame, ByVal valueField As String) As Variant
    Dim ordered As clsTimeSeriesFrame, returns() As Double
    Dim i As Long, wealth As Double
    TS_CumulativeReturn = Null
    If Not PrepareOrderedFrame(frame, ordered) Then Exit Function
    If Not ExtractReturnSeries(ordered, valueField, "valueField", returns) Then Exit Function
    wealth = 1#
    For i = 1 To UBound(returns)
        wealth = wealth * (1# + returns(i))
    Next i
    TS_CumulativeReturn = wealth - 1#
End Function

Public Function TS_AnnualizedReturn(ByVal frame As clsTimeSeriesFrame, _
        ByVal valueField As String, _
        Optional ByVal periodsPerYear As Double = 252#) As Variant
    Dim ordered As clsTimeSeriesFrame, returns() As Double
    Dim i As Long, n As Long, wealth As Double
    ValidatePeriodsPerYear periodsPerYear, "TS_AnnualizedReturn"
    TS_AnnualizedReturn = Null
    If Not PrepareOrderedFrame(frame, ordered) Then Exit Function
    If Not ExtractReturnSeries(ordered, valueField, "valueField", returns) Then Exit Function
    n = UBound(returns)
    wealth = 1#
    For i = 1 To n
        wealth = wealth * (1# + returns(i))
    Next i
    If wealth <= 0# Then
        TS_AnnualizedReturn = -1#
    Else
        TS_AnnualizedReturn = wealth ^ (periodsPerYear / n) - 1#
    End If
End Function

Public Function TS_MaxDrawdown(ByVal frame As clsTimeSeriesFrame, ByVal valueField As String) As Variant
    Dim ordered As clsTimeSeriesFrame, returns() As Double
    Dim i As Long, wealth As Double, peak As Double, drawdown As Double, maxDrawdown As Double
    TS_MaxDrawdown = Null
    If Not PrepareOrderedFrame(frame, ordered) Then Exit Function
    If Not ExtractReturnSeries(ordered, valueField, "valueField", returns) Then Exit Function
    wealth = 1#: peak = 1#: maxDrawdown = 0#
    For i = 1 To UBound(returns)
        wealth = wealth * (1# + returns(i))
        If wealth > peak Then peak = wealth
        If peak > 0# Then
            drawdown = (wealth - peak) / peak
            If drawdown < maxDrawdown Then maxDrawdown = drawdown
        End If
    Next i
    TS_MaxDrawdown = maxDrawdown
End Function

Public Function TS_CalmarRatio(ByVal frame As clsTimeSeriesFrame, _
        ByVal valueField As String, _
        Optional ByVal periodsPerYear As Double = 252#) As Variant
    Dim annReturn As Variant, maxDd As Variant
    ValidatePeriodsPerYear periodsPerYear, "TS_CalmarRatio"
    TS_CalmarRatio = Null
    annReturn = TS_AnnualizedReturn(frame, valueField, periodsPerYear)
    If IsNull(annReturn) Then Exit Function
    maxDd = TS_MaxDrawdown(frame, valueField)
    If IsNull(maxDd) Then Exit Function
    If Abs(CDbl(maxDd)) < 0.0000001 Then Exit Function
    TS_CalmarRatio = CDbl(annReturn) / Abs(CDbl(maxDd))
End Function

Public Function TS_ExcessReturn(ByVal frame As clsTimeSeriesFrame, _
        ByVal valueField As String, _
        ByVal benchmarkField As String) As Variant
    Dim portReturn As Variant, bmReturn As Variant
    TS_ExcessReturn = Null
    portReturn = TS_CumulativeReturn(frame, valueField)
    If IsNull(portReturn) Then Exit Function
    bmReturn = TS_CumulativeReturn(frame, benchmarkField)
    If IsNull(bmReturn) Then Exit Function
    TS_ExcessReturn = CDbl(portReturn) - CDbl(bmReturn)
End Function

Public Function TS_TrackingError(ByVal frame As clsTimeSeriesFrame, _
        ByVal valueField As String, _
        ByVal benchmarkField As String, _
        Optional ByVal periodsPerYear As Double = 252#) As Variant
    Dim ordered As clsTimeSeriesFrame, returns() As Double, benchmarks() As Double, active() As Double
    Dim sd As Variant
    ValidatePeriodsPerYear periodsPerYear, "TS_TrackingError"
    TS_TrackingError = Null
    If Not PrepareOrderedFrame(frame, ordered) Then Exit Function
    If Not ExtractNumericSeries(ordered, valueField, "valueField", returns) Then Exit Function
    If Not ExtractNumericSeries(ordered, benchmarkField, "benchmarkField", benchmarks) Then Exit Function
    active = SubtractVectors(returns, benchmarks)
    sd = SampleStdDev(active)
    If IsNull(sd) Then Exit Function
    TS_TrackingError = CDbl(sd) * Sqr(periodsPerYear)
End Function

Public Function TS_InformationRatio(ByVal frame As clsTimeSeriesFrame, _
        ByVal valueField As String, _
        ByVal benchmarkField As String, _
        Optional ByVal periodsPerYear As Double = 252#) As Variant
    Dim ordered As clsTimeSeriesFrame, returns() As Double, benchmarks() As Double, active() As Double
    Dim sd As Variant, te As Double, meanActive As Double, i As Long, n As Long
    ValidatePeriodsPerYear periodsPerYear, "TS_InformationRatio"
    TS_InformationRatio = Null
    If Not PrepareOrderedFrame(frame, ordered) Then Exit Function
    If Not ExtractNumericSeries(ordered, valueField, "valueField", returns) Then Exit Function
    If Not ExtractNumericSeries(ordered, benchmarkField, "benchmarkField", benchmarks) Then Exit Function
    active = SubtractVectors(returns, benchmarks)
    sd = SampleStdDev(active)
    If IsNull(sd) Then Exit Function
    te = CDbl(sd) * Sqr(periodsPerYear)
    If te = 0# Then Exit Function
    n = UBound(active)
    For i = 1 To n
        meanActive = meanActive + active(i)
    Next i
    meanActive = meanActive / n
    TS_InformationRatio = (meanActive * periodsPerYear) / te
End Function

Public Function TS_SharpeRatio(ByVal frame As clsTimeSeriesFrame, _
        ByVal valueField As String, _
        Optional ByVal riskFreeField As String = "", _
        Optional ByVal periodsPerYear As Double = 252#) As Variant
    Dim ordered As clsTimeSeriesFrame, returns() As Double, riskFree() As Double, excess() As Double
    Dim sd As Variant, meanExcess As Double, i As Long, n As Long
    ValidatePeriodsPerYear periodsPerYear, "TS_SharpeRatio"
    TS_SharpeRatio = Null
    If Not PrepareOrderedFrame(frame, ordered) Then Exit Function
    If Not ExtractNumericSeries(ordered, valueField, "valueField", returns) Then Exit Function
    If Len(Trim$(riskFreeField)) > 0 Then
        If Not ExtractNumericSeries(ordered, riskFreeField, "riskFreeField", riskFree) Then Exit Function
        excess = SubtractVectors(returns, riskFree)
    Else
        excess = returns
    End If
    sd = SampleStdDev(excess)
    If IsNull(sd) Then Exit Function
    If CDbl(sd) = 0# Then Exit Function
    n = UBound(excess)
    For i = 1 To n
        meanExcess = meanExcess + excess(i)
    Next i
    meanExcess = meanExcess / n
    TS_SharpeRatio = (meanExcess / CDbl(sd)) * Sqr(periodsPerYear)
End Function

Public Function TS_SortinoRatio(ByVal frame As clsTimeSeriesFrame, _
        ByVal valueField As String, _
        Optional ByVal riskFreeField As String = "", _
        Optional ByVal targetReturn As Double = 0#, _
        Optional ByVal periodsPerYear As Double = 252#) As Variant
    Dim ordered As clsTimeSeriesFrame, returns() As Double, riskFree() As Double, excess() As Double
    Dim i As Long, n As Long, meanExcess As Double, shortfall As Double, downsideSq As Double, downsideRms As Double
    ValidatePeriodsPerYear periodsPerYear, "TS_SortinoRatio"
    TS_SortinoRatio = Null
    If Not PrepareOrderedFrame(frame, ordered) Then Exit Function
    If Not ExtractNumericSeries(ordered, valueField, "valueField", returns) Then Exit Function
    If Len(Trim$(riskFreeField)) > 0 Then
        If Not ExtractNumericSeries(ordered, riskFreeField, "riskFreeField", riskFree) Then Exit Function
        excess = SubtractVectors(returns, riskFree)
    Else
        excess = returns
    End If
    n = UBound(returns)
    If n < 2 Then Exit Function
    For i = 1 To n
        meanExcess = meanExcess + excess(i)
        ' 下行偏差按原始收益率相对目标收益率计算，保持原有口径一致
        shortfall = returns(i) - targetReturn
        If shortfall < 0# Then downsideSq = downsideSq + shortfall * shortfall
    Next i
    downsideRms = Sqr(downsideSq / n)
    If downsideRms = 0# Then Exit Function
    meanExcess = meanExcess / n
    TS_SortinoRatio = (meanExcess / downsideRms) * Sqr(periodsPerYear)
End Function

' ==============================================================================
' 2. 动态分发路由器（供 clsTimeSeriesFrame.EvaluateMetric 及批量指标列表调用）
' ==============================================================================

Public Function TS_EvalFrameMetric(ByVal metricKey As String, _
        ByVal frame As clsTimeSeriesFrame, _
        ByVal valueField As String, _
        Optional ByVal periodsPerYear As Double = 252#) As Variant
    Select Case UCase$(Trim$(metricKey))
        Case "OBSERVATIONS": TS_EvalFrameMetric = TS_Observations(frame)
        Case "SUM": TS_EvalFrameMetric = TS_Sum(frame, valueField)
        Case "AVERAGE": TS_EvalFrameMetric = TS_Average(frame, valueField)
        Case "MIN": TS_EvalFrameMetric = TS_Min(frame, valueField)
        Case "MAX": TS_EvalFrameMetric = TS_Max(frame, valueField)
        Case "STANDARD_DEVIATION": TS_EvalFrameMetric = TS_StandardDeviation(frame, valueField)
        Case "VOLATILITY": TS_EvalFrameMetric = TS_Volatility(frame, valueField, periodsPerYear)
        Case "CUMULATIVE_RETURN": TS_EvalFrameMetric = TS_CumulativeReturn(frame, valueField)
        Case "ANNUALIZED_RETURN": TS_EvalFrameMetric = TS_AnnualizedReturn(frame, valueField, periodsPerYear)
        Case "MAX_DRAWDOWN": TS_EvalFrameMetric = TS_MaxDrawdown(frame, valueField)
        Case Else
            Err.Raise 5, "Mod_TimeSeriesOps.TS_EvalFrameMetric", "Unsupported frame metric: " & metricKey
    End Select
End Function

Public Function TS_EvalFrameMetricEx(ByVal metricKey As String, _
        ByVal frame As clsTimeSeriesFrame, _
        ByVal valueField As String, _
        Optional ByVal benchmarkField As String = "", _
        Optional ByVal riskFreeField As String = "", _
        Optional ByVal targetReturn As Double = 0#, _
        Optional ByVal periodsPerYear As Double = 252#) As Variant
    Select Case UCase$(Trim$(metricKey))
        Case "RETURN_RATE", "CUMULATIVE_RETURN"
            TS_EvalFrameMetricEx = TS_CumulativeReturn(frame, valueField)
        Case "BENCHMARK_RETURN"
            TS_EvalFrameMetricEx = TS_CumulativeReturn(frame, benchmarkField)
        Case "EXCESS_RETURN"
            TS_EvalFrameMetricEx = TS_ExcessReturn(frame, valueField, benchmarkField)
        Case "ANNUALIZED_RETURN"
            TS_EvalFrameMetricEx = TS_AnnualizedReturn(frame, valueField, periodsPerYear)
        Case "ANNUALIZED_VOLATILITY", "VOLATILITY"
            TS_EvalFrameMetricEx = TS_Volatility(frame, valueField, periodsPerYear)
        Case "MAX_DRAWDOWN"
            TS_EvalFrameMetricEx = TS_MaxDrawdown(frame, valueField)
        Case "CALMAR_RATIO"
            TS_EvalFrameMetricEx = TS_CalmarRatio(frame, valueField, periodsPerYear)
        Case "TRACKING_ERROR"
            TS_EvalFrameMetricEx = TS_TrackingError(frame, valueField, benchmarkField, periodsPerYear)
        Case "INFORMATION_RATIO"
            TS_EvalFrameMetricEx = TS_InformationRatio(frame, valueField, benchmarkField, periodsPerYear)
        Case "SHARPE_RATIO"
            TS_EvalFrameMetricEx = TS_SharpeRatio(frame, valueField, riskFreeField, periodsPerYear)
        Case "SORTINO_RATIO"
            TS_EvalFrameMetricEx = TS_SortinoRatio(frame, valueField, riskFreeField, targetReturn, periodsPerYear)
        Case Else
            TS_EvalFrameMetricEx = TS_EvalFrameMetric(metricKey, frame, valueField, periodsPerYear)
    End Select
End Function

Public Function TS_EvalFrameMetricByGroup(ByVal metricKey As String, _
        ByVal frame As clsTimeSeriesFrame, _
        ByVal dimensions As Variant, _
        ByVal valueField As String, _
        Optional ByVal benchmarkField As String = "", _
        Optional ByVal riskFreeField As String = "", _
        Optional ByVal targetReturn As Double = 0#, _
        Optional ByVal periodsPerYear As Double = 252#) As Variant
    Dim groups As Object, groupKey As Variant, members As Collection
    Dim subset As clsTimeSeriesFrame, output() As Variant
    Dim dimensionCount As Long, dimensionIndex As Long
    Dim sourceRow As Long, outputRow As Long
    If frame Is Nothing Then Err.Raise 5, "Mod_TimeSeriesOps.TS_EvalFrameMetricByGroup", "A time-series frame is required."
    If Not IsArray(dimensions) Then Err.Raise 5, "Mod_TimeSeriesOps.TS_EvalFrameMetricByGroup", "Dimensions must be an array."
    dimensionCount = UBound(dimensions) - LBound(dimensions) + 1
    If dimensionCount < 1 Then Err.Raise 5, "Mod_TimeSeriesOps.TS_EvalFrameMetricByGroup", "At least one series dimension is required."
    Set groups = frame.GroupRows(dimensions)
    ReDim output(1 To groups.count + 1, 1 To dimensionCount + 2)
    For dimensionIndex = 0 To dimensionCount - 1
        output(1, dimensionIndex + 1) = CStr(dimensions(LBound(dimensions) + dimensionIndex))
    Next dimensionIndex
    output(1, dimensionCount + 1) = "OBSERVATIONS"
    output(1, dimensionCount + 2) = "METRIC_VALUE"
    outputRow = 1
    For Each groupKey In groups.Keys
        outputRow = outputRow + 1
        Set members = groups(groupKey)
        Set subset = frame.SelectRows(members)
        sourceRow = CLng(members(1))
        For dimensionIndex = 0 To dimensionCount - 1
            output(outputRow, dimensionIndex + 1) = frame.ValueAt(sourceRow, CStr(dimensions(LBound(dimensions) + dimensionIndex)))
        Next dimensionIndex
        output(outputRow, dimensionCount + 1) = subset.rowCount
        output(outputRow, dimensionCount + 2) = subset.EvaluateMetric(metricKey, valueField, benchmarkField, riskFreeField, targetReturn, periodsPerYear)
    Next groupKey
    TS_EvalFrameMetricByGroup = output
End Function

' ==============================================================================
' 3. 内部校验与向量化提取原语（供各完备公式复用）
' ==============================================================================

Private Sub ValidatePeriodsPerYear(ByVal periodsPerYear As Double, ByVal caller As String)
    If periodsPerYear <= 0# Then Err.Raise 5, "Mod_TimeSeriesOps." & caller, "Periods per year must be positive."
End Sub

Private Function PrepareOrderedFrame(ByVal frame As clsTimeSeriesFrame, ByRef outOrdered As clsTimeSeriesFrame) As Boolean
    If frame Is Nothing Then Exit Function
    frame.ValidateSingleSeries
    If frame.rowCount = 0 Then Exit Function
    Set outOrdered = frame.SortedByDate()
    PrepareOrderedFrame = True
End Function

Private Function ExtractNumericSeries(ByVal ordered As clsTimeSeriesFrame, _
        ByVal fieldName As String, _
        ByVal argumentName As String, _
        ByRef outValues() As Double) As Boolean
    Dim n As Long, i As Long, rawVal As Variant
    If Len(Trim$(fieldName)) = 0 Then
        Err.Raise 5, "Mod_TimeSeriesOps.ExtractNumericSeries", argumentName & " is required for this metric."
    End If
    Call ordered.FieldIndex(fieldName)
    n = ordered.rowCount
    If n = 0 Then Exit Function
    ReDim outValues(1 To n)
    For i = 1 To n
        If Not ordered.HasDate(i) Then Err.Raise 5, "Mod_TimeSeriesOps.ExtractNumericSeries", "Every observation requires a valid date."
        rawVal = ordered.ValueAt(i, fieldName)
        If IsNull(rawVal) Or IsEmpty(rawVal) Then Exit Function
        If IsError(rawVal) Then Err.Raise 13, "Mod_TimeSeriesOps.ExtractNumericSeries", "Error value in " & fieldName
        If Not IsNumeric(rawVal) Then Err.Raise 13, "Mod_TimeSeriesOps.ExtractNumericSeries", "Non-numeric value in " & fieldName
        outValues(i) = CDbl(rawVal)
    Next i
    ExtractNumericSeries = True
End Function

Private Function ExtractReturnSeries(ByVal ordered As clsTimeSeriesFrame, _
        ByVal fieldName As String, _
        ByVal argumentName As String, _
        ByRef outReturns() As Double) As Boolean
    Dim i As Long
    If Not ExtractNumericSeries(ordered, fieldName, argumentName, outReturns) Then Exit Function
    For i = 1 To UBound(outReturns)
        If outReturns(i) <= -1# Then
            Err.Raise 5, "Mod_TimeSeriesOps.ExtractReturnSeries", "Returns must be greater than -100%."
        End If
    Next i
    ExtractReturnSeries = True
End Function

Private Function SampleStdDev(ByRef values() As Double) As Variant
    Dim n As Long, i As Long, meanValue As Double, diff As Double, sumSquares As Double
    n = UBound(values)
    If n < 2 Then
        SampleStdDev = Null
        Exit Function
    End If
    For i = 1 To n
        meanValue = meanValue + values(i)
    Next i
    meanValue = meanValue / n
    For i = 1 To n
        diff = values(i) - meanValue
        sumSquares = sumSquares + diff * diff
    Next i
    SampleStdDev = Sqr(sumSquares / (n - 1))
End Function

Private Function SubtractVectors(ByRef leftSeries() As Double, ByRef rightSeries() As Double) As Double()
    Dim n As Long, i As Long, diff() As Double
    n = UBound(leftSeries)
    ReDim diff(1 To n)
    For i = 1 To n
        diff(i) = leftSeries(i) - rightSeries(i)
    Next i
    SubtractVectors = diff
End Function
