VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTimeSeriesFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Data As clsDataTable
Private m_DateField As String
Private m_DateValues() As Date
Private m_HasDate() As Boolean
Private m_DateRows As Object
Private m_RowCount As Long
Private m_FieldCount As Long
Private m_Loaded As Boolean

Public Sub Load(ByVal table As Variant, ByVal DateField As String)
    Dim data As clsDataTable
    If m_Loaded Then Err.Raise 5, "clsTimeSeriesFrame.Load", "Time-series frames are immutable."
    Set data = New clsDataTable
    data.Load table
    LoadData data, DateField
End Sub

Public Sub LoadData(ByVal data As clsDataTable, ByVal DateField As String)
    Dim r As Long, Value As Variant
    If m_Loaded Then Err.Raise 5, "clsTimeSeriesFrame.LoadData", "Time-series frames are immutable."
    If data Is Nothing Then Err.Raise 5, "clsTimeSeriesFrame.LoadData", "Data table is required."
    If Not data.IsLoaded Then Err.Raise 5, "clsTimeSeriesFrame.LoadData", "Data table is not loaded."
    m_DateField = UCase$(Trim$(DateField))
    Call data.FieldIndex(m_DateField)
    Set m_Data = data
    m_RowCount = data.rowCount
    m_FieldCount = data.fieldCount
    Set m_DateRows = CreateObject("Scripting.Dictionary")
    m_DateRows.CompareMode = 0
    ReDim m_DateValues(1 To MaxLong(1, m_RowCount))
    ReDim m_HasDate(1 To MaxLong(1, m_RowCount))
    For r = 1 To m_RowCount
        Value = data.ValueAt(r, m_DateField)
        If Not IsNull(Value) Then
            If Not IsEmpty(Value) Then
                If Not IsError(Value) Then
                    If IsDate(Value) Then
                        m_DateValues(r) = CDate(Value)
                        m_HasDate(r) = True
                        AddDateRow r
                    End If
                End If
            End If
        End If
    Next r
    m_Loaded = True
End Sub

Public Property Get IsLoaded() As Boolean
    IsLoaded = m_Loaded
End Property

Public Function HasField(ByVal fieldName As String) As Boolean
    HasField = m_Data.HasField(fieldName)
End Function

Public Function ToTable() As Variant
    ToTable = m_Data.ToTable
End Function

Public Function SelectRows(ByVal rows As Collection) As clsTimeSeriesFrame
    Dim result As clsTimeSeriesFrame, selected As clsDataTable
    Set selected = m_Data.SelectRows(rows)
    Set result = New clsTimeSeriesFrame
    result.LoadData selected, m_DateField
    Set SelectRows = result
End Function

Public Function FilterField(ByVal fieldName As String, ByVal Value As Variant) As clsTimeSeriesFrame
    Dim result As clsTimeSeriesFrame, selected As clsDataTable
    Set selected = m_Data.FilterField(fieldName, Value)
    Set result = New clsTimeSeriesFrame
    result.LoadData selected, m_DateField
    Set FilterField = result
End Function

Public Function GroupByCalendarYear() As Object
    Dim buckets As Object, result As Object, rows As Collection
    Dim r As Long, key As Variant, yearKey As String
    Set buckets = CreateObject("Scripting.Dictionary")
    Set result = CreateObject("Scripting.Dictionary")
    For r = 1 To m_RowCount
        If Not m_HasDate(r) Then Err.Raise 5, "clsTimeSeriesFrame.GroupByCalendarYear", "Missing or invalid date at row " & r
        yearKey = CStr(Year(m_DateValues(r)))
        If Not buckets.exists(yearKey) Then
            Set rows = New Collection
            buckets.Add yearKey, rows
        End If
        buckets(yearKey).Add r
    Next r
    For Each key In buckets.Keys
        result.Add CStr(key), SelectRows(buckets(key))
    Next key
    Set GroupByCalendarYear = result
End Function

Public Function EvaluateMetric(ByVal metricKey As String, ByVal valueField As String, Optional ByVal benchmarkField As String = "", Optional ByVal riskFreeField As String = "", Optional ByVal targetReturn As Double = 0#, Optional ByVal periodsPerYear As Double = 252#) As Variant
    EvaluateMetric = TS_EvalFrameMetricEx(metricKey, Me, valueField, benchmarkField, riskFreeField, targetReturn, periodsPerYear)
End Function

Public Function EvaluateMetricByGroup(ByVal metricKey As String, ByVal dimensions As Variant, ByVal valueField As String, Optional ByVal benchmarkField As String = "", Optional ByVal riskFreeField As String = "", Optional ByVal targetReturn As Double = 0#, Optional ByVal periodsPerYear As Double = 252#) As Variant
    EvaluateMetricByGroup = TS_EvalFrameMetricByGroup(metricKey, Me, dimensions, valueField, benchmarkField, riskFreeField, targetReturn, periodsPerYear)
End Function

Public Property Get rowCount() As Long
    rowCount = m_RowCount
End Property

Public Property Get startDate() As Variant
    Dim r As Long, firstDate As Date, found As Boolean
    startDate = Null
    For r = 1 To m_RowCount
        If m_HasDate(r) Then
            If Not found Then firstDate = m_DateValues(r): found = True
            If m_DateValues(r) < firstDate Then firstDate = m_DateValues(r)
        End If
    Next r
    If found Then startDate = firstDate
End Property

Public Property Get endDate() As Variant
    Dim r As Long, lastDate As Date, found As Boolean
    endDate = Null
    For r = 1 To m_RowCount
        If m_HasDate(r) Then
            If Not found Then lastDate = m_DateValues(r): found = True
            If m_DateValues(r) > lastDate Then lastDate = m_DateValues(r)
        End If
    Next r
    If found Then endDate = lastDate
End Property

' A scalar metric operates on a single daily series, after dimension filtering/aggregation.
Public Sub ValidateSingleSeries()
    Dim r As Long, key As Variant, rows As Collection
    If Not m_Loaded Then Err.Raise 5, "clsTimeSeriesFrame.ValidateSingleSeries", "Frame is not loaded."
    For r = 1 To m_RowCount
        If Not m_HasDate(r) Then Err.Raise 5, "clsTimeSeriesFrame.ValidateSingleSeries", "Missing or invalid date at row " & r
    Next r
    For Each key In m_DateRows.Keys
        Set rows = m_DateRows(key)
        If rows.count > 1 Then Err.Raise 5, "clsTimeSeriesFrame.ValidateSingleSeries", "Multiple observations on " & CStr(key) & "; filter or aggregate the series first."
    Next key
End Sub

Public Function SliceByIndices(ByRef indices() As Long, ByVal sliceCount As Long) As clsTimeSeriesFrame
    Dim rows As New Collection, i As Long, selected As clsTimeSeriesFrame
    If sliceCount < 0 Then Err.Raise 5, "clsTimeSeriesFrame.SliceByIndices", "Slice count cannot be negative."
    If sliceCount > 0 Then
        If sliceCount > UBound(indices) - LBound(indices) + 1 Then Err.Raise 5, "clsTimeSeriesFrame.SliceByIndices", "Insufficient slice indices."
        For i = 0 To sliceCount - 1
            rows.Add indices(LBound(indices) + i)
        Next i
    End If
    Set selected = SelectRows(rows)
    Set SliceByIndices = selected.SortedByDate()
End Function

Public Property Get fieldCount() As Long
    fieldCount = m_FieldCount
End Property

Public Property Get DateField() As String
    DateField = m_DateField
End Property

Public Property Get HasDate(ByVal rowNumber As Long) As Boolean
    ValidateRow rowNumber
    HasDate = m_HasDate(rowNumber)
End Property

Public Property Get RowDate(ByVal rowNumber As Long) As Variant
    ValidateRow rowNumber
    If m_HasDate(rowNumber) Then RowDate = m_DateValues(rowNumber) Else RowDate = Null
End Property

Public Property Get ValueAt(ByVal rowNumber As Long, ByVal fieldName As String) As Variant
    ValueAt = m_Data.ValueAt(rowNumber, fieldName)
End Property

Public Function FieldIndex(ByVal fieldName As String) As Long
    FieldIndex = m_Data.FieldIndex(fieldName)
End Function

Public Function RowsAtDate(ByVal observationDate As Date) As Collection
    Dim key As String, rows As Collection, sourceRows As Collection, item As Variant
    ' Daily lookup groups timestamps by calendar date while retaining original values.
    key = DateKey(observationDate)
    Set rows = New Collection
    If m_DateRows.exists(key) Then
        Set sourceRows = m_DateRows(key)
        For Each item In sourceRows
            rows.Add CLng(item)
        Next item
    End If
    Set RowsAtDate = rows
End Function

Public Function FilterDateRange(ByVal startDate As Date, ByVal endDate As Date) As clsTimeSeriesFrame
    Dim result As clsTimeSeriesFrame, selected() As Variant
    Dim r As Long, c As Long, outRow As Long, firstDate As Date, lastDate As Date
    If dateValue(startDate) > dateValue(endDate) Then Err.Raise 5, "clsTimeSeriesFrame.FilterDateRange", "Start date is after end date."
    firstDate = dateValue(startDate)
    lastDate = dateValue(endDate)
    ReDim selected(1 To m_RowCount + 1, 1 To m_FieldCount)
    For c = 1 To m_FieldCount
        selected(1, c) = m_Data.fieldName(c)
    Next c
    outRow = 1
    For r = 1 To m_RowCount
        If m_HasDate(r) Then
            If dateValue(m_DateValues(r)) >= firstDate And dateValue(m_DateValues(r)) <= lastDate Then
                outRow = outRow + 1
                For c = 1 To m_FieldCount
                    selected(outRow, c) = m_Data.ValueAtColumn(r, c)
                Next c
            End If
        End If
    Next r
    TrimTableRows selected, outRow, m_FieldCount
    Set result = New clsTimeSeriesFrame
    result.Load selected, m_DateField
    Set FilterDateRange = result
End Function

Public Function GroupRows(ByVal fieldNames As Variant) As Object
    Set GroupRows = m_Data.GroupRows(fieldNames)
End Function

Public Function UniqueRowIndex(ByVal fieldNames As Variant) As Object
    Set UniqueRowIndex = m_Data.UniqueRowIndex(fieldNames)
End Function

' Explicit aggregation policy for same-date records: flows SUM or snapshots LAST.
' Any missing member makes the group value Null; no implicit zero/fill.
Public Function AggregateByDate(ByVal dimensions As Variant, ByVal valueField As String, ByVal method As String) As clsTimeSeriesFrame
    Dim ordered As clsTimeSeriesFrame, daily As clsTimeSeriesFrame, result As clsTimeSeriesFrame
    Dim table As Variant, fields() As Variant, groups As Object, members As Collection
    Dim output() As Variant, key As Variant, item As Variant, Value As Variant
    Dim r As Long, c As Long, n As Long, dimCount As Long, outputRow As Long
    Dim valueColumn As Long, total As Double, missing As Boolean
    method = UCase$(Trim$(method))
    If method <> "SUM" And method <> "LAST" Then Err.Raise 5, "clsTimeSeriesFrame.AggregateByDate", "Choose SUM or LAST."
    If Not IsArray(dimensions) Then Err.Raise 5, "clsTimeSeriesFrame.AggregateByDate", "Dimensions must be an array."
    dimCount = UBound(dimensions) - LBound(dimensions) + 1
    valueColumn = FieldIndex(valueField)
    Set ordered = SortedByDate()
    table = ordered.ToTable()
    n = ordered.rowCount
    For r = 1 To n
        If Not ordered.HasDate(r) Then Err.Raise 5, "clsTimeSeriesFrame.AggregateByDate", "Invalid date at sorted row " & r
        table(r + 1, FieldIndex(m_DateField)) = dateValue(CDate(ordered.RowDate(r)))
    Next r
    Set daily = New clsTimeSeriesFrame
    daily.Load table, m_DateField
    ReDim fields(0 To dimCount)
    fields(0) = m_DateField
    For c = 1 To dimCount
        fields(c) = CStr(dimensions(LBound(dimensions) + c - 1))
        If UCase$(Trim$(CStr(fields(c)))) = m_DateField Then Err.Raise 5, "clsTimeSeriesFrame.AggregateByDate", "Date is already a grouping field."
        If FieldIndex(CStr(fields(c))) = valueColumn Then Err.Raise 5, "clsTimeSeriesFrame.AggregateByDate", "The measure cannot also be a dimension."
    Next c
    If valueColumn = FieldIndex(m_DateField) Then Err.Raise 5, "clsTimeSeriesFrame.AggregateByDate", "Date cannot be a measure."
    Set groups = daily.GroupRows(fields)
    ReDim output(1 To groups.count + 1, 1 To dimCount + 2)
    For c = 0 To dimCount: output(1, c + 1) = fields(c): Next c
    output(1, dimCount + 2) = UCase$(Trim$(valueField))
    outputRow = 1
    For Each key In groups.Keys
        outputRow = outputRow + 1
        Set members = groups(key)
        For c = 0 To dimCount
            output(outputRow, c + 1) = daily.ValueAt(CLng(members(1)), CStr(fields(c)))
        Next c
        total = 0#: missing = False
        For Each item In members
            Value = table(CLng(item) + 1, valueColumn)
            If IsNull(Value) Or IsEmpty(Value) Then
                missing = True
            ElseIf IsError(Value) Then
                Err.Raise 13, "clsTimeSeriesFrame.AggregateByDate", "Error measure: " & valueField
            ElseIf Not IsNumeric(Value) Then
                Err.Raise 13, "clsTimeSeriesFrame.AggregateByDate", "Non-numeric measure: " & valueField
            Else
                Select Case method
                    Case "SUM": total = total + CDbl(Value)
                    Case "LAST": total = CDbl(Value)
                End Select
            End If
        Next item
        If missing Then
            output(outputRow, dimCount + 2) = Null
        Else
            output(outputRow, dimCount + 2) = total
        End If
    Next key
    Set result = New clsTimeSeriesFrame
    result.Load output, m_DateField
    Set AggregateByDate = result
End Function

Public Function CompoundSeries(ByVal dimensions As Variant, ByVal returnField As String) As clsTimeSeriesFrame
    Dim ordered As clsTimeSeriesFrame, groups As Object, members As Collection
    Dim table As Variant, fields() As Variant, output() As Variant, result As clsTimeSeriesFrame
    Dim key As Variant, member As Variant, Value As Variant, seenDates As Object, dateToken As String
    Dim r As Long, c As Long, n As Long, dimCount As Long, outputRow As Long, returnColumn As Long
    Dim total As Double, firstDate As Date, lastDate As Date, missing As Boolean
    If Not IsArray(dimensions) Then Err.Raise 5, "clsTimeSeriesFrame.CompoundSeries", "Dimensions must be an array."
    dimCount = UBound(dimensions) - LBound(dimensions) + 1
    If dimCount < 1 Then Err.Raise 5, "clsTimeSeriesFrame.CompoundSeries", "At least one series dimension is required."
    returnColumn = FieldIndex(returnField)
    If returnColumn = FieldIndex(m_DateField) Then Err.Raise 5, "clsTimeSeriesFrame.CompoundSeries", "Date cannot be a return field."
    Set ordered = SortedByDate()
    table = ordered.ToTable()
    n = ordered.rowCount
    For r = 1 To n
        If Not ordered.HasDate(r) Then Err.Raise 5, "clsTimeSeriesFrame.CompoundSeries", "Invalid date at sorted row " & r
    Next r
    ReDim fields(0 To dimCount - 1)
    For c = 0 To dimCount - 1
        fields(c) = CStr(dimensions(LBound(dimensions) + c))
        If FieldIndex(CStr(fields(c))) = returnColumn Then Err.Raise 5, "clsTimeSeriesFrame.CompoundSeries", "The return field cannot be a dimension."
    Next c
    Set groups = ordered.GroupRows(fields)
    ReDim output(1 To groups.count + 1, 1 To dimCount + 4)
    For c = 0 To dimCount - 1: output(1, c + 1) = fields(c): Next c
    output(1, dimCount + 1) = "START_DATE"
    output(1, dimCount + 2) = "END_DATE"
    output(1, dimCount + 3) = "OBSERVATIONS"
    output(1, dimCount + 4) = UCase$(Trim$(returnField)) & "_COMPOUNDED"
    outputRow = 1
    For Each key In groups.Keys
        outputRow = outputRow + 1
        Set members = groups(key)
        Set seenDates = CreateObject("Scripting.Dictionary")
        seenDates.CompareMode = 0
        total = 1#: missing = False
        firstDate = ordered.RowDate(CLng(members(1)))
        lastDate = ordered.RowDate(CLng(members(members.count)))
        For c = 0 To dimCount - 1
            output(outputRow, c + 1) = ordered.ValueAt(CLng(members(1)), CStr(fields(c)))
        Next c
        For Each member In members
            r = CLng(member)
            dateToken = CStr(CDbl(dateValue(CDate(ordered.RowDate(r)))))
            If seenDates.exists(dateToken) Then Err.Raise 5, "clsTimeSeriesFrame.CompoundSeries", "Duplicate return observation for one date and dimension key."
            seenDates.Add dateToken, True
            Value = table(r + 1, returnColumn)
            If IsNull(Value) Or IsEmpty(Value) Then
                missing = True
            ElseIf IsError(Value) Then
                Err.Raise 13, "clsTimeSeriesFrame.CompoundSeries", "Error return: " & returnField
            ElseIf Not IsNumeric(Value) Then
                Err.Raise 13, "clsTimeSeriesFrame.CompoundSeries", "Non-numeric return: " & returnField
            ElseIf CDbl(Value) <= -1# Then
                Err.Raise 5, "clsTimeSeriesFrame.CompoundSeries", "Return must be greater than -100%."
            Else
                total = total * (1# + CDbl(Value))
            End If
        Next member
        output(outputRow, dimCount + 1) = firstDate
        output(outputRow, dimCount + 2) = lastDate
        output(outputRow, dimCount + 3) = members.count
        If missing Then output(outputRow, dimCount + 4) = Null Else output(outputRow, dimCount + 4) = total - 1#
    Next key
    Set result = New clsTimeSeriesFrame
    result.Load output, "START_DATE"
    Set CompoundSeries = result
End Function

Public Function SortedByDate() As clsTimeSeriesFrame
    Dim order() As Long, sorted() As Variant, result As clsTimeSeriesFrame
    Dim r As Long, c As Long, outRow As Long
    ReDim order(1 To MaxLong(1, m_RowCount))
    For r = 1 To m_RowCount
        order(r) = r
    Next r
    If m_RowCount > 1 Then QuickSortRows order, 1, m_RowCount

    ReDim sorted(1 To m_RowCount + 1, 1 To m_FieldCount)
    For c = 1 To m_FieldCount
        sorted(1, c) = m_Data.fieldName(c)
    Next c
    For outRow = 1 To m_RowCount
        r = order(outRow)
        For c = 1 To m_FieldCount
            sorted(outRow + 1, c) = m_Data.ValueAtColumn(r, c)
        Next c
    Next outRow
    Set result = New clsTimeSeriesFrame
    result.Load sorted, m_DateField
    Set SortedByDate = result
End Function

Private Sub AddDateRow(ByVal rowNumber As Long)
    Dim key As String, rows As Collection
    key = DateKey(m_DateValues(rowNumber))
    If Not m_DateRows.exists(key) Then
        Set rows = New Collection
        m_DateRows.Add key, rows
    End If
    m_DateRows(key).Add rowNumber
End Sub

Private Sub ValidateRow(ByVal rowNumber As Long)
    If rowNumber < 1 Or rowNumber > m_RowCount Then Err.Raise 9, "clsTimeSeriesFrame", "Row index is out of range."
End Sub

Private Function DateKey(ByVal Value As Date) As String
    DateKey = Format$(dateValue(Value), "yyyy-mm-dd")
End Function

Private Function MaxLong(ByVal leftValue As Long, ByVal rightValue As Long) As Long
    If leftValue > rightValue Then MaxLong = leftValue Else MaxLong = rightValue
End Function

Private Sub QuickSortRows(ByRef order() As Long, ByVal first As Long, ByVal last As Long)
    Dim low As Long, high As Long, pivot As Long, temp As Long
    low = first: high = last: pivot = order((first + last) \ 2)
    Do While low <= high
        Do While CompareRows(order(low), pivot) < 0
            low = low + 1
        Loop
        Do While CompareRows(order(high), pivot) > 0
            high = high - 1
        Loop
        If low <= high Then
            temp = order(low): order(low) = order(high): order(high) = temp
            low = low + 1: high = high - 1
        End If
    Loop
    If first < high Then QuickSortRows order, first, high
    If low < last Then QuickSortRows order, low, last
End Sub

Private Function CompareRows(ByVal leftRow As Long, ByVal rightRow As Long) As Long
    If m_HasDate(leftRow) And Not m_HasDate(rightRow) Then
        CompareRows = -1
    ElseIf Not m_HasDate(leftRow) And m_HasDate(rightRow) Then
        CompareRows = 1
    ElseIf m_HasDate(leftRow) And m_HasDate(rightRow) Then
        If m_DateValues(leftRow) < m_DateValues(rightRow) Then
            CompareRows = -1
        ElseIf m_DateValues(leftRow) > m_DateValues(rightRow) Then
            CompareRows = 1
        ElseIf leftRow < rightRow Then
            CompareRows = -1
        ElseIf leftRow > rightRow Then
            CompareRows = 1
        End If
    ElseIf leftRow < rightRow Then
        CompareRows = -1
    ElseIf leftRow > rightRow Then
        CompareRows = 1
    End If
End Function

Private Sub TrimTableRows(ByRef table As Variant, ByVal rowCount As Long, ByVal fieldCount As Long)
    Dim trimmed() As Variant, r As Long, c As Long
    ReDim trimmed(1 To rowCount, 1 To fieldCount)
    For r = 1 To rowCount
        For c = 1 To fieldCount
            trimmed(r, c) = table(r, c)
        Next c
    Next r
    table = trimmed
End Sub
