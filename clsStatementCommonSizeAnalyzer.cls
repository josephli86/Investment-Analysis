VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStatementCommonSizeAnalyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Load statement layout rows from Config\MEASURES_DEF.xlsx [StatementLayout$]
Public Function LoadLayout(Optional ByVal stmtType As String = "IS", _
        Optional ByVal configPath As String = "") As clsDataTable
    Dim filePath As String, cleanType As String, sqlText As String
    Dim conn As Object, rs As Object, reader As clsRecordsetReader
    Dim rawTbl As clsDataTable, schema As clsDataSchema
    Dim n As Long, description As String

    On Error GoTo Failed
    If Len(Trim$(configPath)) = 0 Then
        filePath = ThisWorkbook.Path & "\Config\MEASURES_DEF.xlsx"
    Else
        filePath = configPath
    End If

    cleanType = UCase$(Trim$(stmtType))
    sqlText = "SELECT STMT_TYPE, LINE_ORDER, ACCT_CODE, ITEM_NAME, SOURCE_TYPE, COMMON_BASE_CODE, INDENT, IS_BOLD, CALC_YOY, ENABLED " & _
              "FROM [StatementLayout$] WHERE UCASE(ENABLED) = 'Y'"
    If Len(cleanType) > 0 Then
        sqlText = sqlText & " AND UCASE(STMT_TYPE) = '" & Replace(cleanType, "'", "''") & "'"
    End If
    sqlText = sqlText & " ORDER BY LINE_ORDER ASC"

    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Mode=Read;Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
    Set rs = conn.Execute(sqlText)
    If rs.EOF Then Err.Raise 5, "clsStatementCommonSizeAnalyzer.LoadLayout", "No enabled rows in StatementLayout$ for STMT_TYPE=" & cleanType

    Set reader = New clsRecordsetReader
    Set rawTbl = reader.Read(rs)
    rs.Close: conn.Close

    Set schema = New clsDataSchema
    schema.Initialize "STATEMENT_LAYOUT", "1", "STMT_TYPE|ACCT_CODE"
    schema.AddField "STMT_TYPE", "TEXT", False
    schema.AddField "LINE_ORDER", "LONG", False
    schema.AddField "ACCT_CODE", "TEXT", False
    schema.AddField "ITEM_NAME", "TEXT", False
    schema.AddField "SOURCE_TYPE", "TEXT", False
    schema.AddField "COMMON_BASE_CODE", "TEXT", True
    schema.AddField "INDENT", "LONG", False
    schema.AddField "IS_BOLD", "TEXT", False
    schema.AddField "CALC_YOY", "TEXT", False
    schema.AddField "ENABLED", "TEXT", False
    schema.Freeze

    Set LoadLayout = schema.Normalize(rawTbl)
    Exit Function

Failed:
    n = Err.number: description = Err.description
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State <> 0 Then rs.Close
    If Not conn Is Nothing Then If conn.State <> 0 Then conn.Close
    On Error GoTo 0
    Err.Raise n, "clsStatementCommonSizeAnalyzer.LoadLayout", description
End Function

' Build canonical enriched long table for a single company across multiple periods
' Columns: COMPANY_CODE | REPORT_PERIOD | STMT_TYPE | LINE_ORDER | ACCT_CODE | ITEM_NAME |
'          AMOUNT | PRIOR_AMOUNT | COMMON_SIZE | YOY_DIFF | YOY_GROWTH | VERSION_KEY | STATUS | DIAGNOSTIC
Public Function BuildSingleCompanyLong(ByVal stmt As clsFinancialStatement, _
        ByVal layoutTable As clsDataTable, _
        ByVal periods As Variant, _
        Optional ByVal ifAdjusted As Long = 2, _
        Optional ByVal engine As clsMeasureEngine = Nothing) As clsDataTable
    Dim activeEngine As clsMeasureEngine, periodArr As Variant
    Dim pIdx As Long, pCount As Long, r As Long, itemCount As Long, outRow As Long
    Dim periodEnd As String, priorPeriod As String
    Dim curVKey As String, curPubl As String, priVKey As String, priPubl As String, priAdj As Long
    Dim acctCode As String, srcType As String, baseCode As String, calcYoy As Boolean
    Dim curVal As Variant, priVal As Variant, baseVal As Variant
    Dim statusText As String, diagText As String, priStatus As String, priDiag As String
    Dim valueIndex As Object, lookupKey As String
    Dim rows() As Variant, resTable As clsDataTable

    If stmt Is Nothing Then Err.Raise 5, "clsStatementCommonSizeAnalyzer", "Financial statement is required."
    If layoutTable Is Nothing Or Not layoutTable.IsLoaded Then Err.Raise 5, "clsStatementCommonSizeAnalyzer", "Layout table is required."
    If ifAdjusted <> 1 And ifAdjusted <> 2 Then Err.Raise 5, "clsStatementCommonSizeAnalyzer", "IfAdjusted must be 1 or 2."
    If engine Is Nothing Then Set activeEngine = MeasureMgr() Else Set activeEngine = engine

    periodArr = NormalizePeriodArray(periods)
    pCount = UBound(periodArr) - LBound(periodArr) + 1
    itemCount = layoutTable.rowCount

    ReDim rows(1 To pCount * itemCount + 1, 1 To 14)
    rows(1, 1) = "COMPANY_CODE": rows(1, 2) = "REPORT_PERIOD": rows(1, 3) = "STMT_TYPE"
    rows(1, 4) = "LINE_ORDER": rows(1, 5) = "ACCT_CODE": rows(1, 6) = "ITEM_NAME"
    rows(1, 7) = "AMOUNT": rows(1, 8) = "PRIOR_AMOUNT": rows(1, 9) = "COMMON_SIZE"
    rows(1, 10) = "YOY_DIFF": rows(1, 11) = "YOY_GROWTH": rows(1, 12) = "VERSION_KEY"
    rows(1, 13) = "STATUS": rows(1, 14) = "DIAGNOSTIC"

    Set valueIndex = CreateObject("Scripting.Dictionary")
    valueIndex.CompareMode = 1
    outRow = 1

    ' Pass 1: evaluate raw accounts & derived metrics for current and restated comparable prior periods
    For pIdx = LBound(periodArr) To UBound(periodArr)
        periodEnd = FinancialPeriodEnd(CStr(periodArr(pIdx)))
        priorPeriod = Format$(DateAdd("yyyy", -1, CDate(periodEnd)), "yyyy-mm-dd")
        curVKey = stmt.SelectedVersionKey(periodEnd, ifAdjusted)

        If Len(curVKey) > 0 Then
            curPubl = stmt.SelectedPublicationDate(periodEnd, ifAdjusted)
            priVKey = stmt.SelectedComparableVersionKey(priorPeriod, curVKey)
            If Len(priVKey) > 0 Then
                priAdj = 1
                priPubl = curPubl
            Else
                priVKey = stmt.SelectedVersionKey(priorPeriod, ifAdjusted)
                priAdj = ifAdjusted
                priPubl = IIf(Len(priVKey) > 0, stmt.SelectedPublicationDate(priorPeriod, ifAdjusted), "")
            End If
        Else
            curPubl = "": priVKey = "": priPubl = "": priAdj = 1
        End If

        For r = 1 To itemCount
            outRow = outRow + 1
            acctCode = UCase$(Trim$(CStr(layoutTable.ValueAt(r, "ACCT_CODE"))))
            srcType = UCase$(Trim$(CStr(layoutTable.ValueAt(r, "SOURCE_TYPE"))))
            calcYoy = (UCase$(Trim$(CStr(layoutTable.ValueAt(r, "CALC_YOY")))) = "Y")

            rows(outRow, 1) = stmt.companyCode
            rows(outRow, 2) = periodEnd
            rows(outRow, 3) = layoutTable.ValueAt(r, "STMT_TYPE")
            rows(outRow, 4) = layoutTable.ValueAt(r, "LINE_ORDER")
            rows(outRow, 5) = acctCode
            rows(outRow, 6) = layoutTable.ValueAt(r, "ITEM_NAME")
            rows(outRow, 12) = curVKey

            If Len(curVKey) = 0 Then
                rows(outRow, 7) = Null: rows(outRow, 8) = Null
                rows(outRow, 13) = "MISSING_REPORT"
                rows(outRow, 14) = "No selected report version for " & periodEnd
            Else
                curVal = EvaluateItemAtVersion(stmt, activeEngine, periodEnd, ifAdjusted, curPubl, curVKey, acctCode, srcType, statusText, diagText)
                rows(outRow, 7) = curVal
                rows(outRow, 13) = statusText
                rows(outRow, 14) = diagText

                If Not IsNull(curVal) Then
                    valueIndex(periodEnd & "|" & acctCode) = CDbl(curVal)
                End If

                If calcYoy And Len(priVKey) > 0 Then
                    priVal = EvaluateItemAtVersion(stmt, activeEngine, priorPeriod, priAdj, priPubl, priVKey, acctCode, srcType, priStatus, priDiag)
                    rows(outRow, 8) = priVal
                Else
                    rows(outRow, 8) = Null
                End If
            End If
        Next r
    Next pIdx

    ' Pass 2: vectorized common-size ratio & YoY growth derivation
    outRow = 1
    For pIdx = LBound(periodArr) To UBound(periodArr)
        periodEnd = FinancialPeriodEnd(CStr(periodArr(pIdx)))
        For r = 1 To itemCount
            outRow = outRow + 1
            curVal = rows(outRow, 7)
            priVal = rows(outRow, 8)
            baseCode = SafeUpperText(layoutTable.ValueAt(r, "COMMON_BASE_CODE"))

            rows(outRow, 9) = Null
            If Not IsNull(curVal) And Len(baseCode) > 0 Then
                lookupKey = periodEnd & "|" & baseCode
                If valueIndex.exists(lookupKey) Then
                    baseVal = valueIndex(lookupKey)
                    If Abs(CDbl(baseVal)) >= 0.0000001 Then
                        rows(outRow, 9) = CDbl(curVal) / CDbl(baseVal)
                    End If
                End If
            End If

            rows(outRow, 10) = Null
            rows(outRow, 11) = Null
            If Not IsNull(curVal) And Not IsNull(priVal) Then
                rows(outRow, 10) = CDbl(curVal) - CDbl(priVal)
                If Abs(CDbl(priVal)) >= 0.0000001 Then
                    rows(outRow, 11) = (CDbl(curVal) - CDbl(priVal)) / Abs(CDbl(priVal))
                End If
            End If
        Next r
    Next pIdx

    Set resTable = New clsDataTable
    resTable.Load rows
    Set BuildSingleCompanyLong = resTable
End Function

' Project single-company long table into multi-period comparison wide table
' displayMode: "BOTH" (Amount + CommonSize per period), "COMMON_SIZE", "AMOUNT", or "YOY_GROWTH"
Public Function PivotSingleCompanyMultiPeriod(ByVal longTable As clsDataTable, _
        ByVal layoutTable As clsDataTable, _
        ByVal periods As Variant, _
        Optional ByVal displayMode As String = "BOTH", _
        Optional ByRef outNotesTable As clsDataTable = Nothing) As clsDataTable
    Dim periodArr As Variant, pCount As Long, pIdx As Long, r As Long, itemCount As Long
    Dim modeKey As String, colsPerPeriod As Long, totalCols As Long, colBase As Long
    Dim pEnd As String, acctCode As String, key As String
    Dim amtIdx As Object, pctIdx As Object, yoyIdx As Object, diagIdx As Object
    Dim wide() As Variant, notes() As Variant, resTable As clsDataTable, resNotes As clsDataTable

    periodArr = NormalizePeriodArray(periods)
    pCount = UBound(periodArr) - LBound(periodArr) + 1
    itemCount = layoutTable.rowCount
    modeKey = UCase$(Trim$(displayMode))

    Select Case modeKey
        Case "BOTH": colsPerPeriod = 2
        Case "COMMON_SIZE", "AMOUNT", "YOY_GROWTH": colsPerPeriod = 1
        Case Else: Err.Raise 5, "clsStatementCommonSizeAnalyzer.PivotSingleCompanyMultiPeriod", "Invalid displayMode: " & displayMode
    End Select

    Set amtIdx = CreateObject("Scripting.Dictionary"): amtIdx.CompareMode = 1
    Set pctIdx = CreateObject("Scripting.Dictionary"): pctIdx.CompareMode = 1
    Set yoyIdx = CreateObject("Scripting.Dictionary"): yoyIdx.CompareMode = 1
    Set diagIdx = CreateObject("Scripting.Dictionary"): diagIdx.CompareMode = 1

    For r = 1 To longTable.rowCount
        key = CStr(longTable.ValueAt(r, "REPORT_PERIOD")) & "|" & UCase$(Trim$(CStr(longTable.ValueAt(r, "ACCT_CODE"))))
        amtIdx(key) = longTable.ValueAt(r, "AMOUNT")
        pctIdx(key) = longTable.ValueAt(r, "COMMON_SIZE")
        yoyIdx(key) = longTable.ValueAt(r, "YOY_GROWTH")
        If CStr(longTable.ValueAt(r, "STATUS")) <> "OK" Then
            diagIdx(key) = "[" & CStr(longTable.ValueAt(r, "STATUS")) & "] " & CStr(longTable.ValueAt(r, "DIAGNOSTIC"))
        End If
    Next r

    totalCols = 4 + pCount * colsPerPeriod
    ReDim wide(1 To itemCount + 1, 1 To totalCols)
    ReDim notes(1 To itemCount + 1, 1 To totalCols)

    wide(1, 1) = "LINE_ORDER": wide(1, 2) = "ACCT_CODE": wide(1, 3) = "ITEM_NAME": wide(1, 4) = "STYLE_FLAG"
    For pIdx = 0 To pCount - 1
        pEnd = FinancialPeriodEnd(CStr(periodArr(LBound(periodArr) + pIdx)))
        colBase = 5 + pIdx * colsPerPeriod
        If modeKey = "BOTH" Then
            wide(1, colBase) = pEnd & "_AMT"
            wide(1, colBase + 1) = pEnd & "_PCT"
        ElseIf modeKey = "COMMON_SIZE" Then
            wide(1, colBase) = pEnd & "_PCT"
        ElseIf modeKey = "AMOUNT" Then
            wide(1, colBase) = pEnd & "_AMT"
        Else
            wide(1, colBase) = pEnd & "_YOY"
        End If
    Next pIdx
    For r = 1 To totalCols: notes(1, r) = wide(1, r): Next r

    For r = 1 To itemCount
        acctCode = UCase$(Trim$(CStr(layoutTable.ValueAt(r, "ACCT_CODE"))))
        wide(r + 1, 1) = layoutTable.ValueAt(r, "LINE_ORDER")
        wide(r + 1, 2) = acctCode
        wide(r + 1, 3) = layoutTable.ValueAt(r, "ITEM_NAME")
        wide(r + 1, 4) = CStr(layoutTable.ValueAt(r, "INDENT")) & "|" & UCase$(Trim$(CStr(layoutTable.ValueAt(r, "IS_BOLD"))))

        For pIdx = 0 To pCount - 1
            pEnd = FinancialPeriodEnd(CStr(periodArr(LBound(periodArr) + pIdx)))
            key = pEnd & "|" & acctCode
            colBase = 5 + pIdx * colsPerPeriod

            If modeKey = "BOTH" Then
                wide(r + 1, colBase) = IIf(amtIdx.exists(key), amtIdx(key), Null)
                wide(r + 1, colBase + 1) = IIf(pctIdx.exists(key), pctIdx(key), Null)
                If diagIdx.exists(key) Then notes(r + 1, colBase) = diagIdx(key)
            ElseIf modeKey = "COMMON_SIZE" Then
                wide(r + 1, colBase) = IIf(pctIdx.exists(key), pctIdx(key), Null)
                If diagIdx.exists(key) Then notes(r + 1, colBase) = diagIdx(key)
            ElseIf modeKey = "AMOUNT" Then
                wide(r + 1, colBase) = IIf(amtIdx.exists(key), amtIdx(key), Null)
                If diagIdx.exists(key) Then notes(r + 1, colBase) = diagIdx(key)
            Else
                wide(r + 1, colBase) = IIf(yoyIdx.exists(key), yoyIdx(key), Null)
                If diagIdx.exists(key) Then notes(r + 1, colBase) = diagIdx(key)
            End If
        Next pIdx
    Next r

    Set resTable = New clsDataTable: resTable.Load wide
    Set resNotes = New clsDataTable: resNotes.Load notes
    Set outNotesTable = resNotes
    Set PivotSingleCompanyMultiPeriod = resTable
End Function

' Build peer same-period comparison wide table across multiple companies
' sliceField: "COMMON_SIZE" (default), "AMOUNT", or "YOY_GROWTH"
Public Function BuildPeerSamePeriodComparison(ByVal statements As Collection, _
        ByVal layoutTable As clsDataTable, _
        ByVal reportPeriod As String, _
        Optional ByVal sliceField As String = "COMMON_SIZE", _
        Optional ByVal ifAdjusted As Long = 2, _
        Optional ByVal engine As clsMeasureEngine = Nothing, _
        Optional ByRef outCombinedLong As clsDataTable = Nothing, _
        Optional ByRef outNotesTable As clsDataTable = Nothing, _
        Optional ByVal companyLabels As Object = Nothing) As clsDataTable
    Dim activeEngine As clsMeasureEngine, periodEnd As String, metricCol As String
    Dim stmt As clsFinancialStatement, compLong As clsDataTable, longTables As New Collection
    Dim cCount As Long, cIdx As Long, r As Long, itemCount As Long, totalCols As Long
    Dim valueIdx As Object, diagIdx As Object, key As String, compCode As String, acctCode As String
    Dim wide() As Variant, notes() As Variant, validVals() As Double, vCount As Long, cellVal As Variant
    Dim resTable As clsDataTable, resNotes As clsDataTable

    If statements Is Nothing Or statements.count = 0 Then
        Err.Raise 5, "clsStatementCommonSizeAnalyzer.BuildPeerSamePeriodComparison", "At least one statement is required."
    End If
    If engine Is Nothing Then Set activeEngine = MeasureMgr() Else Set activeEngine = engine

    periodEnd = FinancialPeriodEnd(reportPeriod)
    metricCol = UCase$(Trim$(sliceField))
    If metricCol <> "COMMON_SIZE" And metricCol <> "AMOUNT" And metricCol <> "YOY_GROWTH" Then
        Err.Raise 5, "clsStatementCommonSizeAnalyzer.BuildPeerSamePeriodComparison", "sliceField must be COMMON_SIZE, AMOUNT, or YOY_GROWTH."
    End If

    Set valueIdx = CreateObject("Scripting.Dictionary"): valueIdx.CompareMode = 1
    Set diagIdx = CreateObject("Scripting.Dictionary"): diagIdx.CompareMode = 1

    For Each stmt In statements
        Set compLong = BuildSingleCompanyLong(stmt, layoutTable, Array(periodEnd), ifAdjusted, activeEngine)
        longTables.Add compLong
        For r = 1 To compLong.rowCount
            key = UCase$(Trim$(CStr(compLong.ValueAt(r, "COMPANY_CODE")))) & "|" & _
                  UCase$(Trim$(CStr(compLong.ValueAt(r, "ACCT_CODE"))))
            valueIdx(key) = compLong.ValueAt(r, metricCol)
            If CStr(compLong.ValueAt(r, "STATUS")) <> "OK" Then
                diagIdx(key) = "[" & CStr(compLong.ValueAt(r, "STATUS")) & "] " & CStr(compLong.ValueAt(r, "DIAGNOSTIC"))
            End If
        Next r
    Next stmt

    Set outCombinedLong = MergeLongTables(longTables)

    cCount = statements.count
    itemCount = layoutTable.rowCount
    totalCols = 6 + cCount
    ReDim wide(1 To itemCount + 1, 1 To totalCols)
    ReDim notes(1 To itemCount + 1, 1 To totalCols)

    wide(1, 1) = "LINE_ORDER": wide(1, 2) = "ACCT_CODE": wide(1, 3) = "ITEM_NAME": wide(1, 4) = "STYLE_FLAG"
    For cIdx = 1 To cCount
        Set stmt = statements(cIdx)
        compCode = UCase$(Trim$(stmt.companyCode))
        If Not companyLabels Is Nothing Then
            If companyLabels.exists(compCode) Then
                wide(1, 4 + cIdx) = CStr(companyLabels(compCode))
            Else
                wide(1, 4 + cIdx) = compCode
            End If
        Else
            wide(1, 4 + cIdx) = compCode
        End If
    Next cIdx
    wide(1, 5 + cCount) = "PEER_MEAN"
    wide(1, 6 + cCount) = "PEER_MEDIAN"
    For r = 1 To totalCols: notes(1, r) = wide(1, r): Next r

    For r = 1 To itemCount
        acctCode = UCase$(Trim$(CStr(layoutTable.ValueAt(r, "ACCT_CODE"))))
        wide(r + 1, 1) = layoutTable.ValueAt(r, "LINE_ORDER")
        wide(r + 1, 2) = acctCode
        wide(r + 1, 3) = layoutTable.ValueAt(r, "ITEM_NAME")
        wide(r + 1, 4) = CStr(layoutTable.ValueAt(r, "INDENT")) & "|" & UCase$(Trim$(CStr(layoutTable.ValueAt(r, "IS_BOLD"))))

        vCount = 0
        ReDim validVals(1 To cCount)
        For cIdx = 1 To cCount
            Set stmt = statements(cIdx)
            compCode = UCase$(Trim$(stmt.companyCode))
            key = compCode & "|" & acctCode
            cellVal = IIf(valueIdx.exists(key), valueIdx(key), Null)
            wide(r + 1, 4 + cIdx) = cellVal
            If diagIdx.exists(key) Then notes(r + 1, 4 + cIdx) = diagIdx(key)

            If Not IsNull(cellVal) And IsNumeric(cellVal) Then
                vCount = vCount + 1
                validVals(vCount) = CDbl(cellVal)
            End If
        Next cIdx

        If vCount > 0 Then
            ReDim Preserve validVals(1 To vCount)
            SortDoubleArray validVals
            wide(r + 1, 5 + cCount) = MeanArray(validVals)
            wide(r + 1, 6 + cCount) = MedianSorted(validVals)
        Else
            wide(r + 1, 5 + cCount) = Null
            wide(r + 1, 6 + cCount) = Null
        End If
    Next r

    Set resTable = New clsDataTable: resTable.Load wide
    Set resNotes = New clsDataTable: resNotes.Load notes
    Set outNotesTable = resNotes
    Set BuildPeerSamePeriodComparison = resTable
End Function

' ==============================================================================
' Internal Evaluation & Statistical Helpers
' ==============================================================================

Private Function EvaluateItemAtVersion(ByVal stmt As clsFinancialStatement, _
        ByVal engine As clsMeasureEngine, _
        ByVal periodEnd As String, _
        ByVal ifAdj As Long, _
        ByVal publDate As String, _
        ByVal versionKey As String, _
        ByVal acctCode As String, _
        ByVal srcType As String, _
        ByRef outStatus As String, _
        ByRef outDiag As String) As Variant
    Dim provider As clsStatementProvider, ctx As clsEvalContext
    Dim v As Variant, rawAmt As Variant

    Set provider = New clsStatementProvider
    provider.Initialize stmt
    provider.PinReportVersion periodEnd, ifAdj, publDate
    Set ctx = New clsEvalContext
    ctx.RegisterProvider provider

    v = engine.EvaluateContext(ctx, periodEnd, acctCode, CInt(ifAdj))
    If IsError(v) Then
        EvaluateItemAtVersion = Null
        outStatus = "ERROR"
        outDiag = engine.LastError
        If Len(outDiag) = 0 Then outDiag = "Evaluation failed: " & acctCode & " / " & versionKey
        Exit Function
    End If

    EvaluateItemAtVersion = CDbl(v)
    If acctCode = "TAX_RATE" Then
        If stmt.TryGetVersionRawValue(versionKey, "TAX_RATE", rawAmt) Then
            outStatus = "OK": outDiag = ""
        Else
            outStatus = "DEFAULT_TAX_RATE"
            outDiag = "Missing TAX_RATE; default 0.25 applied / " & versionKey
        End If
    ElseIf srcType = "RAW" Then
        If stmt.TryGetVersionRawValue(versionKey, acctCode, rawAmt) Then
            outStatus = "OK": outDiag = ""
        Else
            outStatus = "DEFAULT_ZERO"
            outDiag = "Missing or NULL raw account treated as 0 / " & versionKey
        End If
    Else
        outStatus = "OK": outDiag = ""
    End If
End Function

Private Function MergeLongTables(ByVal tables As Collection) As clsDataTable
    Dim tbl As clsDataTable, totalRows As Long, colCount As Long
    Dim merged() As Variant, outRow As Long, r As Long, c As Long, res As clsDataTable
    Set tbl = tables(1)
    colCount = tbl.fieldCount
    For Each tbl In tables
        totalRows = totalRows + tbl.rowCount
    Next tbl
    ReDim merged(1 To totalRows + 1, 1 To colCount)
    Set tbl = tables(1)
    For c = 1 To colCount
        merged(1, c) = tbl.fieldName(c)
    Next c
    outRow = 1
    For Each tbl In tables
        For r = 1 To tbl.rowCount
            outRow = outRow + 1
            For c = 1 To colCount
                merged(outRow, c) = tbl.ValueAtColumn(r, c)
            Next c
        Next r
    Next tbl
    Set res = New clsDataTable
    res.Load merged
    Set MergeLongTables = res
End Function

Private Function NormalizePeriodArray(ByVal periods As Variant) As Variant
    Dim arr() As Variant
    If IsArray(periods) Then
        NormalizePeriodArray = periods
    Else
        ReDim arr(1 To 1)
        arr(1) = periods
        NormalizePeriodArray = arr
    End If
End Function

Private Function SafeUpperText(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    SafeUpperText = UCase$(Trim$(CStr(v)))
End Function

Private Sub SortDoubleArray(ByRef arr() As Double)
    Dim i As Long, j As Long, tmp As Double
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
        Next j
    Next i
End Sub

Private Function MeanArray(ByRef arr() As Double) As Double
    Dim i As Long, n As Long, s As Double
    n = UBound(arr) - LBound(arr) + 1
    For i = LBound(arr) To UBound(arr)
        s = s + arr(i)
    Next i
    MeanArray = s / CDbl(n)
End Function

Private Function MedianSorted(ByRef sorted() As Double) As Double
    Dim n As Long
    n = UBound(sorted) - LBound(sorted) + 1
    If n Mod 2 = 1 Then
        MedianSorted = sorted(LBound(sorted) + n \ 2)
    Else
        MedianSorted = (sorted(LBound(sorted) + n \ 2 - 1) + sorted(LBound(sorted) + n \ 2)) / 2#
    End If
End Function
