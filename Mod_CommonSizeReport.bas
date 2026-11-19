Attribute VB_Name = "Mod_CommonSizeReport"
Option Explicit

' Entry 1: Single-company multi-period common-size & YoY comparison
Public Sub ExportSingleCompanyCommonSize()
    Dim companyCode As String, periodsText As String, stmtType As String, cutoffText As String
    Dim rawParts() As String, periods() As String, i As Long, maxPeriod As String
    Dim finSvc As clsFinancialService, stmt As clsFinancialStatement
    Dim ctx As clsAnalysisContext, app As clsFinancialAnalysisApplication, res As clsAnalysisResult
    Dim wb As Workbook, wsComp As Worksheet, wsLong As Worksheet, wsLineage As Worksheet
    Dim renderer As clsResultTableRenderer

    On Error GoTo Failed
    companyCode = Trim$(InputBox("Enter numeric CompanyCode:", "Single Company Common-Size", "1833"))
    If Len(companyCode) = 0 Then Exit Sub

    periodsText = Trim$(InputBox("Enter fiscal periods (comma-separated, e.g. 2023,2024,2025):", "Single Company Common-Size", "2023,2024,2025"))
    If Len(periodsText) = 0 Then Exit Sub

    stmtType = UCase$(Trim$(InputBox("Statement type (IS or BS):", "Single Company Common-Size", "IS")))
    If Len(stmtType) = 0 Then Exit Sub

    cutoffText = Trim$(InputBox("As-of cutoff date (yyyy-mm-dd):", "Single Company Common-Size", Format$(Date, "yyyy-mm-dd")))
    If Len(cutoffText) = 0 Then Exit Sub

    rawParts = Split(periodsText, ",")
    ReDim periods(0 To UBound(rawParts))
    Dim minPeriod As String: minPeriod = "9999-12-31"
    For i = 0 To UBound(rawParts)
        Dim pClean As String: pClean = Trim$(rawParts(i))
        If Len(pClean) > 0 Then
            periods(i) = FinancialPeriodEnd(pClean)
            If periods(i) > maxPeriod Then maxPeriod = periods(i)
            If periods(i) < minPeriod Then minPeriod = periods(i)
        End If
    Next i

    Set finSvc = New clsFinancialService
    Set stmt = finSvc.GetStatement(companyCode, Format$(DateAdd("yyyy", -1, CDate(minPeriod)), "yyyy-mm-dd"), maxPeriod, 1, cutoffText, "", 2)

    Set ctx = New clsAnalysisContext
    ctx.Initialize CDate(cutoffText), CDate(cutoffText), "CNY", "COMMON_SIZE_V1"

    Set app = New clsFinancialAnalysisApplication
    Set res = app.RunCommonSizeSingleCompany(stmt, ctx, periods, stmtType, "BOTH", 2)

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    Set wsComp = wb.Worksheets(1)
    wsComp.name = "MultiPeriod_CommonSize"
    RenderStyledComparisonSheet res, "MULTI_PERIOD_COMPARISON", "MULTI_PERIOD_NOTES", "BOTH", wsComp

    Set wsLong = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
    wsLong.name = "Detail_Long"
    RenderStyledLongSheet res, "DETAIL_LONG", wsLong

    Set wsLineage = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
    wsLineage.name = "Lineage"
    Set renderer = New clsResultTableRenderer
    renderer.Render res, "LINEAGE", wsLineage

    wsComp.Activate
    MsgBox "Single-company multi-period common-size report exported (RunId=" & res.RunId & ").", vbInformation, "Common-Size Analysis"
    Exit Sub

Failed:
    MsgBox "Export failed: " & Err.description, vbExclamation, "Common-Size Analysis"
End Sub

' Entry 2: Multi-company same-period peer common-size comparison
' Accepts stock tickers (e.g. "600519,000858,000568") or numeric CompanyCodes
Public Sub ExportPeerCommonSizeComparison()
    Dim inputCodes As String, periodText As String, stmtType As String, sliceField As String, cutoffText As String
    Dim tokens() As String, i As Long, compCode As String, labelText As String
    Dim compCodes As New Collection, labelsDict As Object
    Dim finSvc As clsFinancialService, stmt As clsFinancialStatement, stmts As New Collection
    Dim ctx As clsAnalysisContext, app As clsFinancialAnalysisApplication, res As clsAnalysisResult
    Dim wb As Workbook, wsPeer As Worksheet, wsLong As Worksheet, wsLineage As Worksheet
    Dim renderer As clsResultTableRenderer, periodEnd As String

    On Error GoTo Failed
    inputCodes = Trim$(InputBox("Enter stock tickers or CompanyCodes (comma-separated):", "Peer Common-Size Comparison", "600519,000858,000568"))
    If Len(inputCodes) = 0 Then Exit Sub

    periodText = Trim$(InputBox("Report period (YYYY or YYYY-MM-DD):", "Peer Common-Size Comparison", "2025-12-31"))
    If Len(periodText) = 0 Then Exit Sub

    stmtType = UCase$(Trim$(InputBox("Statement type (IS or BS):", "Peer Common-Size Comparison", "IS")))
    If Len(stmtType) = 0 Then Exit Sub

    sliceField = UCase$(Trim$(InputBox("Compare metric (COMMON_SIZE, AMOUNT, or YOY_GROWTH):", "Peer Common-Size Comparison", "COMMON_SIZE")))
    If Len(sliceField) = 0 Then Exit Sub

    cutoffText = Trim$(InputBox("As-of cutoff date (yyyy-mm-dd):", "Peer Common-Size Comparison", Format$(Date, "yyyy-mm-dd")))
    If Len(cutoffText) = 0 Then Exit Sub

    periodEnd = FinancialPeriodEnd(periodText)
    Set labelsDict = CreateObject("Scripting.Dictionary")
    labelsDict.CompareMode = 1

    tokens = Split(inputCodes, ",")
    For i = 0 To UBound(tokens)
        If Len(Trim$(tokens(i))) > 0 Then
            ResolveTickerOrCompanyCode Trim$(tokens(i)), compCode, labelText
            If Not labelsDict.exists(compCode) Then
                labelsDict.Add compCode, labelText
                compCodes.Add compCode
            End If
        End If
    Next i

    Set finSvc = New clsFinancialService
    For i = 1 To compCodes.count
        compCode = CStr(compCodes(i))
        Set stmt = finSvc.GetStatement(compCode, Format$(DateAdd("yyyy", -1, CDate(periodEnd)), "yyyy-mm-dd"), periodEnd, 1, cutoffText, "", 2)
        stmts.Add stmt
    Next i

    Set ctx = New clsAnalysisContext
    ctx.Initialize CDate(cutoffText), CDate(cutoffText), "CNY", "COMMON_SIZE_PEER_V1"

    Set app = New clsFinancialAnalysisApplication
    Set res = app.RunCommonSizePeerComparison(stmts, ctx, periodEnd, stmtType, sliceField, 2, Nothing, labelsDict)

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    Set wsPeer = wb.Worksheets(1)
    wsPeer.name = "Peer_CommonSize"
    RenderStyledComparisonSheet res, "PEER_COMPARISON", "PEER_NOTES", sliceField, wsPeer

    Set wsLong = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
    wsLong.name = "Peer_Detail_Long"
    RenderStyledLongSheet res, "DETAIL_LONG", wsLong

    Set wsLineage = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
    wsLineage.name = "Lineage"
    Set renderer = New clsResultTableRenderer
    renderer.Render res, "LINEAGE", wsLineage

    wsPeer.Activate
    MsgBox "Peer common-size comparison exported (RunId=" & res.RunId & ").", vbInformation, "Peer Common-Size Comparison"
    Exit Sub

Failed:
    MsgBox "Peer comparison failed: " & Err.description, vbExclamation, "Peer Common-Size Comparison"
End Sub

' ==============================================================================
' Styled Excel Presentation Helpers
' ==============================================================================

Private Sub RenderStyledComparisonSheet(ByVal result As clsAnalysisResult, _
        ByVal dataTableName As String, _
        ByVal notesTableName As String, _
        ByVal sliceMode As String, _
        ByVal ws As Worksheet)
    Dim data As Variant, notes As Variant, rCount As Long, cCount As Long
    Dim r As Long, c As Long, styleParts() As String, indentLvl As Long, isBold As Boolean
    Dim colHeader As String, noteText As String

    data = result.table(dataTableName).ToTable()
    notes = result.table(notesTableName).ToTable()
    rCount = UBound(data, 1)
    cCount = UBound(data, 2)

    ws.Cells.Clear
    ws.Cells.ClearComments

    For r = 1 To rCount
        For c = 1 To cCount
            If IsNull(data(r, c)) Or IsEmpty(data(r, c)) Then data(r, c) = Empty
        Next c
    Next r

    ws.Range(ws.Cells(1, 1), ws.Cells(rCount, cCount)).Value2 = data

    ' Header styling
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, cCount))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
    End With

    ' Apply number formats by column role
    If rCount > 1 Then
        ws.Range(ws.Cells(2, 2), ws.Cells(rCount, 3)).NumberFormat = "@"
        For c = 5 To cCount
            colHeader = UCase$(CStr(data(1, c)))
            If Right$(colHeader, 4) = "_AMT" Or UCase$(sliceMode) = "AMOUNT" Then
                ws.Range(ws.Cells(2, c), ws.Cells(rCount, c)).NumberFormat = "#,##0.00;[Red]-#,##0.00;0.00"
            Else
                ws.Range(ws.Cells(2, c), ws.Cells(rCount, c)).NumberFormat = "0.00%;[Red]-0.00%;0.00%"
            End If
        Next c
    End If

    ' Apply row indentation, subtotal highlighting, and audit comments
    For r = 2 To rCount
        styleParts = Split(CStr(data(r, 4)), "|")
        indentLvl = CLng(styleParts(0))
        isBold = (styleParts(1) = "Y")

        ws.Cells(r, 3).IndentLevel = indentLvl
        If isBold Then
            With ws.Range(ws.Cells(r, 1), ws.Cells(r, cCount))
                .Font.Bold = True
                .Interior.Color = RGB(237, 242, 248)
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeBottom).LineStyle = xlDouble
            End With
        End If

        For c = 5 To cCount
            noteText = CStr(notes(r, c))
            If Len(noteText) > 0 Then
                ws.Cells(r, c).AddComment noteText
            End If
        Next c
    Next r

    ' Hide internal STYLE_FLAG helper column (col 4) and autofit visible columns
    ws.columns(4).Hidden = True
    ws.columns.AutoFit
    ws.Range(ws.Cells(1, 1), ws.Cells(rCount, cCount)).AutoFilter
End Sub

Private Sub RenderStyledLongSheet(ByVal result As clsAnalysisResult, ByVal tableName As String, ByVal ws As Worksheet)
    Dim data As Variant, rCount As Long, cCount As Long, r As Long, c As Long
    data = result.table(tableName).ToTable()
    rCount = UBound(data, 1)
    cCount = UBound(data, 2)

    ws.Cells.Clear
    For r = 1 To rCount
        For c = 1 To cCount
            If IsNull(data(r, c)) Or IsEmpty(data(r, c)) Then data(r, c) = Empty
        Next c
    Next r

    If rCount > 1 Then
        ws.Range(ws.Cells(2, 1), ws.Cells(rCount, 6)).NumberFormat = "@"
        ws.Range(ws.Cells(2, 7), ws.Cells(rCount, 8)).NumberFormat = "#,##0.00;[Red]-#,##0.00;0.00"
        ws.Range(ws.Cells(2, 9), ws.Cells(rCount, 9)).NumberFormat = "0.00%;[Red]-0.00%;0.00%"
        ws.Range(ws.Cells(2, 10), ws.Cells(rCount, 10)).NumberFormat = "#,##0.00;[Red]-#,##0.00;0.00"
        ws.Range(ws.Cells(2, 11), ws.Cells(rCount, 11)).NumberFormat = "0.00%;[Red]-0.00%;0.00%"
    End If

    ws.Range(ws.Cells(1, 1), ws.Cells(rCount, cCount)).Value2 = data
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, cCount))
        .Font.Bold = True
        .Interior.Color = RGB(237, 242, 248)
    End With
    ws.Range(ws.Cells(1, 1), ws.Cells(rCount, cCount)).AutoFilter
    ws.columns.AutoFit
End Sub

Private Sub ResolveTickerOrCompanyCode(ByVal token As String, ByRef outCompanyCode As String, ByRef outLabel As String)
    Dim masterSvc As clsSecurityMasterService, masterTbl As clsDataTable
    Dim r As Long, cleanTok As String, rowSec As String, rowComp As String, rowName As String

    cleanTok = UCase$(Trim$(token))
    If InStr(cleanTok, ".") > 0 Then cleanTok = Left$(cleanTok, InStr(cleanTok, ".") - 1)

    Set masterSvc = New clsSecurityMasterService
    Set masterTbl = masterSvc.LoadMasterTable()

    For r = 1 To masterTbl.rowCount
        rowSec = UCase$(Trim$(CStr(masterTbl.ValueAt(r, "SEC_CODE"))))
        If rowSec = cleanTok And Not IsNull(masterTbl.ValueAt(r, "COMPANY_CODE")) Then
            rowComp = Trim$(CStr(masterTbl.ValueAt(r, "COMPANY_CODE")))
            If Len(rowComp) > 0 Then
                outCompanyCode = rowComp
                rowName = IIf(IsNull(masterTbl.ValueAt(r, "SEC_NAME")), rowSec, Trim$(CStr(masterTbl.ValueAt(r, "SEC_NAME"))))
                outLabel = rowName & " (" & rowSec & ")"
                Exit Sub
            End If
        End If
    Next r

    ' Fallback: treat token directly as numeric CompanyCode
    outCompanyCode = cleanTok
    outLabel = "Company_" & cleanTok
End Sub
