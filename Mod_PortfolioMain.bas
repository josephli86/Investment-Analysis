Attribute VB_Name = "Mod_PortfolioMain"
Option Explicit

Public Sub DebugPortfolioPanelTable()
    Dim code As Variant, startText As Variant, endText As Variant, tableChoice As Variant
    Dim startDay As Date, endDay As Date, panel As clsPortfolioPanel
    Dim output As Variant, tableName As String, book As Workbook, sheet As Worksheet
    Dim analyzer As clsPortfolioAnalyzer

    On Error GoTo Failed
    code = Application.InputBox("Portfolio code (text, leading zeros preserved):", "Debug Panel Table", Type:=2)
    If VarType(code) = vbBoolean Then Exit Sub
    startText = Application.InputBox("Start date: yyyy-mm-dd", "Debug Panel Table", Format$(Date - 30, "yyyy-mm-dd"), Type:=2)
    If VarType(startText) = vbBoolean Then Exit Sub
    startDay = ParsePortfolioDate(CStr(startText))
    endText = Application.InputBox("End date: yyyy-mm-dd", "Debug Panel Table", Format$(Date - 1, "yyyy-mm-dd"), Type:=2)
    If VarType(endText) = vbBoolean Then Exit Sub
    endDay = ParsePortfolioDate(CStr(endText))
    If startDay > endDay Then Err.Raise 5, , "Start date cannot be after end date."

    tableChoice = Application.InputBox("Select table number:" & vbCrLf & _
        "1 = Holdings" & vbCrLf & "2 = Performance" & vbCrLf & _
        "3 = Profit" & vbCrLf & "4 = Orders" & vbCrLf & _
        "5 = IndustryAnalysis" & vbCrLf & "6 = Summary", _
        "Debug Panel Table", Type:=1)
    If VarType(tableChoice) = vbBoolean Then Exit Sub
    If tableChoice < 1 Or tableChoice > 6 Or tableChoice <> Fix(tableChoice) Then Err.Raise 5, , "Choose a table number from 1 to 6."

    Set panel = LoadPortfolioPanel(CStr(code), startDay, endDay)
    Select Case CLng(tableChoice)
    Case 1
        tableName = "Holdings": output = panel.HoldingsTable
    Case 2
        tableName = "Performance": output = panel.PerformanceTable
    Case 3
        tableName = "Profit": output = panel.ProfitTable
    Case 4
        tableName = "Orders": output = panel.ordersTable
    Case 5
        tableName = "IndustryAnalysis"
        Set analyzer = New clsPortfolioAnalyzer
        analyzer.Initialize panel
        output = analyzer.IndustryAnalysis
    Case 6
        tableName = "Summary"
        Set book = RenderPortfolioReport(panel)
        book.Activate
        Exit Sub
    End Select

    Set book = Application.Workbooks.Add(xlWBATWorksheet)
    Set sheet = book.Worksheets(1)
    sheet.name = tableName
    WritePortfolioTable sheet, output
    book.Activate
    Exit Sub
Failed:
    MsgBox "Debug panel table failed: " & Err.description, vbExclamation, "Debug Panel Table"
End Sub

Public Sub RunPortfolioAnalysis()
    Dim code As Variant, startText As Variant, endText As Variant
    Dim startDay As Date, endDay As Date, report As Workbook
    Dim panel As clsPortfolioPanel

    On Error GoTo Failed

    code = Application.InputBox("Portfolio code (text, leading zeros preserved):", "Portfolio analysis", Type:=2)
    If VarType(code) = vbBoolean Then Exit Sub
    startText = Application.InputBox("Start date: yyyy-mm-dd", "Portfolio analysis", Format$(Date - 30, "yyyy-mm-dd"), Type:=2)
    If VarType(startText) = vbBoolean Then Exit Sub
    startDay = ParsePortfolioDate(CStr(startText))
    endText = Application.InputBox("End date: yyyy-mm-dd", "Portfolio analysis", Format$(Date - 1, "yyyy-mm-dd"), Type:=2)
    If VarType(endText) = vbBoolean Then Exit Sub
    endDay = ParsePortfolioDate(CStr(endText))
    If startDay > endDay Then Err.Raise 5, , "Start date cannot be after end date."

    Set panel = LoadPortfolioPanel(CStr(code), startDay, endDay)
    Set report = RenderPortfolioReport(panel)
    report.Activate
    Exit Sub
Failed:
    MsgBox Err.description, vbExclamation, "Portfolio analysis"
End Sub

Public Sub RunPortfolioReturnAnalysis()
    Dim code As Variant, endText As Variant
    Dim endDay As Date, report As Workbook
    Dim panel As clsPortfolioPanel

    On Error GoTo Failed

    code = Application.InputBox("Portfolio code (text, leading zeros preserved):", "Annual return analysis", Type:=2)
    If VarType(code) = vbBoolean Then Exit Sub
    endText = Application.InputBox("End date: yyyy-mm-dd", "Annual return analysis", Format$(Date - 1, "yyyy-mm-dd"), Type:=2)
    If VarType(endText) = vbBoolean Then Exit Sub
    endDay = ParsePortfolioDate(CStr(endText))

    Set panel = LoadPortfolioReturnPanel(CStr(code), endDay)
    Set report = RenderPortfolioReturnReport(panel)
    report.Activate
    Exit Sub
Failed:
    MsgBox Err.description, vbExclamation, "Annual return analysis"
End Sub

Public Function ParsePortfolioDate(ByVal text As String) As Date
    Dim result As Date
    If Not text Like "####-##-##" Then Err.Raise 5, , "Use yyyy-mm-dd."
    result = DateSerial(CInt(Left$(text, 4)), CInt(Mid$(text, 6, 2)), CInt(Right$(text, 2)))
    If Format$(result, "yyyy-mm-dd") <> text Then Err.Raise 5, , "Invalid valuation date."
    ParsePortfolioDate = result
End Function

Public Function LoadPortfolioPanel(ByVal portCode As String, _
        ByVal startDate As Date, _
        ByVal endDate As Date) As clsPortfolioPanel

    Dim service As clsPortfolioService
    Set service = New clsPortfolioService
    Set LoadPortfolioPanel = service.GetPanel(Array(portCode), startDate, endDate)
End Function

Public Function LoadPortfolioReturnPanel(ByVal portCode As String, ByVal endDate As Date) As clsPortfolioPanel
    Dim service As clsPortfolioService
    Set service = New clsPortfolioService
    Set LoadPortfolioReturnPanel = service.GetReturnPanel(Array(portCode), endDate)
End Function

Public Function CreatePortfolioReport(ByVal portCode As String, _
        ByVal valuationDate As Date, _
        Optional ByVal amountsComparableConfirmed As Boolean = False) As Workbook
    Dim panel As clsPortfolioPanel
    Set panel = LoadPortfolioPanel(portCode, valuationDate, valuationDate)
    Set CreatePortfolioReport = RenderPortfolioReport(panel, amountsComparableConfirmed)
End Function

Public Function CreatePortfolioTimeSeriesReport(ByVal portCode As String, _
        ByVal startDate As Date, _
        ByVal endDate As Date) As Workbook
    Dim panel As clsPortfolioPanel
    Set panel = LoadPortfolioPanel(portCode, startDate, endDate)
    Set CreatePortfolioTimeSeriesReport = RenderPortfolioSeriesReport(panel)
End Function

Public Function RenderPortfolioSeriesReport(ByVal panel As clsPortfolioPanel) As Workbook
    Dim book As Workbook, sheet As Worksheet, table(1 To 6, 1 To 2) As Variant
    If panel Is Nothing Then Err.Raise 5, , "Portfolio panel is required."
    Set book = Application.Workbooks.Add(xlWBATWorksheet)
    Set sheet = book.Worksheets(1)
    sheet.name = "TimeSeriesSummary"
    table(1, 1) = "PORT_CODE": table(1, 2) = panel.PortfolioCodes
    table(2, 1) = "START_DATE": table(2, 2) = Format$(panel.startDate, "yyyy-mm-dd")
    table(3, 1) = "END_DATE": table(3, 2) = Format$(panel.endDate, "yyyy-mm-dd")
    table(4, 1) = "HOLDING_ROWS": table(4, 2) = panel.HoldingRowCount
    table(5, 1) = "PERFORMANCE_ROWS": table(5, 2) = panel.PerformanceRowCount
    table(6, 1) = "STATUS": table(6, 2) = "Panel data loaded for requested range."
    WritePortfolioTable sheet, table
    Set RenderPortfolioSeriesReport = book
End Function

Public Function RenderPortfolioReport(ByVal panel As clsPortfolioPanel, Optional ByVal amountsComparableConfirmed As Boolean = False) As Workbook
    Dim book As Workbook, sheet As Worksheet, result As clsAnalysisResult
    Dim context As clsAnalysisContext, app As clsPortfolioAnalysisApplication, returnResult As clsAnalysisResult
    Dim industryResult As clsAnalysisResult, attributionResult As clsAnalysisResult, holdingResult As clsAnalysisResult
    Dim resultName As Variant, tableName As Variant, table As Variant
    Dim number As Long, description As String
    On Error GoTo Failed
    If panel Is Nothing Then Err.Raise 5, , "Portfolio panel is required."
    Set context = New clsAnalysisContext
    context.Initialize panel.endDate, panel.endDate, "CNY", "PORTFOLIO_ANALYSIS_V1"
    Set app = New clsPortfolioAnalysisApplication
    Set returnResult = app.RunReturnAnalysis(panel, context)
    Set industryResult = app.RunIndustryAnalysis(panel, context)
    Set attributionResult = app.RunAbsoluteAttribution(panel, context)
    Set holdingResult = app.RunHoldingSummary(panel, context, amountsComparableConfirmed)
    Set book = Application.Workbooks.Add(xlWBATWorksheet)
    Set sheet = book.Worksheets(1): sheet.name = "Summary"
    RenderPortfolioSummary panel, holdingResult, sheet
    For Each resultName In Array("HOLDINGS", "PERFORMANCE", "PROFIT", "ORDERS", "INDUSTRY_ANALYSIS", "ABSOLUTE_ATTRIBUTION", "RETURNS_ANALYSIS")
        Select Case CStr(resultName)
            Case "HOLDINGS": tableName = "Holdings": Set result = app.RunPanelTable(panel, context, "HOLDINGS", panel.HoldingsTable)
            Case "PERFORMANCE": tableName = "Performance": Set result = app.RunPanelTable(panel, context, "PERFORMANCE", panel.PerformanceTable)
            Case "PROFIT": tableName = "Profit": Set result = app.RunPanelTable(panel, context, "PROFIT", panel.ProfitTable)
            Case "ORDERS": tableName = "Orders": Set result = app.RunPanelTable(panel, context, "ORDERS", panel.ordersTable)
            Case "INDUSTRY_ANALYSIS": tableName = "IndustryAnalysis": Set result = industryResult
            Case "ABSOLUTE_ATTRIBUTION": tableName = "AbsoluteAttribution": Set result = attributionResult
            Case "RETURNS_ANALYSIS": tableName = "ReturnsAnalysis": Set result = returnResult
        End Select
        Set sheet = book.Worksheets.Add(After:=book.Worksheets(book.Worksheets.count))
        sheet.name = CStr(tableName)
        table = result.table(CStr(resultName)).ToTable()
        RenderPortfolioResultTable result, CStr(resultName), sheet
        If CStr(resultName) = "RETURNS_ANALYSIS" Then WriteReturnDefinitions sheet, UBound(table, 1)
    Next resultName
    book.Worksheets("Summary").Activate
    Set RenderPortfolioReport = book
    Exit Function
Failed:
    number = Err.number: description = Err.description
    If Not book Is Nothing Then book.Close SaveChanges:=False
    Err.Raise number, "Mod_PortfolioMain.RenderPortfolioReport", description
End Function

Private Sub RenderPortfolioSummary(ByVal panel As clsPortfolioPanel, ByVal holdingResult As clsAnalysisResult, ByVal sheet As Worksheet)
    Dim summary(1 To 10, 1 To 2) As Variant
    summary(1, 1) = "PORT_CODE": summary(1, 2) = panel.PortfolioCodes
    summary(2, 1) = "START_DATE": summary(2, 2) = Format$(panel.startDate, "yyyy-mm-dd")
    summary(3, 1) = "END_DATE": summary(3, 2) = Format$(panel.endDate, "yyyy-mm-dd")
    summary(4, 1) = "SOURCE": summary(4, 2) = panel.sourceAlias
    summary(5, 1) = "HOLDING_ROWS": summary(5, 2) = panel.HoldingRowCount
    summary(6, 1) = "PERFORMANCE_ROWS": summary(6, 2) = panel.PerformanceRowCount
    summary(7, 1) = "INDUSTRY_ROWS": summary(7, 2) = panel.IndustryRowCount
    summary(8, 1) = "TOTAL_HOLDING_VALUE": summary(8, 2) = holdingResult.table("HOLDING_SUMMARY").ValueAt(1, "VALUE")
    summary(9, 1) = "STATUS": summary(9, 2) = "RTN is decimal; security PnL excludes fund expenses; no reconciliation to RTN."
    summary(10, 1) = "METHOD": summary(10, 2) = "Frongello; contribution = DAILYPROFIT / prior valuation NAV."
    WritePortfolioTable sheet, summary
End Sub

Private Sub RenderPortfolioResultTable(ByVal result As clsAnalysisResult, ByVal resultTable As String, ByVal sheet As Worksheet)
    Dim renderer As clsResultTableRenderer
    Set renderer = New clsResultTableRenderer
    renderer.Render result, resultTable, sheet
    With sheet.UsedRange
        .rows(1).Interior.Color = RGB(218, 232, 247)
    End With
End Sub

Public Function RenderPortfolioReturnReport(ByVal panel As clsPortfolioPanel) As Workbook
    Dim book As Workbook, sheet As Worksheet
    Dim table As Variant, summary(1 To 5, 1 To 2) As Variant
    Dim context As clsAnalysisContext, app As clsPortfolioAnalysisApplication, result As clsAnalysisResult, renderer As clsResultTableRenderer
    If panel Is Nothing Then Err.Raise 5, , "Portfolio panel is required."
    Set context = New clsAnalysisContext
    context.Initialize panel.endDate, panel.endDate, "CNY", "PORTFOLIO_RETURN_V1"
    Set app = New clsPortfolioAnalysisApplication
    Set result = app.RunReturnAnalysis(panel, context)
    table = result.table("RETURNS_ANALYSIS").ToTable()
    Set book = Application.Workbooks.Add(xlWBATWorksheet)
    Set sheet = book.Worksheets(1)
    sheet.name = "Summary"
    summary(1, 1) = "PORT_CODE": summary(1, 2) = panel.PortfolioCodes
    summary(2, 1) = "START_DATE": summary(2, 2) = Format$(panel.startDate, "yyyy-mm-dd")
    summary(3, 1) = "END_DATE": summary(3, 2) = Format$(panel.endDate, "yyyy-mm-dd")
    summary(4, 1) = "PERFORMANCE_ROWS": summary(4, 2) = panel.PerformanceRowCount
    summary(5, 1) = "STATUS": summary(5, 2) = "Annual return analysis only; holdings, attribution and industry data were not loaded into this report."
    WritePortfolioTable sheet, summary
    Set sheet = book.Worksheets.Add(After:=book.Worksheets(book.Worksheets.count))
    sheet.name = "ReturnsAnalysis"
    Set renderer = New clsResultTableRenderer
    renderer.Render result, "RETURNS_ANALYSIS", sheet
    WriteReturnDefinitions sheet, UBound(table, 1)
    book.Worksheets("Summary").Activate
    Set RenderPortfolioReturnReport = book
End Function

Private Sub WriteReturnDefinitions(ByVal sheet As Worksheet, ByVal rowCount As Long)
    sheet.Cells(rowCount + 3, 1).Value = "Definitions"
    sheet.Cells(rowCount + 4, 1).Value = "One row per calendar year from the earliest available return observation (used as inception proxy) through the selected end date."
    sheet.Cells(rowCount + 5, 1).Value = "Returns are compounded from decimal daily RTN and RTN_BM; excess = portfolio annual return - benchmark annual return."
    sheet.Cells(rowCount + 6, 1).Value = "Annualized volatility and Sharpe use sample daily standard deviation; ratios use zero risk-free rate and Sqr(252)."
    sheet.Cells(rowCount + 7, 1).Value = "Sortino uses zero target and downside RMS across all daily observations, annualized by Sqr(252)."
End Sub

Private Sub WritePortfolioTable(ByVal sheet As Worksheet, ByVal table As Variant)
    Dim r As Long, c As Long, rowCount As Long, colCount As Long, hasString As Boolean
    rowCount = UBound(table, 1)
    colCount = UBound(table, 2)
    If rowCount > sheet.rows.count Then Err.Raise 6, , "Too many rows for Excel."

    For r = 1 To rowCount
        For c = 1 To colCount
            If IsNull(table(r, c)) Then table(r, c) = Empty
        Next c
    Next r

    ' Apply "@" text format to columns containing string data so leading zeros survive without literal apostrophes.
    For c = 1 To colCount
        hasString = False
        For r = 2 To rowCount
            If VarType(table(r, c)) = vbString Then
                hasString = True
                Exit For
            End If
        Next r
        If hasString Then
            sheet.Cells(1, c).Resize(rowCount, 1).NumberFormat = "@"
        End If
    Next c

    With sheet.Cells(1, 1).Resize(rowCount, colCount)
        .Value2 = table
        .rows(1).Font.Bold = True
        .rows(1).Interior.Color = RGB(218, 232, 247)
        .columns.AutoFit
    End With
End Sub
