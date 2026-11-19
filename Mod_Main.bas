Attribute VB_Name = "Mod_Main"
Option Explicit

Public Sub RunPipeline_Test()
    Dim companyCode As String, reportDate As String, asOfDate As String
    Dim srv As New clsFinancialService, stmt As clsFinancialStatement
    Dim context As clsAnalysisContext, app As clsFinancialAnalysisApplication
    Dim result As clsAnalysisResult, renderer As clsResultTableRenderer
    Dim ws As Worksheet, periodEnd As String, cutoff As String
    On Error GoTo Failed
    companyCode = Trim$(InputBox("Enter numeric CompanyCode (not a security ticker):", "Financial Engine"))
    If companyCode = "" Then Exit Sub
    reportDate = Trim$(InputBox("Annual report end date (yyyy-mm-dd):", "Financial Engine", "2025-12-31"))
    If reportDate = "" Then Exit Sub
    asOfDate = Trim$(InputBox("Publication cutoff (yyyy-mm-dd):", "Financial Engine", Format$(Date, "yyyy-mm-dd")))
    If asOfDate = "" Then Exit Sub

    periodEnd = FinancialPeriodEnd(reportDate)
    cutoff = Format$(CDate(asOfDate), "yyyy-mm-dd")
    Set stmt = srv.GetReport(companyCode, periodEnd, 1, 2, cutoff)

    Set context = New clsAnalysisContext
    context.Initialize CDate(periodEnd), CDate(cutoff), "CNY", "FINANCIAL_SERIES_V1"
    Set app = New clsFinancialAnalysisApplication
    Set result = app.RunSelectedSeries(stmt, context, 2)

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    Set renderer = New clsResultTableRenderer
    renderer.Render result, "SELECTED_SERIES", ws
    ws.rows(1).Interior.Color = RGB(218, 232, 247)
    MsgBox "Selected adj2 report series exported to " & ws.name & " (RunId=" & result.RunId & "). NULL amounts are blank; zero remains zero. No adjusted-flag fallback is used.", vbInformation, "Financial Engine"
    Exit Sub
Failed:
    MsgBox "Export failed: " & Err.description, vbExclamation, "Financial Engine"
End Sub

Public Sub ClearFinancialCache()
    ResetGlobals
    Application.CalculateFull
End Sub
