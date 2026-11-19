Attribute VB_Name = "Mod_MetricAudit"
Option Explicit

' Manual entry point. No database access occurs until this macro is run.
Public Sub ExportMetricAudit()
    Dim company As String, period As String, metric As String, cutoff As String
    Dim ws As Worksheet
    On Error GoTo Failed
    company = Trim$(InputBox("Vendor CompanyCode (not stock ticker):", "Metric audit"))
    If Len(company) = 0 Then Exit Sub
    period = Trim$(InputBox("Report period (YYYY or YYYY-MM-DD):", "Metric audit", "2025-12-31"))
    If Len(period) = 0 Then Exit Sub
    metric = UCase$(Trim$(InputBox("Metric code:", "Metric audit", "RNOA")))
    If Len(metric) = 0 Then Exit Sub
    cutoff = Trim$(InputBox("As-of date (YYYY-MM-DD):", "Metric audit", Format$(Date, "yyyy-mm-dd")))
    If Len(cutoff) = 0 Then Exit Sub
    Set ws = CreateMetricAudit(company, period, metric, 1, 2, cutoff)
    ws.Activate
    MsgBox "Audit exported to " & ws.name & ". Check Status and Diagnostic, including ERROR rows.", vbInformation
    Exit Sub
Failed:
    MsgBox "Audit export failed: " & Err.description, vbExclamation
End Sub

' For other flags, exact publication pins, or deeper nested AVG history, call
' this API explicitly. It creates a sheet in this code's workbook, not ActiveWorkbook.
Public Function CreateMetricAudit(ByVal company As String, ByVal period As String, ByVal metric As String, _
        Optional ByVal ifMerged As Long = 1, Optional ByVal ifAdjusted As Integer = 2, _
        Optional ByVal asOfDate As String = "", Optional ByVal publishedDate As String = "", _
        Optional ByVal historyYears As Long = 5) As Worksheet
    Dim service As clsFinancialService, stmt As clsFinancialStatement
    Dim context As clsAnalysisContext, app As clsFinancialAnalysisApplication, result As clsAnalysisResult
    Dim auditTable As clsDataTable, lineageData As Variant, ws As Worksheet
    Dim rows As Variant, headers As Variant, metadata(1 To 7, 1 To 2) As Variant
    Dim i As Long, j As Long, count As Long, lastRow As Long, suffix As Long
    Dim baseName As String, candidate As String, diagnostic As String
    Dim n As Long, description As String
    On Error GoTo Failed
    If historyYears < 0 Or historyYears > 100 Then Err.Raise 5, , "History years must be 0..100."
    If ifMerged <> 1 And ifMerged <> 2 Then Err.Raise 5, , "IfMerged must be 1 or 2."
    If ifAdjusted <> 1 And ifAdjusted <> 2 Then Err.Raise 5, , "IfAdjusted must be 1 or 2."
    period = FinancialPeriodEnd(period)
    metric = UCase$(Trim$(metric))
    If Len(metric) = 0 Then Err.Raise 5, , "Metric code is required."
    If Len(asOfDate) = 0 Then asOfDate = Format$(Date, "yyyy-mm-dd")
    asOfDate = Format$(CDate(asOfDate), "yyyy-mm-dd")
    If Len(publishedDate) > 0 Then publishedDate = Format$(CDate(publishedDate), "yyyy-mm-dd")

    Set service = New clsFinancialService
    Set stmt = service.GetStatement(company, Format$(DateAdd("yyyy", -historyYears, CDate(period)), "yyyy-mm-dd"), period, ifMerged, asOfDate, publishedDate, ifAdjusted)

    Set context = New clsAnalysisContext
    context.Initialize CDate(period), CDate(asOfDate), "CNY", "FINANCIAL_AUDIT_V1"
    Set app = New clsFinancialAnalysisApplication
    Set result = app.RunAuditMetric(stmt, context, period, metric, CLng(ifAdjusted))

    Set auditTable = result.table("METRIC_AUDIT")
    count = auditTable.rowCount
    If count = 0 Then Err.Raise 5, , "No audit rows produced."
    ReDim rows(1 To count, 1 To 14)
    For i = 1 To count
        For j = 1 To 14
            rows(i, j) = auditTable.ValueAtColumn(i, j)
        Next j
    Next i
    If CStr(rows(1, 13)) = "ERROR" Then diagnostic = CStr(rows(1, 14))

    lastRow = 9 + count
    baseName = "Audit_" & Format$(Now, "yymmdd_hhnnss")
    candidate = baseName
    Do While AuditSheetExists(candidate)
        suffix = suffix + 1
        candidate = baseName & "_" & CStr(suffix)
    Loop
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    ws.name = candidate
    ws.Range("A1:N7").NumberFormat = "@"
    metadata(1, 1) = "CompanyCode": metadata(1, 2) = stmt.companyCode
    metadata(2, 1) = "Metric / period": metadata(2, 2) = metric & " / " & period
    metadata(3, 1) = "Merged / adjusted": metadata(3, 2) = CStr(ifMerged) & " / " & CStr(ifAdjusted)
    metadata(4, 1) = "As-of / end-period pin": metadata(4, 2) = asOfDate & " / " & publishedDate
    metadata(5, 1) = "History years / exported": metadata(5, 2) = CStr(historyYears) & " / " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    metadata(6, 1) = "Run ID / result status": metadata(6, 2) = result.RunId & " / " & result.status
    metadata(7, 1) = "Evaluation diagnostic": metadata(7, 2) = diagnostic
    ws.Range("A1:B7").Value2 = metadata
    ws.Range("A8").Value2 = "Snapshot after local overrides; child values are NOT additive contributions. NULL/missing raw values are blank; see DEFAULT_ZERO."
    headers = Array("NodeID", "ParentID", "Level", "NodeType", "Identifier / role", "ReportPeriod", "AccountCode", "Expression", "RawValue", "CalculatedValue", "SelectedPublication", "SelectedVersionKey", "Status", "Diagnostic")
    For i = 0 To UBound(headers)
        ws.Cells(9, i + 1).Value2 = headers(i)
    Next i
    ws.Range("E10:H" & CStr(lastRow)).NumberFormat = "@"
    ws.Range("K10:N" & CStr(lastRow)).NumberFormat = "@"
    ws.Range("A10").Resize(count, 14).Value2 = rows
    ws.Range("I10:J" & CStr(lastRow)).NumberFormat = "#,##0.000000;[Red]-#,##0.000000;0.000000"
    For i = 1 To count
        ws.Cells(i + 9, 5).IndentLevel = WorksheetFunction.Min(CLng(rows(i, 3)), 15)
        Select Case CStr(rows(i, 13))
        Case "DEFAULT_ZERO": ws.Range(ws.Cells(i + 9, 1), ws.Cells(i + 9, 14)).Interior.Color = RGB(255, 242, 204)
        Case "ERROR": ws.Range(ws.Cells(i + 9, 1), ws.Cells(i + 9, 14)).Interior.Color = RGB(255, 199, 206)
        End Select
    Next i
    With ws.Range("A9:N9")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
    End With
    ws.Range("A9:N" & CStr(lastRow)).AutoFilter

    ' Append dataset lineage provenance below the audit tree.
    lineageData = result.table("LINEAGE").ToTable()
    ws.Cells(lastRow + 2, 1).Value2 = "Lineage"
    ws.Cells(lastRow + 2, 1).Font.Bold = True
    ws.Range(ws.Cells(lastRow + 3, 1), ws.Cells(lastRow + 2 + UBound(lineageData, 1), UBound(lineageData, 2))).NumberFormat = "@"
    ws.Cells(lastRow + 3, 1).Resize(UBound(lineageData, 1), UBound(lineageData, 2)).Value2 = lineageData
    ws.rows(lastRow + 3).Font.Bold = True

    ws.columns("A:D").ColumnWidth = 12
    ws.columns("E:G").ColumnWidth = 28
    ws.columns("H:H").ColumnWidth = 60
    ws.columns("I:K").ColumnWidth = 20
    ws.columns("L:L").ColumnWidth = 60
    ws.columns("M:M").ColumnWidth = 18
    ws.columns("N:N").ColumnWidth = 65
    Set CreateMetricAudit = ws
    Exit Function
Failed:
    n = Err.number
    description = Err.description
    If Not ws Is Nothing Then description = description & " (Partial sheet retained: " & ws.name & ")"
    Err.Raise n, "Mod_MetricAudit.CreateMetricAudit", description
End Function

Private Function AuditSheetExists(ByVal sheetName As String) As Boolean
    Dim item As Object
    For Each item In ThisWorkbook.Sheets
        If StrComp(item.name, sheetName, vbTextCompare) = 0 Then
            AuditSheetExists = True
            Exit Function
        End If
    Next item
End Function
