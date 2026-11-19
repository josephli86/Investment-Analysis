Attribute VB_Name = "Mod_ReportTemplate"
Option Explicit

Private Const TEMPLATE_FILE As String = "\Config\RPT_TEMPLATE_01.xlsx"

Public Sub ExportTemplateReport()
    Dim company As String, firstYear As String, lastYear As String
    Dim merged As String, adjusted As String, cutoff As String, anchor As Range, result As Range
    On Error GoTo Failed
    company = AskText("CompanyCode (numeric, not ticker):", "")
    If company = "" Then Exit Sub
    firstYear = AskText("Start fiscal year (annual December reports):", CStr(Year(Date) - 3))
    If firstYear = "" Then Exit Sub
    lastYear = AskText("End fiscal year:", CStr(Year(Date) - 1))
    If lastYear = "" Then Exit Sub
    merged = AskText("IfMerged: 1 or 2", "1")
    If merged = "" Then Exit Sub
    adjusted = AskText("IfAdjusted: 1 or 2", "2")
    If adjusted = "" Then Exit Sub
    cutoff = AskText("As-of cutoff (yyyy-mm-dd):", Format$(Date, "yyyy-mm-dd"))
    If cutoff = "" Then Exit Sub
    If Not firstYear Like "####" Or Not lastYear Like "####" Then Err.Raise 5, , "Enter four-digit fiscal years."
    If merged <> "1" And merged <> "2" Then Err.Raise 5, , "IfMerged must be 1 or 2."
    If adjusted <> "1" And adjusted <> "2" Then Err.Raise 5, , "IfAdjusted must be 1 or 2."
    On Error Resume Next
    Set anchor = Application.InputBox("Select ONE output start cell. Prefer a blank new sheet or a blank area. Existing content requires confirmation.", "Template report", Type:=8)
    On Error GoTo Failed
    If anchor Is Nothing Then Exit Sub
    Set result = RenderTemplateReport(company, CLng(firstYear), CLng(lastYear), anchor, CLng(merged), CLng(adjusted), cutoff)
    If Not result Is Nothing Then MsgBox "Report snapshot created: " & result.Address(External:=True) & vbCrLf & "Excel error cells have diagnostic notes. Save the output workbook to retain it.", vbInformation
    Exit Sub
Failed:
    MsgBox "Report failed: " & Err.description, vbExclamation
End Sub

Private Function AskText(ByVal prompt As String, ByVal defaultValue As String) As String
    Dim answer As Variant
    answer = Application.InputBox(prompt, "Template report", defaultValue, Type:=2)
    If VarType(answer) = vbBoolean Then Exit Function
    AskText = Trim$(CStr(answer))
End Function

Public Function RenderTemplateReport(ByVal companyCode As String, ByVal startYear As Long, ByVal endYear As Long, ByVal outputStart As Range, Optional ByVal ifMerged As Long = 1, Optional ByVal ifAdjusted As Long = 2, Optional ByVal asOfDate As String = "") As Range
    Dim service As clsFinancialService, stmt As clsFinancialStatement, engine As clsMeasureEngine
    Dim context As clsAnalysisContext, app As clsFinancialAnalysisApplication, analysisResult As clsAnalysisResult
    Dim valuesData As clsDataTable, notesData As clsDataTable
    Dim templateBook As Workbook, source As Worksheet, target As Range, owned As Boolean
    Dim eventsBefore As Boolean, screenBefore As Boolean, securityBefore As Long, changed As Boolean
    Dim codes As Variant, values() As Variant, notes() As String, defs As Object, raw As Object
    Dim r As Long, col As Long, count As Long, item As Variant
    Dim n As Long, description As String, templatePath As String
    On Error GoTo Failed
    Set service = New clsFinancialService
    companyCode = service.NormalizeCompanyCode(companyCode)
    If startYear < 1901 Or endYear > 9999 Or startYear > endYear Then Err.Raise 5, , "Years must satisfy 1901 <= start <= end <= 9999."
    If ifMerged <> 1 And ifMerged <> 2 Then Err.Raise 5, , "IfMerged must be 1 or 2."
    If ifAdjusted <> 1 And ifAdjusted <> 2 Then Err.Raise 5, , "IfAdjusted must be 1 or 2."
    If asOfDate = "" Then asOfDate = Format$(Date, "yyyy-mm-dd")
    asOfDate = Format$(CDate(asOfDate), "yyyy-mm-dd")
    count = endYear - startYear + 1
    Set target = CheckTarget(outputStart, count)
    If HasContent(target) Then
        If MsgBox("Replace ALL content and notes in " & target.Address(External:=True) & "?", vbYesNo + vbExclamation + vbDefaultButton2, "Confirm report overwrite") <> vbYes Then Exit Function
    End If
    eventsBefore = Application.EnableEvents
    screenBefore = Application.ScreenUpdating
    securityBefore = Application.AutomationSecurity
    changed = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    templatePath = ThisWorkbook.Path & TEMPLATE_FILE
    Set templateBook = FindOpenBook(templatePath)
    If templateBook Is Nothing Then
        Application.AutomationSecurity = 3
        Set templateBook = Application.Workbooks.Open(fileName:=templatePath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, Notify:=False)
        owned = True
        Application.AutomationSecurity = securityBefore
    End If
    Set source = templateBook.Worksheets("Sheet3")
    If source.Range("A6").Value2 <> "ACCT_CODE" Then Err.Raise 5, , "Template Sheet3!A6 must be ACCT_CODE."
    For Each item In source.Range("A1:A45").Cells
        If item.MergeCells Or item.HasFormula Then Err.Raise 5, , "Template must contain unmerged literal cells."
    Next item
    codes = source.Range("A1:A45").Value2
    Set defs = ReadDefinitions()
    Set raw = ReadRawCodes()
    ValidateCodes codes, defs, raw
    Set engine = MeasureMgr()
    If engine.DefinitionCount = 0 Then Err.Raise 5, , "No metric definitions loaded."

    Set stmt = service.GetStatement(companyCode, CStr(startYear - 1) & "-12-31", CStr(endYear) & "-12-31", ifMerged, asOfDate, "", ifAdjusted)
    Set context = New clsAnalysisContext
    context.Initialize DateSerial(endYear, 12, 31), CDate(asOfDate), "CNY", "FINANCIAL_TEMPLATE_V1"
    Set app = New clsFinancialAnalysisApplication
    Set analysisResult = app.RunTemplateSnapshot(stmt, context, service.sourceAlias, startYear, endYear, codes, defs, raw, ifAdjusted, engine)

    Set valuesData = analysisResult.table("TEMPLATE_VALUES")
    Set notesData = analysisResult.table("TEMPLATE_NOTES")
    ReDim values(1 To 45, 1 To count + 1)
    ReDim notes(1 To 45, 1 To count + 1)
    For r = 1 To 45
        For col = 1 To count + 1
            values(r, col) = valuesData.ValueAtColumn(r, col)
            notes(r, col) = CStr(notesData.ValueAtColumn(r, col))
        Next col
    Next r

    Set target = CheckTarget(outputStart, count)
    target.ClearContents
    target.ClearComments
    source.Range("A1:A45").copy
    target.columns(1).PasteSpecial Paste:=xlPasteFormats
    For col = 2 To count + 1
        target.columns(1).copy
        target.columns(col).PasteSpecial Paste:=xlPasteFormats
    Next col
    Application.CutCopyMode = False
    target.Value2 = values
    target.columns(1).ColumnWidth = source.columns(1).ColumnWidth
    For col = 2 To count + 1
        target.columns(col).ColumnWidth = 24
    Next col
    For r = 1 To 45
        target.rows(r).RowHeight = source.rows(r).RowHeight
        For col = 2 To count + 1
            If notes(r, col) <> "" Then target.Cells(r, col).AddComment notes(r, col)
        Next col
    Next r
    target.rows(1).Resize(5).WrapText = True
    target.rows(1).Resize(5).RowHeight = 42
    Set RenderTemplateReport = target
CleanExit:
    If owned Then templateBook.Close SaveChanges:=False
    If changed Then
        Application.AutomationSecurity = securityBefore
        Application.EnableEvents = eventsBefore
        Application.ScreenUpdating = screenBefore
    End If
    Exit Function
Failed:
    n = Err.number
    description = Err.description
    On Error Resume Next
    Application.CutCopyMode = False
    If owned Then templateBook.Close SaveChanges:=False
    If changed Then
        Application.AutomationSecurity = securityBefore
        Application.EnableEvents = eventsBefore
        Application.ScreenUpdating = screenBefore
    End If
    On Error GoTo 0
    Err.Raise n, "Mod_ReportTemplate.RenderTemplateReport", description
End Function

Private Function FindOpenBook(ByVal fullPath As String) As Workbook
    Dim book As Workbook
    For Each book In Application.Workbooks
        If StrComp(book.FullName, fullPath, vbTextCompare) = 0 Then
            Set FindOpenBook = book
            Exit Function
        End If
    Next book
End Function

Private Function CheckTarget(ByVal anchor As Range, ByVal years As Long) As Range
    Dim sheet As Worksheet, book As Workbook, area As Range, cell As Range, table As ListObject
    Dim root As String, bookPath As String, spill As Variant
    If anchor Is Nothing Then Err.Raise 5, , "An output start cell is required."
    If Not (anchor.Application Is Application) Then Err.Raise 5, , "Output must belong to this Excel instance."
    If anchor.Areas.count <> 1 Or anchor.Cells.CountLarge <> 1 Then Err.Raise 5, , "Select exactly one start cell."
    Set sheet = anchor.Worksheet
    Set book = sheet.parent
    root = LCase$(ThisWorkbook.Path & "\Config\")
    bookPath = LCase$(book.FullName)
    If Left$(bookPath, Len(root)) = root Then Err.Raise 5, , "Output cannot overwrite Config source workbooks."
    If book.ReadOnly Then Err.Raise 5, , "Output workbook is read-only."
    If sheet.ProtectContents Or sheet.ProtectDrawingObjects Then Err.Raise 5, , "Output sheet is protected."
    If anchor.row + 44 > sheet.rows.count Or anchor.column + years > sheet.columns.count Then Err.Raise 5, , "Output exceeds worksheet bounds."
    Set area = anchor.Resize(45, years + 1)
    For Each table In sheet.ListObjects
        If Not Application.Intersect(area, table.Range) Is Nothing Then Err.Raise 5, , "Output intersects an Excel table."
    Next table
    For Each cell In area.Cells
        If cell.MergeCells Then Err.Raise 5, , "Output intersects merged cells."
        If cell.HasArray Then Err.Raise 5, , "Output intersects an array formula."
        spill = False
        On Error Resume Next
        spill = CallByName(cell, "HasSpill", VbGet)
        On Error GoTo 0
        If Not IsNull(spill) Then
            If spill Then Err.Raise 5, , "Output intersects a dynamic spill range."
        End If
    Next cell
    Set CheckTarget = area
End Function

Private Function HasContent(ByVal area As Range) As Boolean
    Dim cell As Range, note As Comment, threaded As Object
    HasContent = (Application.CountA(area) > 0)
    For Each cell In area.Cells
        Set note = cell.Comment
        If Not note Is Nothing Then HasContent = True
        Set threaded = Nothing
        On Error Resume Next
        Set threaded = CallByName(cell, "CommentThreaded", VbGet)
        On Error GoTo 0
        If Not threaded Is Nothing Then Err.Raise 5, , "Output intersects threaded comments; choose a blank area."
    Next cell
End Function

Private Function ReadDefinitions() As Object
    Dim conn As Object, rs As Object, result As Object, code As String
    Dim n As Long, description As String
    On Error GoTo Failed
    Set result = CreateObject("Scripting.Dictionary")
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.Path & "\Config\MEASURES_DEF.xlsx;Mode=Read;Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
    Set rs = conn.Execute("SELECT METRIC_CODE,ENABLED FROM [MeasuresDef$]")
    Do While Not rs.EOF
        If Not IsNull(rs.fields(0).Value) Then
            code = UCase$(Trim$(CStr(rs.fields(0).Value)))
            If code <> "" Then
                If result.exists(code) Then Err.Raise 5, , "Duplicate metric definition: " & code
                result.Add code, False
                If Not IsNull(rs.fields(1).Value) Then result(code) = (UCase$(Trim$(CStr(rs.fields(1).Value))) = "Y")
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    conn.Close
    Set ReadDefinitions = result
    Exit Function
Failed:
    n = Err.number
    description = Err.description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not conn Is Nothing Then conn.Close
    On Error GoTo 0
    Err.Raise n, "Mod_ReportTemplate.ReadDefinitions", description
End Function

Private Function ReadRawCodes() As Object
    Dim config As clsConfig, rx As Object, matches As Object
    Dim result As Object, item As Object, alias As String, sql As String
    Set config = New clsConfig
    sql = config.ReadSQLFile(ChrW(&H8D22) & ChrW(&H62A5) & ".sql")
    Set rx = CreateObject("VBScript.RegExp")
    rx.IgnoreCase = True
    rx.Pattern = "\bFROM\s+jydb\."
    Set matches = rx.Execute(sql)
    If matches.count <> 1 Then Err.Raise 5, , "Expected one financial SQL FROM jydb clause."
    sql = Left$(sql, matches(0).FirstIndex)
    rx.Pattern = "\bAS\s+([A-Za-z_][A-Za-z0-9_]*)"
    rx.Global = True
    Set matches = rx.Execute(sql)
    Set result = CreateObject("Scripting.Dictionary")
    For Each item In matches
        alias = UCase$(item.SubMatches(0))
        Select Case alias
            Case "COMPANY_CODE", "END_DATE", "IF_MERGED", "IF_ADJUSTED", "INFO_PUBL_DATE"
            Case Else
                result(alias) = alias
        End Select
    Next item
    If result.count = 0 Then Err.Raise 5, , "Financial SQL contains no account aliases."
    result("TAX_RATE") = "TAX_RATE"
    Set ReadRawCodes = result
End Function

Private Sub ValidateCodes(ByVal codes As Variant, ByVal definitions As Object, ByVal raw As Object)
    Dim row As Long, code As String, unknown As String, disabled As String
    For row = 8 To 45
        If IsError(codes(row, 1)) Then Err.Raise 5, , "Template code contains an Excel error at row " & CStr(row)
        code = UCase$(Trim$(CStr(codes(row, 1))))
        If code <> "" Then
            If code = "TAX_RATE" Then
                ' Built-in version account, also supported when absent from SQL.
            ElseIf definitions.exists(code) Then
                If Not definitions(code) Then disabled = disabled & code & " (A" & CStr(row) & "); "
            ElseIf Not raw.exists(code) Then
                unknown = unknown & code & " (A" & CStr(row) & "); "
            End If
        End If
    Next row
    If unknown <> "" Or disabled <> "" Then Err.Raise 5, , "Template validation failed. Unknown/missing definitions: " & unknown & " Disabled: " & disabled & " Add approved enabled definitions to MeasuresDef; no business formulas are inferred."
End Sub
