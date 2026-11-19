Attribute VB_Name = "Mod_PlatformTests"
Option Explicit

' Edit these values to select one approved company and a narrow report-period window.
Public Sub TestFinancialCompanySample()
    Const COMPANY_CODE As String = "1833"
    Const REPORT_BEGIN As String = "2025-01-01"
    Const REPORT_END As String = "2025-12-31"
    Const KNOWLEDGE_AS_OF As String = "2026-11-16"
    Const IF_MERGED As Long = 1

    Dim config As clsConfig, manager As clsConnectionManager, connection As Object
    Dim builder As clsSqlBuilder, source As clsDbDataSource, loader As IDataSetSource
    Dim schema As clsDataSchema, dataset As clsDataSet, data As clsDataTable
    Dim aliasName As String, sqlText As String, row As Long, col As Long, Value As Variant

    On Error GoTo Failed
    Set config = New clsConfig
    If config.GetUseMock(config.GetSourceAliasByModule("FINANCIAL")) Then
        Err.Raise 5, "TestFinancialCompanySample", "Financial source is configured for mock mode."
    End If

    Set manager = New clsConnectionManager
    aliasName = config.GetSourceAliasByModule("FINANCIAL")
    Set connection = manager.GetConnection(aliasName)

    Set builder = New clsSqlBuilder
    Call builder.LoadTemplate(ChrW(&H8D22) & ChrW(&H62A5) & ".sql")
    Call builder.SetNumber("CompanyCode", COMPANY_CODE)
    Call builder.SetDate("ReportBegin", REPORT_BEGIN, False)
    Call builder.SetDate("ReportEnd", REPORT_END, False)
    Call builder.SetDate("AsOfDate", KNOWLEDGE_AS_OF, False)
    Call builder.SetNumber("IfMerged", IF_MERGED)
    sqlText = builder.Build()

    Set schema = NewFinancialWideSchema()
    Set source = New clsDbDataSource
    source.Initialize connection, sqlText, schema, aliasName, "FINANCIAL_SAMPLE_" & COMPANY_CODE, _
                      "SQL_TEMPLATE:FINANCIAL_REPORT:1", "KNOWN"
    Set loader = source
    Set dataset = loader.LoadData()
    Set data = dataset.data

    Debug.Print "[FIN-SAMPLE] PASS; source=" & aliasName & "; company=" & COMPANY_CODE & _
                "; period=" & REPORT_BEGIN & ".." & REPORT_END & "; asOf=" & KNOWLEDGE_AS_OF & _
                "; rows=" & CStr(data.rowCount) & "; columns=" & CStr(data.fieldCount)
    Debug.Print "[FIN-SAMPLE] Column names and first-row VBA types:"
    For col = 1 To data.fieldCount
        If data.rowCount > 0 Then
            Value = data.ValueAtColumn(1, col)
            Debug.Print "  " & data.fieldName(col) & "; VarType=" & CStr(VarType(Value)) & _
                        "; Null=" & CStr(IsNull(Value))
        Else
            Debug.Print "  " & data.fieldName(col) & "; no sample value"
        End If
    Next col
    For row = 1 To data.rowCount
        Debug.Print "[FIN-SAMPLE] identity row=" & CStr(row) & _
                    "; company=" & CStr(data.ValueAt(row, "COMPANY_CODE")) & _
                    "; end=" & Format$(CDate(data.ValueAt(row, "END_DATE")), "yyyy-mm-dd") & _
                    "; merged=" & CStr(data.ValueAt(row, "IF_MERGED")) & _
                    "; adjusted=" & CStr(data.ValueAt(row, "IF_ADJUSTED")) & _
                    "; published=" & Format$(CDate(data.ValueAt(row, "INFO_PUBL_DATE")), "yyyy-mm-dd") & _
                    "; total_assets=" & SampleValueText(data.ValueAt(row, "TOTAL_ASSETS")) & _
                    "; net_profit=" & SampleValueText(data.ValueAt(row, "NET_PROFIT"))
    Next row

CleanUp:
    On Error Resume Next
    If Not manager Is Nothing Then manager.CloseAll
    On Error GoTo 0
    Exit Sub

Failed:
    Debug.Print "[FIN-SAMPLE] FAIL; source=" & aliasName & "; number=" & CStr(Err.number) & _
                "; description=" & Err.description
    Resume CleanUp
End Sub

Private Function SampleValueText(ByVal Value As Variant) As String
    If IsNull(Value) Then
        SampleValueText = "<NULL>"
    ElseIf IsEmpty(Value) Then
        SampleValueText = "<EMPTY>"
    ElseIf IsError(Value) Then
        SampleValueText = "<ERROR>"
    Else
        SampleValueText = CStr(Value)
    End If
End Function

' Reads bounded live financial and portfolio samples into dedicated worksheets.
' Edit the constants before running; outputs overwrite only the named DBTEST sheets.
Public Sub TestRealDatabaseSamplesToSheets()
    Const COMPANY_CODE As String = "1833"
    Const REPORT_BEGIN As String = "2025-01-01"
    Const REPORT_END As String = "2025-12-31"
    Const KNOWLEDGE_AS_OF As String = "2026-11-16"
    Const IF_MERGED As Long = 1
    Const IF_ADJUSTED As Long = 2
    Const PORTFOLIO_CODE As String = "519008"
    Const PORTFOLIO_START As String = "2026-11-01"
    Const PORTFOLIO_END As String = "2026-11-30"

    Dim config As clsConfig, manager As clsConnectionManager, connection As Object
    Dim builder As clsSqlBuilder, source As clsDbDataSource, loader As IDataSetSource
    Dim schema As clsDataSchema, dataset As clsDataSet, financialData As clsDataTable
    Dim portfolioService As clsPortfolioService, panel As clsPortfolioPanel
    Dim aliasName As String, sqlText As String

    On Error GoTo Failed
    Set config = New clsConfig
    aliasName = config.GetSourceAliasByModule("FINANCIAL")
    If config.GetUseMock(aliasName) Then Err.Raise 5, "TestRealDatabaseSamplesToSheets", "Financial source is configured for mock mode."

    Set manager = New clsConnectionManager
    Set connection = manager.GetConnection(aliasName)

    Set builder = New clsSqlBuilder
    Call builder.LoadTemplate(ChrW(&H8D22) & ChrW(&H62A5) & ".sql")
    Call builder.SetNumber("CompanyCode", COMPANY_CODE)
    Call builder.SetDate("ReportBegin", REPORT_BEGIN, False)
    Call builder.SetDate("ReportEnd", REPORT_END, False)
    Call builder.SetDate("AsOfDate", KNOWLEDGE_AS_OF, False)
    Call builder.SetNumber("IfMerged", IF_MERGED)
    sqlText = builder.Build()

    Set schema = NewFinancialWideSchema()
    Set source = New clsDbDataSource
    source.Initialize connection, sqlText, schema, aliasName, "FINANCIAL_SAMPLE_" & COMPANY_CODE, _
                      "SQL_TEMPLATE:FINANCIAL_REPORT:1", "KNOWN"
    Set loader = source
    Set dataset = loader.LoadData()
    Set financialData = dataset.data
    WriteSampleTable "DBTEST_FINANCIAL", financialData.ToTable

    If config.GetUseMock(config.GetSourceAliasByModule("PORTFOLIO")) Then _
        Err.Raise 5, "TestRealDatabaseSamplesToSheets", "Portfolio source is configured for mock mode."
    Set portfolioService = New clsPortfolioService
    Set panel = portfolioService.GetPanel(Array(PORTFOLIO_CODE), CDate(PORTFOLIO_START), CDate(PORTFOLIO_END))
    WriteSampleTable "DBTEST_HOLDINGS", panel.HoldingsTable
    WriteSampleTable "DBTEST_PERFORMANCE", panel.PerformanceTable
    WriteSampleTable "DBTEST_PROFIT", panel.ProfitTable
    WriteSampleTable "DBTEST_ORDERS", panel.ordersTable
    WriteSampleTable "DBTEST_INDUSTRY", panel.industryTable

    Debug.Print "[REAL-SAMPLE] PASS; financial source=" & aliasName & "; rows=" & CStr(financialData.rowCount)
    Debug.Print "[REAL-SAMPLE] PASS; portfolio source=" & panel.sourceAlias & "; portfolio=" & PORTFOLIO_CODE & _
                "; date=" & PORTFOLIO_START & ".." & PORTFOLIO_END & "; holdings=" & CStr(panel.HoldingRowCount) & _
                "; performance=" & CStr(panel.PerformanceRowCount) & "; profit=" & CStr(panel.ProfitRowCount) & _
                "; orders=" & CStr(panel.OrderRowCount) & "; industry=" & CStr(panel.IndustryRowCount)
    Debug.Print "[REAL-SAMPLE] Data written to DBTEST_FINANCIAL and DBTEST_HOLDINGS/PERFORMANCE/PROFIT/ORDERS/INDUSTRY."

CleanUp:
    On Error Resume Next
    If Not manager Is Nothing Then manager.CloseAll
    ConnMgr.CloseAll
    On Error GoTo 0
    Exit Sub

Failed:
    Debug.Print "[REAL-SAMPLE] FAIL; " & SafeErrorText(Err.number, Err.description)
    Resume CleanUp
End Sub

Private Sub WriteSampleTable(ByVal sheetName As String, ByVal table As Variant)
    Dim ws As Worksheet, rowCount As Long, columnCount As Long
    rowCount = UBound(table, 1)
    columnCount = UBound(table, 2)
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = sheetName
    End If
    ws.Cells.ClearContents
    ws.Cells(1, 1).Resize(rowCount, columnCount).Value2 = table
    ws.rows(1).Font.Bold = True
    ws.rows(1).AutoFilter
    ws.columns.AutoFit
End Sub

' Run this Sub from the VBA Immediate window or Macro dialog in a trusted test workbook.
' It performs only SELECT-based connection and DataSet smoke tests. It does not print credentials.
Public Sub TestRealDatabaseConnections()
    Dim config As clsConfig, manager As clsConnectionManager, connection As Object
    Dim accessConnection As Object, source As clsDbDataSource, loader As IDataSetSource
    Dim dataset As clsDataSet, schema As clsDataSchema, successCount As Long, failureCount As Long
    Dim aliasName As String, ValueText As String

    On Error GoTo UnexpectedFailure
    Set config = New clsConfig
    Set manager = New clsConnectionManager

    aliasName = config.GetSourceAliasByModule("FINANCIAL")
    On Error GoTo OracleFailure
    Set connection = manager.GetConnection(aliasName)
    Set schema = NewConnectionSmokeSchema()
    Set source = New clsDbDataSource
    source.Initialize connection, "SELECT 1 AS CONNECTION_TEST FROM DUAL", schema, aliasName, "CONNECTION_SMOKE_ORACLE", "LIVE", "KNOWN"
    Set loader = source
    Set dataset = loader.LoadData()
    ValueText = CStr(dataset.data.ValueAt(1, "CONNECTION_TEST"))
    Debug.Print "[REAL-DB] Oracle PASS; source=" & aliasName & "; rows=" & CStr(dataset.data.rowCount) & "; value=" & ValueText
    successCount = successCount + 1
    GoTo AccessStart

OracleFailure:
    failureCount = failureCount + 1
    Debug.Print "[REAL-DB] Oracle FAIL; source=" & aliasName & "; " & SafeErrorText(Err.number, Err.description)
    Err.Clear

AccessStart:
    On Error GoTo AccessFailure
    Set accessConnection = CreateObject("ADODB.Connection")
    accessConnection.ConnectionTimeout = 15
    accessConnection.CommandTimeout = 120
    accessConnection.Open config.GetLocalConnString()
    Set schema = NewConnectionSmokeSchema()
    Set source = New clsDbDataSource
    source.Initialize accessConnection, "SELECT 1 AS CONNECTION_TEST", schema, config.LocalSourceAlias, "CONNECTION_SMOKE_ACCESS", "LIVE", "KNOWN"
    Set loader = source
    Set dataset = loader.LoadData()
    ValueText = CStr(dataset.data.ValueAt(1, "CONNECTION_TEST"))
    Debug.Print "[REAL-DB] Access PASS; source=" & config.LocalSourceAlias & "; rows=" & CStr(dataset.data.rowCount) & "; value=" & ValueText
    successCount = successCount + 1
    GoTo Finish

AccessFailure:
    failureCount = failureCount + 1
    Debug.Print "[REAL-DB] Access FAIL; source=" & config.LocalSourceAlias & "; " & SafeErrorText(Err.number, Err.description)

Finish:
    On Error Resume Next
    If Not accessConnection Is Nothing Then
        If accessConnection.State = 1 Then accessConnection.Close
    End If
    manager.CloseAll
    On Error GoTo 0
    Debug.Print "[REAL-DB] Summary; passed=" & CStr(successCount) & "; failed=" & CStr(failureCount)
    Exit Sub

UnexpectedFailure:
    Debug.Print "[REAL-DB] Test failure; " & SafeErrorText(Err.number, Err.description)
    Resume Finish
End Sub

Private Function NewConnectionSmokeSchema() As clsDataSchema
    Dim result As clsDataSchema
    Set result = New clsDataSchema
    result.Initialize "CONNECTION_SMOKE", "1", "CONNECTION_TEST"
    result.AddField "CONNECTION_TEST", "LONG", False
    result.Freeze
    Set NewConnectionSmokeSchema = result
End Function

Private Function SafeErrorText(ByVal errorNumber As Long, ByVal errorDescription As String) As String
    SafeErrorText = "number=" & CStr(errorNumber) & "; description=" & errorDescription
End Function

Public Function RunPlatformOfflineTests() As String
    Dim count As Long, sample As Variant, table As clsDataTable, schema As clsDataSchema
    Dim dataset As clsDataSet, context As clsAnalysisContext, request As clsAnalysisRequest
    Dim source As clsMemoryDataSource, loader As IDataSetSource, loaded As clsDataSet, schedule As clsDataTable
    Dim values() As Variant, model As clsFixedIncomeModel, result As clsAnalysisResult, output As clsDataTable
    Dim price As Double, dv01 As Double, shocked As Double, names As Variant, i As Long, errorNumber As Long
    Dim orderSchema As clsDataSchema, orderData As clsDataTable, orderValues() As Variant
    On Error GoTo Failed
    ReDim sample(1 To 3, 1 To 3)
    sample(1, 1) = "PORT_CODE": sample(1, 2) = "T_DATE": sample(1, 3) = "AMOUNT"
    sample(2, 1) = "000001": sample(2, 2) = DateSerial(2026, 11, 13): sample(2, 3) = CDec(0)
    sample(3, 1) = "000002": sample(3, 2) = DateSerial(2026, 11, 14): sample(3, 3) = Null
    Set table = New clsDataTable: table.Load sample
    Set schema = New clsDataSchema: schema.Initialize "SAMPLE", "1", "PORT_CODE+T_DATE"
    schema.AddField "PORT_CODE", "TEXT", False: schema.AddField "T_DATE", "DATE", False
    schema.AddField "AMOUNT", "DECIMAL": schema.Freeze
    Set dataset = New clsDataSet: dataset.Initialize table, schema, "MEMORY", "SAMPLE_QUERY", "FIXTURE-1", "KNOWN"
    Set source = New clsMemoryDataSource: source.Initialize dataset: Set loader = source: Set loaded = loader.LoadData()
    Call Check(loaded.data.rowCount = 2, count)
    Call Check(CStr(loaded.data.ValueAt(1, "PORT_CODE")) = "000001", count)
    Call Check(CDec(loaded.data.ValueAt(1, "AMOUNT")) = CDec(0), count)
    Call Check(IsNull(loaded.data.ValueAt(2, "AMOUNT")), count)
    Call Check(loaded.sourceAlias = "MEMORY" And loaded.revisionId = "FIXTURE-1", count)

    Dim dateValues() As Variant, dateTable As clsDataTable, dateSchema As clsDataSchema
    ReDim dateValues(1 To 2, 1 To 1)
    dateValues(1, 1) = "T_DATE": dateValues(2, 1) = "2026-11-01"
    Set dateTable = New clsDataTable: dateTable.Load dateValues
    Set dateSchema = New clsDataSchema: dateSchema.Initialize "DATE_COERCION", "1", "T_DATE"
    dateSchema.AddField "T_DATE", "DATE", False: dateSchema.Freeze
    Set dateTable = dateSchema.Normalize(dateTable)
    Check VarType(dateTable.ValueAt(1, "T_DATE")) = vbDate, count
    Check dateTable.ValueAt(1, "T_DATE") = DateSerial(2026, 11, 1), count

    Set source = New clsMemoryDataSource
    On Error Resume Next
    Err.Clear
    source.Initialize dataset
    source.Initialize dataset
    errorNumber = Err.number
    Err.Clear
    On Error GoTo Failed
    Call Check(errorNumber = 5, count)

    Set context = New clsAnalysisContext: context.Initialize DateSerial(2026, 11, 14), DateSerial(2026, 11, 14), "CNY", "BASELINE-1"
    Set request = New clsAnalysisRequest: request.Initialize "BOND_RATE_SCENARIO", context
    request.SetParameter "ANNUAL_RATE", 0.05
    Set result = New clsAnalysisResult: result.Initialize request
    result.AddSource loaded
    result.AddTable "SAMPLE", loaded.data
    Call Check(result.table("LINEAGE").rowCount = 1, count)
    Call Check(result.table("DIAGNOSTICS").rowCount = 0, count)
    Call Check(result.table("PARAMETERS").rowCount = 1, count)
    names = result.TableNames
    For i = LBound(names) To UBound(names)
        If names(i) = "DIAGNOSTICS" Then Exit For
    Next i
    Call Check(i <= UBound(names), count)
    result.Freeze

    Set orderSchema = NewPortfolioSchema("PORT_ORDERS")
    ReDim orderValues(1 To 2, 1 To 10)
    names = Array("PORT_CODE", "I_CODE", "M_TYPE", "I_NAME", "A_TYPE", "T_TYPE", "TRADE_ATTR", "ISSUER", "T_DATE", "T_AMOUNT")
    For i = 0 To UBound(names): orderValues(1, i + 1) = names(i): Next i
    orderValues(2, 1) = "000001": orderValues(2, 9) = DateSerial(2026, 11, 14): orderValues(2, 10) = CDec(12.5)
    Set orderData = New clsDataTable: orderData.Load orderValues
    Set orderData = orderSchema.Normalize(orderData)
    Call Check(orderData.ValueAt(1, "T_DATE") = DateSerial(2026, 11, 14), count)

    Set orderData = New clsDataTable: orderData.Load EmptyTableForPlatform(names)
    Set orderData = orderSchema.Normalize(orderData)
    Call Check(orderData.rowCount = 0 And orderData.HasField("T_DATE"), count)

    Set context = New clsAnalysisContext: context.Initialize DateSerial(2026, 11, 14), DateSerial(2026, 11, 14), "CNY", "BASELINE-1"
    Set request = New clsAnalysisRequest: request.Initialize "BOND_RATE_SCENARIO", context
    request.SetParameter "ANNUAL_RATE", 0.05
    request.Freeze
    ReDim values(1 To 3, 1 To 2)
    values(1, 1) = "PERIODS": values(1, 2) = "CASHFLOW"
    values(2, 1) = 1#: values(2, 2) = 5#
    values(3, 1) = 2#: values(3, 2) = 105#
    Set schedule = New clsDataTable: schedule.Load values
    Set model = New clsFixedIncomeModel
    price = model.price(schedule, 0.05)
    Call Check(Abs(price - 100#) < 0.0000001, count)
    dv01 = model.dv01(schedule, 0.05)
    Call Check(dv01 > 0# And Abs(dv01 - (model.price(schedule, 0.0499) - model.price(schedule, 0.0501)) / 2#) < 0.0000001, count)
    Set result = model.RunRateScenario(request, schedule, 1000000#, 100#)
    Set output = result.table("BOND_RATE_SCENARIO")
    shocked = CDbl(output.ValueAt(1, "PNL"))
    Call Check(shocked < 0# And Abs(CDbl(output.ValueAt(1, "BASE_PRICE")) - 100#) < 0.0000001, count)
    Call Check(result.table("DIAGNOSTICS").rowCount = 0, count)
    Call Check(result.table("LINEAGE").rowCount = 0, count)
    Call Check(result.status = "SUCCESS", count)
    Call Check(result.table("BOND_RATE_SCENARIO").rowCount = 1, count)
    On Error Resume Next
    Err.Clear
    result.AddDiagnostic "AFTER_FREEZE", "BOND", "TEST", "Must be rejected."
    errorNumber = Err.number
    Err.Clear
    On Error GoTo Failed
    Call Check(errorNumber = 5, count)

    RunPlatformOfflineTests = "PASS: " & CStr(count) & " platform assertions."
    Exit Function
Failed:
    Err.Raise Err.number, "Mod_PlatformTests", "Assertion " & CStr(count + 1) & " failed: " & Err.description
End Function

Public Function RunReaderOfflineTests() As String
    Dim count As Long
    On Error GoTo Failed
    RunReaderTests count
    RunDataSourceContractTests count
    RunReaderOfflineTests = "PASS: " & CStr(count) & " reader assertions."
    Exit Function
Failed:
    RunReaderOfflineTests = "FAIL: reader assertion " & CStr(count + 1) & "; " & CStr(Err.number) & "; " & Err.description
End Function

Public Function RunFinancialOfflineTests() As String
    Dim count As Long
    On Error GoTo Failed
    RunFinancialTests count
    RunFinancialOfflineTests = "PASS: " & CStr(count) & " financial assertions."
    Exit Function
Failed:
    RunFinancialOfflineTests = "FAIL: financial assertion " & CStr(count + 1) & "; " & CStr(Err.number) & "; " & Err.description
End Function

Private Sub RunReaderTests(ByRef count As Long)
    Dim rs As Object, reader As clsRecordsetReader, data As clsDataTable, stamp As Date
    Dim emptyData As clsDataTable, n As Long
    Set rs = CreateObject("ADODB.Recordset")
    rs.fields.Append "CODE", 200, 30
    rs.fields.Append "AMOUNT", 14, , 32
    rs.fields("AMOUNT").Precision = 28: rs.fields("AMOUNT").NumericScale = 6
    rs.fields.Append "OBSERVED_AT", 135
    rs.Open
    stamp = DateSerial(2026, 11, 14) + TimeSerial(15, 16, 17)
    rs.AddNew: rs.fields("CODE").Value = "000001234567890123"
    rs.fields("AMOUNT").Value = CDec("1234567890.1234"): rs.fields("OBSERVED_AT").Value = stamp: rs.Update
    rs.AddNew: rs.fields("CODE").Value = "000002": rs.fields("AMOUNT").Value = Null
    rs.fields("OBSERVED_AT").Value = stamp: rs.Update: rs.MoveFirst
    Set reader = New clsRecordsetReader: Set data = reader.Read(rs)
    Check data.rowCount = 2 And data.fieldCount = 3, count
    Check data.ValueAt(1, "CODE") = "000001234567890123", count
    Check VarType(data.ValueAt(1, "AMOUNT")) = vbDecimal, count
    Check data.ValueAt(1, "AMOUNT") = CDec("1234567890.1234"), count
    Check IsNull(data.ValueAt(2, "AMOUNT")), count
    Check VarType(data.ValueAt(1, "OBSERVED_AT")) = vbDate And data.ValueAt(1, "OBSERVED_AT") = stamp, count
    Check rs.State = 1 And rs.EOF, count
    Set emptyData = reader.Read(rs)
    Check emptyData.rowCount = 0 And emptyData.HasField("OBSERVED_AT"), count
    rs.Close
    Check data.ValueAt(1, "CODE") = "000001234567890123", count
    On Error Resume Next
    Err.Clear
    Set emptyData = reader.Read(rs)
    n = Err.number
    On Error GoTo 0
    Check n = 5, count
End Sub

Private Sub RunDataSourceContractTests(ByRef count As Long)
    Dim schema As clsDataSchema, data As clsDataTable, dataset As clsDataSet
    Dim values() As Variant, source As clsMemoryDataSource, loader As IDataSetSource, loaded As clsDataSet
    ReDim values(1 To 2, 1 To 1)
    values(1, 1) = "CODE": values(2, 1) = "000001"
    Set data = New clsDataTable: data.Load values
    Set schema = New clsDataSchema: schema.Initialize "SOURCE_CONTRACT", "1", "CODE"
    schema.AddField "CODE", "TEXT", False: schema.Freeze
    Set dataset = New clsDataSet: dataset.Initialize data, schema, "MEMORY", "CONTRACT", "REV-1", "KNOWN"
    Set source = New clsMemoryDataSource: source.Initialize dataset
    Set loader = source: Set loaded = loader.LoadData()
    Check loaded.data.ValueAt(1, "CODE") = "000001", count
    Check loaded.schema.datasetName = "SOURCE_CONTRACT", count
    Check loaded.sourceAlias = "MEMORY" And loaded.queryId = "CONTRACT", count
    Check loaded.revisionId = "REV-1" And loaded.AvailabilityEvidence = "KNOWN", count
End Sub

Private Sub RunFinancialTests(ByRef count As Long)
    Dim wide(1 To 5, 1 To 6) As Variant, fields As Variant, i As Long
    Dim data As clsDataTable, dataset As clsDataSet, schema As clsDataSchema
    Dim statement As clsFinancialStatement, view As clsFinancialStatement, adapter As clsWideFactAdapter
    Dim engine As clsMeasureEngine, definitions() As Variant
    Dim context As clsAnalysisContext, app As clsFinancialAnalysisApplication, result As clsAnalysisResult
    Dim ovTable(1 To 2, 1 To 7) As Variant, ovSchema As clsDataSchema, ovData As clsDataTable, ovSet As clsDataSet
    Dim errorNumber As Long

    fields = Array("COMPANY_CODE", "END_DATE", "IF_MERGED", "IF_ADJUSTED", "INFO_PUBL_DATE", "SALES")
    For i = 0 To UBound(fields): wide(1, i + 1) = fields(i): Next i
    For i = 2 To 5
        wide(i, 1) = "1001": wide(i, 3) = 1
        If i Mod 2 = 0 Then
            wide(i, 2) = DateSerial(2025, 12, 31): wide(i, 4) = 2
        Else
            wide(i, 2) = DateSerial(2024, 12, 31): wide(i, 4) = 1
        End If
        If i <= 3 Then wide(i, 5) = DateSerial(2026, 3, 30) Else wide(i, 5) = DateSerial(2026, 4, 30)
    Next i
    wide(2, 6) = 120#: wide(3, 6) = 100#: wide(4, 6) = 150#: wide(5, 6) = 125#
    Set data = New clsDataTable: data.Load wide
    Set schema = NewFinancialWideSchema()
    Set dataset = New clsDataSet: dataset.Initialize data, schema, "MEMORY", "FINANCIAL_FIXTURE", "FIN-1", "KNOWN"
    Set statement = New clsFinancialStatement: statement.companyCode = "1001"
    Set adapter = New clsWideFactAdapter: adapter.LoadData dataset.data, statement
    statement.AddSource dataset

    ' Verify override schema and multi-source lineage propagation.
    ovTable(1, 1) = "COMPANY_CODE": ovTable(1, 2) = "END_DATE": ovTable(1, 3) = "IF_MERGED"
    ovTable(1, 4) = "IF_ADJUSTED": ovTable(1, 5) = "INFO_PUBL_DATE": ovTable(1, 6) = "ACCOUNT_CODE": ovTable(1, 7) = "AMOUNT"
    ovTable(2, 1) = "1001": ovTable(2, 2) = "2025-12-31": ovTable(2, 3) = 1
    ovTable(2, 4) = 2: ovTable(2, 5) = "2026-03-30": ovTable(2, 6) = "PROVISION": ovTable(2, 7) = CDec(10)
    Set ovData = New clsDataTable: ovData.Load ovTable
    Set ovSchema = NewFinancialOverrideSchema()
    Set ovSet = New clsDataSet: ovSet.Initialize ovData, ovSchema, "LOCAL_ACC", "MANUAL_FINANCIALS", "OV-1", "KNOWN"
    statement.AddSource ovSet

    Set engine = New clsMeasureEngine
    ReDim definitions(0 To 5, 0 To 0)
    definitions(0, 0) = 1: definitions(1, 0) = "SALES_GROWTH": definitions(2, 0) = "GROWTH"
    definitions(3, 0) = "SALES": definitions(4, 0) = "": definitions(5, 0) = "Y"
    ReDim Preserve definitions(0 To 5, 0 To 1)
    definitions(0, 1) = 2: definitions(1, 1) = "PROVISION_CHECK": definitions(2, 1) = "DEFINED"
    definitions(3, 1) = "PROVISION": definitions(4, 1) = "": definitions(5, 1) = "Y"
    engine.LoadRows definitions: engine.ValidateDefinitions

    Set view = statement.CloneView(): view.ConfigureView 1, "2026-04-01"
    Set context = New clsAnalysisContext: context.Initialize DateSerial(2026, 4, 1), DateSerial(2026, 4, 1), "CNY", "FIN-POLICY-1"
    Set app = New clsFinancialAnalysisApplication
    Set result = app.RunMetricWithEngine(view, context, "2025", "SALES", 2, engine)
    Check result.table("FINANCIAL_METRIC").ValueAt(1, "VALUE") = 120#, count
    Check result.table("LINEAGE").rowCount = 2 And result.status = "SUCCESS", count

    Set result = app.RunMetricWithEngine(view, context, "2025", "PROVISION_CHECK", 2, engine)
    Check result.status = "SUCCESS" And result.table("FINANCIAL_METRIC").ValueAt(1, "VALUE") = 0#, count

    Set result = app.RunMetricWithEngine(view, context, "2025", "SALES_GROWTH", 2, engine)
    Check Abs(CDbl(result.table("FINANCIAL_METRIC").ValueAt(1, "VALUE")) - 0.2) < 0.0000001, count

    ' Audit and selected series unified entry points.
    Set result = app.RunAuditMetricWithEngine(view, context, "2025", "SALES_GROWTH", 2, engine)
    Check result.status = "SUCCESS" And result.table("METRIC_AUDIT").fieldCount = 14 And result.table("LINEAGE").rowCount = 2, count
    Set result = app.RunSelectedSeries(view, context, 2)
    Check result.status = "SUCCESS" And result.table("SELECTED_SERIES").fieldCount = 7 And result.table("LINEAGE").rowCount = 2, count

    ' Reject report period beyond valuationDate.
    On Error Resume Next
    Err.Clear
    Set result = app.RunMetricWithEngine(view, context, "2026", "SALES", 2, engine)
    errorNumber = Err.number
    Err.Clear
    On Error GoTo 0
    Check errorNumber = 5, count

    view.ConfigureView 1, "2026-05-01"
    Set context = New clsAnalysisContext: context.Initialize DateSerial(2026, 5, 1), DateSerial(2026, 5, 1), "CNY", "FIN-POLICY-1"
    Set result = app.RunMetricWithEngine(view, context, "2025", "SALES", 2, engine)
    Check result.table("FINANCIAL_METRIC").ValueAt(1, "VALUE") = 150#, count
    Set result = app.RunMetricWithEngine(view, context, "2025", "SALES_GROWTH", 2, engine)
    Check Abs(CDbl(result.table("FINANCIAL_METRIC").ValueAt(1, "VALUE")) - 0.2) < 0.0000001, count

    view.ConfigureView 1, "2026-03-29"
    Set context = New clsAnalysisContext: context.Initialize DateSerial(2026, 3, 29), DateSerial(2026, 3, 29), "CNY", "FIN-POLICY-1"
    Set result = app.RunMetricWithEngine(view, context, "2025", "SALES", 2, engine)
    Check result.status = "FAILED" And IsNull(result.table("FINANCIAL_METRIC").ValueAt(1, "VALUE")), count
    Check result.table("DIAGNOSTICS").rowCount = 1, count
End Sub

Private Function EmptyTableForPlatform(ByVal fieldNames As Variant) As Variant
    Dim names As Variant, values() As Variant, i As Long
    names = fieldNames
    ReDim values(1 To 1, 1 To UBound(names) - LBound(names) + 1)
    For i = LBound(names) To UBound(names): values(1, i - LBound(names) + 1) = names(i): Next i
    EmptyTableForPlatform = values
End Function

Private Sub Check(ByVal Condition As Boolean, ByRef count As Long)
    If Not Condition Then Err.Raise 5, "Mod_PlatformTests", "Expected condition was false."
    count = count + 1
End Sub

Public Sub EquityTest()

    Dim eqSvc As New clsEquityService
    Dim panel As clsEquityPanel
    Dim fullTs As clsTimeSeriesFrame, evalTs As clsTimeSeriesFrame
    Dim indTable As clsDataTable, sharpe As Variant, maxDd As Variant

    ' 1. 自动通过 DIM_SECURITY_MASTER 解析 InnerCode，并前置 400 天预热缓冲拉取后复权行情
    Set panel = eqSvc.GetQuotePanel(Array("600519", "000858"), DateSerial(2026, 1, 1), DateSerial(2026, 11, 15))

    ' 2. 提取贵州茅台（含预热期计算均线/MACD/RSI/布林带，输出时自动裁掉预热期）
    Set fullTs = panel.QuoteFrameForStock("600519", includeWarmup:=True)
    Set indTable = TA_BuildIndicatorTable(fullTs, displayStartDate:=panel.startDate)

    ' 3. 截取目标评估期，直接复用 Mod_TimeSeriesOps 计算夏普比率与最大回撤
    Set evalTs = panel.QuoteFrameForStock("600519", includeWarmup:=False)
    sharpe = TS_SharpeRatio(evalTs, "DAILY_RTN")
    maxDd = TS_MaxDrawdown(evalTs, "DAILY_RTN")

End Sub
