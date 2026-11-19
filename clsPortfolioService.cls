VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPortfolioService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const SQL_HOLDINGS As String = "PORT_Holdings.sql"
Private Const SQL_PERFORMANCE As String = "PORT_Measures.sql"
Private Const SQL_PROFIT As String = "PORT_PLDetail.sql"
Private Const SQL_ORDERS As String = "PORT_Orders.sql"
Private Const SQL_INDUSTRY As String = "PORT_StockIndustry.sql"
Private Const PRIOR_NAV_BUFFER_DAYS As Long = 15

Public Function GetPanel(ByVal portCodes As Variant, _
        ByVal startDate As Date, _
        ByVal endDate As Date) As clsPortfolioPanel

    Set GetPanel = BuildPanel(portCodes, startDate, endDate, True, True)
End Function

Public Function GetReturnPanel(ByVal portCodes As Variant, ByVal endDate As Date) As clsPortfolioPanel
    Set GetReturnPanel = BuildReturnPanel(portCodes, endDate)
End Function

Private Function BuildReturnPanel(ByVal portCodes As Variant, ByVal endDate As Date) As clsPortfolioPanel
    Dim config As clsConfig, connection As Object, performanceData As clsDataSet
    Dim sources As New Collection
    Dim source As String, codeList As String
    Dim panel As clsPortfolioPanel, holdings As clsTimeSeriesFrame, performance As clsTimeSeriesFrame
    Dim profit As clsTimeSeriesFrame, orders As clsDataTable, industry As clsDataTable
    Dim number As Long, description As String, startDate As Date

    startDate = DateSerial(1900, 1, 1)
    If dateValue(startDate) > dateValue(endDate) Then Err.Raise 5, , "Panel start date is after end date."
    codeList = BuildPortCodeList(portCodes)
    Set config = New clsConfig
    source = config.GetSourceAliasByModule("PORTFOLIO")
    If Len(Trim$(source)) = 0 Then Err.Raise 5, , "Missing PORTFOLIO source."
    On Error GoTo Failed
    Set connection = ConnMgr().GetConnection(source)
    Set performanceData = ExecuteQueryEndOnly(connection, source, SQL_PERFORMANCE, portCodes, endDate)
    sources.Add performanceData
    Set performance = ReadTimeSeries(performanceData, "T_DATE")
    Set holdings = EmptyTimeSeries(Array("PORT_CODE", "T_DATE", "I_CODE", "H_EVAL", "H_AI", "H_NAV_PCT", "A_TYPE"), "T_DATE")
    Set profit = EmptyTimeSeries(Array("PORT_CODE", "T_DATE", "I_CODE", "DAILYPROFIT"), "T_DATE")
    Set orders = EmptyDataTable(Array("PORT_CODE", "T_DATE", "ORDER_ID"))
    Set industry = EmptyDataTable(Array("SEC_CODE", "START_DATE", "EXPIRATION_DATE", "SW_SECTOR_CODE", "SW_SECTOR"))
    Set panel = New clsPortfolioPanel
    panel.Load codeList, startDate, endDate, source, holdings, performance, profit, orders, industry, sources
    Set BuildReturnPanel = panel
    Exit Function
Failed:
    number = Err.number: description = Err.description
    Err.Raise number, "clsPortfolioService.BuildReturnPanel", "Annual return data load failed. portfolios=" & codeList & "; end=" & Format$(endDate, "yyyy-mm-dd") & "; detail=" & description
End Function

Private Function BuildPanel(ByVal portCodes As Variant, ByVal startDate As Date, ByVal endDate As Date, ByVal includePositionData As Boolean, ByVal includePerformance As Boolean) As clsPortfolioPanel
    Dim config As clsConfig, connection As Object, profitConnection As Object, industryConnection As Object
    Dim source As String, profitSource As String, industrySource As String, codeList As String
    Dim mainBoardCodes As Object, stibCodes As Object, perfStartDate As Date
    Dim panel As clsPortfolioPanel, holdings As clsTimeSeriesFrame, performance As clsTimeSeriesFrame
    Dim profit As clsTimeSeriesFrame, orders As clsDataTable, industry As clsDataTable
    Dim holdingsData As clsDataSet, performanceData As clsDataSet, profitData As clsDataSet
    Dim ordersData As clsDataSet, industryData As clsDataSet, sources As New Collection
    Dim number As Long, description As String

    If dateValue(startDate) > dateValue(endDate) Then Err.Raise 5, , "Panel start date is after end date."
    codeList = BuildPortCodeList(portCodes)

    Set config = New clsConfig
    source = config.GetSourceAliasByModule("PORTFOLIO")
    If includePositionData Then profitSource = config.GetSourceAliasByRoutingKey("PORTFOLIO_PL")
    If includePositionData Then industrySource = config.GetSourceAliasByModule("FINANCIAL")
    If Len(Trim$(source)) = 0 Then Err.Raise 5, , "Missing PORTFOLIO source."
    If includePositionData And Len(Trim$(profitSource)) = 0 Then Err.Raise 5, , "Missing PORTFOLIO_PL source."
    If includePositionData And Len(Trim$(industrySource)) = 0 Then Err.Raise 5, , "Missing FINANCIAL source."
    On Error GoTo Failed
    Set connection = ConnMgr().GetConnection(source)
    If includePositionData Then Set profitConnection = ConnMgr().GetConnection(profitSource)
    If includePositionData Then Set industryConnection = ConnMgr().GetConnection(industrySource)

    If includePositionData Then
        Set holdingsData = ExecuteQuery(connection, source, SQL_HOLDINGS, portCodes, startDate, endDate)
        sources.Add holdingsData
        Set holdings = ReadTimeSeries(holdingsData, "T_DATE")
        BuildIndustryLists holdings, mainBoardCodes, stibCodes
        Set profitData = ExecuteQuery(profitConnection, profitSource, SQL_PROFIT, portCodes, startDate, endDate)
        sources.Add profitData
        Set profit = ReadTimeSeries(profitData, "T_DATE")
        Set ordersData = ExecuteQuery(connection, source, SQL_ORDERS, portCodes, startDate, endDate)
        sources.Add ordersData
        Set orders = ordersData.data
        Set industryData = ExecuteIndustryQuery(industryConnection, industrySource, mainBoardCodes, stibCodes)
        sources.Add industryData
        Set industry = industryData.data
    Else
        Set holdings = EmptyTimeSeries(Array("PORT_CODE", "T_DATE", "I_CODE", "H_EVAL", "H_AI", "H_NAV_PCT", "A_TYPE"), "T_DATE")
        Set profit = EmptyTimeSeries(Array("PORT_CODE", "T_DATE", "I_CODE", "DAILYPROFIT"), "T_DATE")
        Set orders = EmptyDataTable(Array("PORT_CODE", "T_DATE", "ORDER_ID"))
        Set industry = EmptyDataTable(Array("SEC_CODE", "START_DATE", "EXPIRATION_DATE", "SW_SECTOR_CODE", "SW_SECTOR"))
    End If
    If includePerformance Then
        If includePositionData Then
            ' Pull a 15-day lookback buffer so day-1 attribution can resolve prior trading day NAV.
            perfStartDate = DateAdd("d", -PRIOR_NAV_BUFFER_DAYS, dateValue(startDate))
            Set performanceData = ExecuteQuery(connection, source, SQL_PERFORMANCE, portCodes, perfStartDate, endDate)
        Else
            Set performanceData = ExecuteQueryEndOnly(connection, source, SQL_PERFORMANCE, portCodes, endDate)
        End If
        sources.Add performanceData
        Set performance = ReadTimeSeries(performanceData, "T_DATE")
    Else
        Set performance = EmptyTimeSeries(Array("PORT_CODE", "T_DATE", "P_TOTALNAV", "RTN", "RTN_BM"), "T_DATE")
    End If

    Set panel = New clsPortfolioPanel
    panel.Load codeList, startDate, endDate, source, holdings, performance, profit, orders, industry, sources
    Set BuildPanel = panel
    Exit Function
Failed:
    number = Err.number: description = Err.description
    Err.Raise number, "clsPortfolioService.BuildPanel", "Portfolio panel load failed. portfolios=" & codeList & "; start=" & Format$(startDate, "yyyy-mm-dd") & "; end=" & Format$(endDate, "yyyy-mm-dd") & "; detail=" & description
End Function

Private Function EmptyTimeSeries(ByVal fieldNames As Variant, ByVal DateField As String) As clsTimeSeriesFrame
    Dim result As clsTimeSeriesFrame
    Set result = New clsTimeSeriesFrame
    result.Load emptyTable(fieldNames), DateField
    Set EmptyTimeSeries = result
End Function

Private Function EmptyDataTable(ByVal fieldNames As Variant) As clsDataTable
    Dim result As clsDataTable
    Set result = New clsDataTable
    result.Load emptyTable(fieldNames)
    Set EmptyDataTable = result
End Function

Private Function emptyTable(ByVal fieldNames As Variant) As Variant
    Dim names As Variant, result() As Variant, i As Long
    names = fieldNames
    ReDim result(1 To 1, 1 To UBound(names) - LBound(names) + 1)
    For i = LBound(names) To UBound(names): result(1, i - LBound(names) + 1) = UCase$(CStr(names(i))): Next i
    emptyTable = result
End Function

Private Function ReadTimeSeries(ByVal dataset As clsDataSet, ByVal DateField As String) As clsTimeSeriesFrame
    Dim result As clsTimeSeriesFrame
    Set result = New clsTimeSeriesFrame
    result.LoadData dataset.data, DateField
    Set ReadTimeSeries = result
End Function

Private Function BuildPortCodeList(ByVal portCodes As Variant) As String
    Dim builder As clsSqlBuilder
    Set builder = New clsSqlBuilder
    Call builder.LoadRawSql("{PortList}", "Portfolio code list")
    Call builder.SetInList("PortList", portCodes)
    ' Retain the quoted PortfolioCodes display contract used by panel reports.
    BuildPortCodeList = builder.Build()
End Function

Private Sub BuildIndustryLists(ByVal holdings As clsTimeSeriesFrame, _
        ByRef mainBoardCodes As Object, _
        ByRef stibCodes As Object)

    Dim r As Long, Value As Variant, market As String, code As String

    Set mainBoardCodes = CreateObject("Scripting.Dictionary")
    Set stibCodes = CreateObject("Scripting.Dictionary")
    mainBoardCodes.CompareMode = 1
    stibCodes.CompareMode = 1
    If Not holdings.HasField("I_CODE") Or Not holdings.HasField("M_TYPE") Then Exit Sub
    For r = 1 To holdings.rowCount
        Value = holdings.ValueAt(r, "I_CODE")
        If Not IsNull(Value) Then
            code = Trim$(CStr(Value))
            If Len(code) > 0 Then
                Value = holdings.ValueAt(r, "M_TYPE")
                If Not IsNull(Value) Then
                    market = CStr(Value)
                    If IsMainlandMarket(market) Then
                        If Left$(code, 3) = "688" Then
                            If Not stibCodes.exists(code) Then stibCodes.Add code, code
                        ElseIf Not mainBoardCodes.exists(code) Then
                            mainBoardCodes.Add code, code
                        End If
                    End If
                End If
            End If
        End If
    Next r
End Sub

Private Function IsMainlandMarket(ByVal marketType As String) As Boolean
    Dim Value As String, shLabel As String, szLabel As String, shszLabel As String, shszMktLabel As String
    Value = UCase$(Trim$(marketType))
    shLabel = ChrW(&H6CAA) & ChrW(&H5E02)
    szLabel = ChrW(&H6DF1) & ChrW(&H5E02)
    shszLabel = ChrW(&H6CAA) & ChrW(&H6DF1)
    shszMktLabel = ChrW(&H6CAA) & ChrW(&H6DF1) & ChrW(&H5E02) & ChrW(&H573A)
    IsMainlandMarket = (Value = "CN" Or Value = "A" Or Value = "ASHARE" Or Value = "A_SHARE" Or _
        Value = "SH" Or Value = "SZ" Or Value = "SSE" Or Value = "SZSE" Or _
        Value = "XSHE" Or Value = "XSHG" Or Value = shLabel Or Value = szLabel Or _
        Value = shszLabel Or Value = shszMktLabel Or Value = "83" Or Value = "90")
End Function

Private Function ExecuteQuery(ByVal connection As Object, ByVal sourceAlias As String, ByVal sqlFileName As String, ByVal portCodes As Variant, ByVal startDate As Date, ByVal endDate As Date) As clsDataSet
    Dim builder As clsSqlBuilder, executor As clsDbExecutor, schema As clsDataSchema
    Dim datasetName As String, data As clsDataSet
    Set builder = New clsSqlBuilder
    Call builder.LoadTemplate(sqlFileName)
    Call builder.SetInList("PortList", portCodes)
    Call builder.SetDate("StartDate", startDate)
    Call builder.SetDate("EndDate", endDate)
    Select Case UCase$(sqlFileName)
        Case UCase$(SQL_HOLDINGS): datasetName = "PORT_HOLDINGS"
        Case UCase$(SQL_PERFORMANCE): datasetName = "PORT_PERFORMANCE"
        Case UCase$(SQL_PROFIT): datasetName = "PORT_PROFIT"
        Case UCase$(SQL_ORDERS): datasetName = "PORT_ORDERS"
        Case Else: Err.Raise 5, "clsPortfolioService.ExecuteQuery", "Unregistered portfolio dataset: " & sqlFileName
    End Select
    Set schema = NewPortfolioSchema(datasetName)
    Set executor = New clsDbExecutor
    Set data = executor.ReadDataSet(connection, builder.Build(), schema, sourceAlias, sqlFileName)
    Set ExecuteQuery = data
End Function

Private Function ExecuteQueryEndOnly(ByVal connection As Object, ByVal sourceAlias As String, ByVal sqlFileName As String, ByVal portCodes As Variant, ByVal endDate As Date) As clsDataSet
    Set ExecuteQueryEndOnly = ExecuteQuery(connection, sourceAlias, sqlFileName, portCodes, DateSerial(1900, 1, 1), endDate)
End Function

Private Function ExecuteIndustryQuery(ByVal connection As Object, ByVal sourceAlias As String, ByVal mainBoardCodes As Object, ByVal stibCodes As Object) As clsDataSet
    Dim sqlText As String, builder As clsSqlBuilder, executor As clsDbExecutor
    Dim schema As clsDataSchema, data As clsDataSet
    Set builder = New clsSqlBuilder
    Call builder.LoadTemplate(SQL_INDUSTRY)
    Call builder.SetChunkedInPredicate("MainBoardPredicate", "b.AStockCode", mainBoardCodes)
    Call builder.SetChunkedInPredicate("StibPredicate", "b.SecuCode", stibCodes)
    sqlText = builder.Build()
    Set schema = NewPortfolioSchema("PORT_INDUSTRY")
    Set executor = New clsDbExecutor
    Set data = executor.ReadDataSet(connection, sqlText, schema, sourceAlias, SQL_INDUSTRY)
    Set ExecuteIndustryQuery = data
End Function
