Attribute VB_Name = "Mod_EquityMain"
Option Explicit

Public Sub TestExportEquityQuotes()
    Dim eqSvc As clsEquityService
    Dim panel As clsEquityPanel
    Dim fullTs As clsTimeSeriesFrame
    Dim evalTs As clsTimeSeriesFrame
    Dim indTable As clsDataTable
    Dim sharpe As Variant, maxDd As Variant
    Dim t0 As Double

    On Error GoTo Failed
    t0 = Timer
    Application.ScreenUpdating = False

    ' 1. 拉取行情面板（默认自动前置 400 天预热期用于均线/MACD计算）
    Set eqSvc = New clsEquityService
    Set panel = eqSvc.GetQuotePanel(Array("600519", "000858"), DateSerial(2026, 1, 1), DateSerial(2026, 11, 15))

    ' 2. 导出匹配到的主数据映射表（核对 INNER_CODE、COMPANY_CODE 是否正确）
    WriteDataTableToSheet "EQ_MASTER", panel.MasterTable

    ' 3. 导出全量拉取的后复权行情表（核对 CLOSE_PRICE、ADJ_FACTOR、ADJ_CLOSE、DAILY_RTN）
    '    注：此处导出 panel.QuoteFrame 包含预热期；若只想看目标区间，可改为导出单个标的的 evalTs
    WriteTimeSeriesToSheet "EQ_QUOTES", panel.quoteFrame

    ' 4. 导出单只股票（600519）的技术指标明细表（自动裁掉预热期，仅保留 2026-01-01 起）
    Set fullTs = panel.QuoteFrameForStock("600519", includeWarmup:=True)
    Set indTable = TA_BuildIndicatorTable(fullTs, displayStartDate:=panel.startDate)
    WriteDataTableToSheet "EQ_TECH_600519", indTable

    ' 5. 顺便计算目标区间内的夏普比率与最大回撤并弹窗提示
    Set evalTs = panel.QuoteFrameForStock("600519", includeWarmup:=False)
    sharpe = TS_SharpeRatio(evalTs, "DAILY_RTN")
    maxDd = TS_MaxDrawdown(evalTs, "DAILY_RTN")

    Application.ScreenUpdating = True
    MsgBox "行情与技术指标导出完成！耗时: " & Format$(Timer - t0, "0.00") & " 秒" & vbCrLf & _
           "主数据行数: " & panel.MasterTable.rowCount & vbCrLf & _
           "总行情行数(含预热): " & panel.quoteFrame.rowCount & vbCrLf & _
           "600519 区间夏普比率: " & IIf(IsNull(sharpe), "N/A", Format$(CDbl(sharpe), "0.0000")) & vbCrLf & _
           "600519 区间最大回撤: " & IIf(IsNull(maxDd), "N/A", Format$(CDbl(maxDd), "0.00%")), _
           vbInformation, "股票行情导出核对"
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    MsgBox "行情导出失败: " & Err.description, vbExclamation, "股票行情导出核对"
End Sub

' ==============================================================================
' 通用工作表输出辅助过程（自动保留代码前导零并设置日期/数值格式）
' ==============================================================================

Public Sub WriteDataTableToSheet(ByVal sheetName As String, ByVal tbl As clsDataTable)
    Dim arr() As Variant, r As Long, c As Long, rCount As Long, cCount As Long
    If tbl Is Nothing Or Not tbl.IsLoaded Then Exit Sub
    rCount = tbl.rowCount
    cCount = tbl.fieldCount
    ReDim arr(1 To rCount + 1, 1 To cCount)
    For c = 1 To cCount
        arr(1, c) = tbl.fieldName(c)
    Next c
    For r = 1 To rCount
        For c = 1 To cCount
            arr(r + 1, c) = tbl.ValueAtColumn(r, c)
        Next c
    Next r
    WriteVariantArrayToSheet sheetName, arr
End Sub

Public Sub WriteTimeSeriesToSheet(ByVal sheetName As String, ByVal ts As clsTimeSeriesFrame)
    If ts Is Nothing Or Not ts.IsLoaded Then Exit Sub
    WriteVariantArrayToSheet sheetName, ts.ToTable()
End Sub

Private Sub WriteVariantArrayToSheet(ByVal sheetName As String, ByVal dataArr As Variant)
    Dim ws As Worksheet, rCount As Long, cCount As Long, c As Long, colName As String
    rCount = UBound(dataArr, 1)
    cCount = UBound(dataArr, 2)

    Set ws = GetOrCreateWorksheet(ThisWorkbook, sheetName)
    ws.Cells.Clear

    ' 先按表头预设列格式（确保文本代码保留前导零），再整块写入二维数组
    If rCount > 1 Then
        For c = 1 To cCount
            colName = UCase$(Trim$(CStr(dataArr(1, c))))
            Select Case colName
                Case "SEC_CODE", "COMPANY_CODE", "MARKET_CODE", "ASSET_CLASS", "SUB_TYPE", "SOURCE_TABLE"
                    ws.Range(ws.Cells(2, c), ws.Cells(rCount, c)).NumberFormat = "@"
                Case "T_DATE", "LIST_DATE", "DELIST_DATE", "UPDATE_TIME"
                    ws.Range(ws.Cells(2, c), ws.Cells(rCount, c)).NumberFormat = "yyyy-mm-dd"
                Case "DAILY_RTN", "TURNOVER_RATE"
                    ws.Range(ws.Cells(2, c), ws.Cells(rCount, c)).NumberFormat = "0.00%"
                Case "ADJ_FACTOR"
                    ws.Range(ws.Cells(2, c), ws.Cells(rCount, c)).NumberFormat = "0.000000"
                Case "VOLUME", "AMOUNT"
                    ws.Range(ws.Cells(2, c), ws.Cells(rCount, c)).NumberFormat = "#,##0.00"
                Case "INNER_CODE"
                    ws.Range(ws.Cells(2, c), ws.Cells(rCount, c)).NumberFormat = "0"
                Case Else
                    ws.Range(ws.Cells(2, c), ws.Cells(rCount, c)).NumberFormat = "#,##0.0000"
            End Select
        Next c
    End If

    ws.Range(ws.Cells(1, 1), ws.Cells(rCount, cCount)).Value = dataArr

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, cCount))
        .Font.Bold = True
        .Interior.Color = RGB(237, 242, 248)
    End With
    ws.Range(ws.Cells(1, 1), ws.Cells(rCount, cCount)).AutoFilter
    ws.columns.AutoFit
End Sub

Private Function GetOrCreateWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        ws.name = sheetName
    End If
    Set GetOrCreateWorksheet = ws
End Function

Public Sub ExportAnnualAdjustedValuationReport()
    Dim eqSvc As clsEquityService
    Dim panel As clsEquityPanel
    Dim analyzer As clsEquityValuationAnalyzer
    Dim codes As Variant, codeItem As Variant
    Dim dailyTbl As clsDataTable, seriesList As New Collection, summaryTbl As clsDataTable
    Dim t0 As Double

    On Error GoTo Failed
    t0 = Timer
    Application.ScreenUpdating = False

    codes = Array("600519", "000858")

    ' 1. 加载行情与总市值面板（仅算估值倍数无需均线预热，lookbackDays 设为 0）
    Set eqSvc = New clsEquityService
    Set panel = eqSvc.GetQuotePanel(codes, DateSerial(2026, 1, 1), DateSerial(2026, 11, 15), lookbackDays:=0)

    ' 2. 初始化分析器（如果 STK_DailyQuote 中 TOTAL_MV 单位为万元，mvUnitMultiplier 设为 10000#；若已换算为元，改为 1#）
    Set analyzer = New clsEquityValuationAnalyzer
    analyzer.Initialize ifAdjusted:=2, mvUnitMultiplier:=1#

    ' 3. 逐股计算日频 ADJ_PE = TOTAL_MV / 最近一期已披露年报 NOPAT
    For Each codeItem In codes
        Set dailyTbl = analyzer.BuildDailyLfyRatioSeries( _
            panel:=panel, _
            secCodeOrInnerCode:=CStr(codeItem), _
            metricCode:="NOPAT", _
            ratioName:="ADJ_PE", _
            useFloatMv:=False)
        seriesList.Add dailyTbl
        WriteDataTableToSheet "ADJ_PE_LFY_" & CStr(codeItem), dailyTbl
    Next codeItem

    ' 4. 输出多股票横向对比与历史分位表
    Set summaryTbl = analyzer.BuildValuationSummary(seriesList, ratioName:="ADJ_PE")
    WriteDataTableToSheet "EQ_ADJ_PE_LFY_SUMMARY", summaryTbl

    Application.ScreenUpdating = True
    MsgBox "基于最近一期年报 (LFY) 的调整PE联动计算完成！" & vbCrLf & _
           "耗时: " & Format$(Timer - t0, "0.00") & " 秒", vbInformation, "年报调整PE估值分析"
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    MsgBox "年报调整PE计算失败: " & Err.description, vbExclamation, "年报调整PE估值分析"
End Sub
