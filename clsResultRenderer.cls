VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsResultRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' clsResultRenderer: 统一结果呈现与报表交付引擎
' 负责解耦分析计算与交付介质，支持四种交付模式：
'   1. DELIV_DATA_ONLY (0): 纯内存对象，直接返回 clsAnalysisResult
'   2. DELIV_UDF_SCALAR (1): 单元格公式求值，返回标量或 CVErr
'   3. DELIV_CURRENT_WORKBOOK (2): 在当前打开的工作簿中追加/刷新指定名称工作表
'   4. DELIV_NEW_WORKBOOK (3): 独立生成全新、带专业排版的报告工作簿
' ==============================================================================

Public Function Render(ByVal result As clsAnalysisResult, _
                       Optional ByVal deliveryMode As Long = 2, _
                       Optional ByVal targetTitle As String = "ANALYSIS_REPORT") As Variant
    If result Is Nothing Then Err.Raise 5, "clsResultRenderer", "Result required."

    Select Case deliveryMode
        Case DELIV_DATA_ONLY
            Set Render = result

        Case DELIV_UDF_SCALAR
            Render = ExtractScalar(result)

        Case DELIV_CURRENT_WORKBOOK
            Set Render = RenderToCurrentWorkbook(result, targetTitle)

        Case DELIV_NEW_WORKBOOK
            Set Render = RenderToNewWorkbook(result, targetTitle)

        Case Else
            Err.Raise 5, "clsResultRenderer", "Invalid delivery mode: " & CStr(deliveryMode)
    End Select
End Function

Public Function RenderToCurrentWorkbook(ByVal result As clsAnalysisResult, ByVal prefix As String) As Object
    Dim wb As Object, ws As Object, tblName As Variant
    Dim primaryTblName As String, hasRenderedPrimary As Boolean
    Dim cleanPrefix As String

    If result Is Nothing Then Err.Raise 5, "clsResultRenderer", "Result required."
    cleanPrefix = CleanSheetName(prefix)
    If Len(cleanPrefix) = 0 Then cleanPrefix = "REPORT"

    On Error Resume Next
    Set wb = Application.ActiveWorkbook
    On Error GoTo 0
    If wb Is Nothing Then
        Set wb = Application.Workbooks.Add
    End If

    ' 找到主业务表（排除保留表 DIAGNOSTICS/LINEAGE/PARAMETERS）
    primaryTblName = FindPrimaryTableName(result)

    ' 1. 渲染主表到以 prefix 命名的工作表
    Set ws = GetOrCreateWorksheet(wb, cleanPrefix)
    RenderSingleTableToSheet result.table(primaryTblName), ws, cleanPrefix & " (" & primaryTblName & ")"
    Set RenderToCurrentWorkbook = ws

    ' 2. 如果存在警告或错误，自动追加诊断工作表
    If result.status <> "SUCCESS" Then
        Dim wsDiag As Object
        Set wsDiag = GetOrCreateWorksheet(wb, Left$(cleanPrefix & "_Diag", 31))
        RenderSingleTableToSheet result.table("DIAGNOSTICS"), wsDiag, "DIAGNOSTICS (" & result.status & ")"
    End If
End Function

Public Function RenderToNewWorkbook(ByVal result As clsAnalysisResult, ByVal reportTitle As String) As Object
    Dim wb As Object, ws As Object, tblName As Variant
    Dim names As Variant, i As Long, sheetName As String
    Dim isFirst As Boolean

    If result Is Nothing Then Err.Raise 5, "clsResultRenderer", "Result required."
    Set wb = Application.Workbooks.Add(-4167) ' xlWBATWorksheet

    names = result.TableNames
    isFirst = True

    For i = LBound(names) To UBound(names)
        tblName = CStr(names(i))
        ' 仅在有内容时渲染诊断与谱系表
        If tblName = "DIAGNOSTICS" And result.Diagnostics.count = 0 Then GoTo NextTable
        If tblName = "LINEAGE" And result.Lineage.count = 0 Then GoTo NextTable

        If isFirst Then
            Set ws = wb.Worksheets(1)
            isFirst = False
        Else
            Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        End If

        sheetName = CleanSheetName(CStr(tblName))
        ws.name = sheetName
        RenderSingleTableToSheet result.table(CStr(tblName)), ws, reportTitle & " - " & CStr(tblName)
NextTable:
    Next i

    Set RenderToNewWorkbook = wb
End Function

Public Sub RenderSingleTableToSheet(ByVal tbl As clsDataTable, ByVal ws As Object, ByVal title As String)
    Dim arr() As Variant, r As Long, c As Long, rCount As Long, cCount As Long
    Dim colName As String

    If tbl Is Nothing Or ws Is Nothing Then Exit Sub
    If Not tbl.IsLoaded Then Exit Sub

    rCount = tbl.rowCount
    cCount = tbl.fieldCount
    If rCount = 0 Or cCount = 0 Then Exit Sub

    ws.Cells.Clear

    ReDim arr(1 To rCount + 1, 1 To cCount)
    For c = 1 To cCount
        arr(1, c) = tbl.fieldName(c)
    Next c
    For r = 1 To rCount
        For c = 1 To cCount
            arr(r + 1, c) = tbl.ValueAtColumn(r, c)
            If IsNull(arr(r + 1, c)) Then arr(r + 1, c) = Empty
        Next c
    Next r

    ' 设置列格式
    For c = 1 To cCount
        colName = UCase$(Trim$(CStr(arr(1, c))))
        Select Case colName
            Case "SEC_CODE", "COMPANY_CODE", "INNER_CODE", "PORT_CODE", "MARKET_CODE", "STATUS", "SECTOR_CODE"
                ws.Range(ws.Cells(2, c), ws.Cells(rCount + 1, c)).NumberFormat = "@"
            Case "T_DATE", "PUB_DATE", "REPORT_DATE", "LIST_DATE", "UPDATE_TIME", "FISCAL_YEAR_END"
                ws.Range(ws.Cells(2, c), ws.Cells(rCount + 1, c)).NumberFormat = "yyyy-mm-dd"
            Case "DAILY_RTN", "RTN", "WEIGHT", "H_NAV_PCT", "ALLOC_PCT", "SELECTION_PCT", "TOTAL_PCT", "TURNOVER_RATE", "PERCENTILE"
                ws.Range(ws.Cells(2, c), ws.Cells(rCount + 1, c)).NumberFormat = "0.00%"
            Case "PE", "PB", "PS", "EV_EBITDA", "SHARPE", "MAX_DD", "VOLATILITY", "ANNUAL_RTN"
                ws.Range(ws.Cells(2, c), ws.Cells(rCount + 1, c)).NumberFormat = "0.0000"
            Case "VOLUME", "AMOUNT", "MV_CNY", "TOTAL_MV", "FLOAT_MV", "H_EVAL", "DAILYPROFIT", "P_TOTALNAV", "TOTAL_ASSETS"
                ws.Range(ws.Cells(2, c), ws.Cells(rCount + 1, c)).NumberFormat = "#,##0.00"
        End Select
    Next c

    ' 批量写入数据
    With ws.Range(ws.Cells(1, 1), ws.Cells(rCount + 1, cCount))
        .Value2 = arr
        .rows(1).Font.Bold = True
        .rows(1).Interior.Color = RGB(31, 73, 125) ' 深蓝表头
        .rows(1).Font.Color = RGB(255, 255, 255)  ' 白色文字
        .Borders.LineStyle = 1                     ' xlContinuous
        .Borders.Color = RGB(216, 216, 216)
        .columns.AutoFit
    End With
End Sub

Private Function ExtractScalar(ByVal result As clsAnalysisResult) As Variant
    Dim primaryTblName As String, tbl As clsDataTable
    If result.status = "FAILED" Then
        ExtractScalar = CVErr(2015) ' xlErrValue
        Exit Function
    End If
    primaryTblName = FindPrimaryTableName(result)
    Set tbl = result.table(primaryTblName)
    If tbl.rowCount >= 1 And tbl.fieldCount >= 1 Then
        ' 优先返回第1行第2列或第1列数值
        If tbl.fieldCount >= 2 Then
            ExtractScalar = tbl.ValueAtColumn(1, 2)
        Else
            ExtractScalar = tbl.ValueAtColumn(1, 1)
        End If
    Else
        ExtractScalar = CVErr(2042) ' xlErrNA
    End If
End Function

Private Function FindPrimaryTableName(ByVal result As clsAnalysisResult) As String
    Dim names As Variant, i As Long, nameStr As String
    names = result.TableNames
    For i = LBound(names) To UBound(names)
        nameStr = CStr(names(i))
        If nameStr <> "DIAGNOSTICS" And nameStr <> "LINEAGE" And nameStr <> "PARAMETERS" Then
            FindPrimaryTableName = nameStr
            Exit Function
        End If
    Next i
    FindPrimaryTableName = "PARAMETERS"
End Function

Private Function GetOrCreateWorksheet(ByVal wb As Object, ByVal sheetName As String) As Object
    Dim ws As Object
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        ws.name = sheetName
    End If
    Set GetOrCreateWorksheet = ws
End Function

Private Function CleanSheetName(ByVal rawName As String) As String
    Dim s As String, badChars As Variant, i As Long
    s = Trim$(rawName)
    badChars = Array(":", "\", "/", "?", "*", "[", "]")
    For i = LBound(badChars) To UBound(badChars)
        s = Replace(s, badChars(i), "_")
    Next i
    If Len(s) > 31 Then s = Left$(s, 31)
    CleanSheetName = s
End Function
