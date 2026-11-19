VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEquityService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const SQL_DAILY_QUOTE As String = "STOCK_Quotes.sql"
Private Const DEFAULT_LOOKBACK_DAYS As Long = 400

Public Function GetQuotePanel(ByVal secCodes As Variant, _
        ByVal startDate As Date, _
        ByVal endDate As Date, _
        Optional ByVal lookbackDays As Long = DEFAULT_LOOKBACK_DAYS) As clsEquityPanel
    Dim config As clsConfig, sourceAlias As String, conn As Object
    Dim masterSvc As clsSecurityMasterService, fullMaster As clsDataTable, subMaster As clsDataTable
    Dim innerCodes() As Long, displayCodes As String, loadedStart As Date
    Dim builder As clsSqlBuilder, executor As clsDbExecutor, schema As clsDataSchema
    Dim quoteDataset As clsDataSet, quoteFrame As clsTimeSeriesFrame
    Dim sources As New Collection, panel As clsEquityPanel
    Dim errNum As Long, errDesc As String

    If dateValue(startDate) > dateValue(endDate) Then Err.Raise 5, "clsEquityService.GetQuotePanel", "Start date is after end date."
    If lookbackDays < 0 Then Err.Raise 5, "clsEquityService.GetQuotePanel", "Lookback days cannot be negative."

    On Error GoTo Failed
    Set masterSvc = New clsSecurityMasterService
    Set fullMaster = masterSvc.LoadMasterTable()
    Set subMaster = ExtractEquitySubMaster(fullMaster, secCodes, innerCodes, displayCodes)

    loadedStart = DateAdd("d", -lookbackDays, dateValue(startDate))

    Set config = New clsConfig
    sourceAlias = config.GetSourceAliasByModule("FINANCIAL")
    If Len(Trim$(sourceAlias)) = 0 Then Err.Raise 5, "clsEquityService.GetQuotePanel", "Missing FINANCIAL source alias."
    Set conn = ConnMgr().GetConnection(sourceAlias)

    Set builder = New clsSqlBuilder
    Call builder.LoadTemplate(SQL_DAILY_QUOTE)
    Call builder.SetChunkedInPredicate("InnerCodePred_AF", "InnerCode", innerCodes, True)
    Call builder.SetChunkedInPredicate("InnerCodePred_Q", "q.InnerCode", innerCodes, True)
    Call builder.SetDate("StartDate", loadedStart)
    Call builder.SetDate("EndDate", dateValue(endDate))

    Set schema = NewStockQuoteSchema()
    Set executor = New clsDbExecutor
    Set quoteDataset = executor.ReadDataSet(conn, builder.Build(), schema, sourceAlias, SQL_DAILY_QUOTE)
    sources.Add quoteDataset

    Set quoteFrame = New clsTimeSeriesFrame
    quoteFrame.LoadData quoteDataset.data, "T_DATE"

    Set panel = New clsEquityPanel
    panel.Load displayCodes, startDate, endDate, loadedStart, sourceAlias, subMaster, quoteFrame, sources
    Set GetQuotePanel = panel
    Exit Function

Failed:
    errNum = Err.number: errDesc = Err.description
    Err.Raise errNum, "clsEquityService.GetQuotePanel", "Equity quote load failed. codes=" & displayCodes & "; detail=" & errDesc
End Function

Private Function ExtractEquitySubMaster(ByVal fullMaster As clsDataTable, _
        ByVal secCodes As Variant, _
        ByRef outInnerCodes() As Long, _
        ByRef outDisplayCodes As String) As clsDataTable
    Dim codeList As Variant, idxBySec As Object, r As Long
    Dim assetClass As String, secCode As String, mktCode As String, compositeKey As String
    Dim reqItem As Variant, reqText As String, dotPos As Long, reqSec As String, reqMkt As String
    Dim matchedRows As New Collection, seenInner As Object, innerVal As Long
    Dim outTable() As Variant, cCount As Long, c As Long, rowIdx As Long, resTbl As clsDataTable

    codeList = NormalizeCodeArray(secCodes)
    Set idxBySec = CreateObject("Scripting.Dictionary")
    idxBySec.CompareMode = 1
    Set seenInner = CreateObject("Scripting.Dictionary")

    ' 为 EQUITY 构建 SEC_CODE 与 SEC_CODE.MARKET_CODE 双键索引（A股优先于港股）
    For r = 1 To fullMaster.rowCount
        assetClass = UCase$(Trim$(CStr(fullMaster.ValueAt(r, "ASSET_CLASS"))))
        If assetClass = "EQUITY" Then
            secCode = UCase$(Trim$(CStr(fullMaster.ValueAt(r, "SEC_CODE"))))
            mktCode = UCase$(Trim$(CStr(fullMaster.ValueAt(r, "MARKET_CODE"))))
            compositeKey = secCode & "." & mktCode
            If Not idxBySec.exists(compositeKey) Then idxBySec.Add compositeKey, r
            If Not idxBySec.exists(secCode) Then
                idxBySec.Add secCode, r
            ElseIf mktCode = "SSE" Or mktCode = "SZSE" Or mktCode = "BSE" Then
                idxBySec(secCode) = r
            End If
        End If
    Next r

    For Each reqItem In codeList
        reqText = UCase$(Trim$(CStr(reqItem)))
        If Len(reqText) > 0 Then
            dotPos = InStr(reqText, ".")
            If dotPos > 0 Then
                reqSec = Left$(reqText, dotPos - 1)
                reqMkt = Mid$(reqText, dotPos + 1)
                compositeKey = reqSec & "." & reqMkt
                If Not idxBySec.exists(compositeKey) Then
                    Err.Raise 5, "clsEquityService.ExtractEquitySubMaster", "Equity code not found in DIM_SECURITY_MASTER: " & reqText
                End If
                r = CLng(idxBySec(compositeKey))
            Else
                If Not idxBySec.exists(reqText) Then
                    Err.Raise 5, "clsEquityService.ExtractEquitySubMaster", "Equity code not found in DIM_SECURITY_MASTER: " & reqText
                End If
                r = CLng(idxBySec(reqText))
            End If

            innerVal = CLng(fullMaster.ValueAt(r, "INNER_CODE"))
            If Not seenInner.exists(innerVal) Then
                seenInner.Add innerVal, True
                matchedRows.Add r
                If Len(outDisplayCodes) > 0 Then outDisplayCodes = outDisplayCodes & ","
                outDisplayCodes = outDisplayCodes & CStr(fullMaster.ValueAt(r, "SEC_CODE")) & "." & CStr(fullMaster.ValueAt(r, "MARKET_CODE"))
            End If
        End If
    Next reqItem

    If matchedRows.count = 0 Then Err.Raise 5, "clsEquityService.ExtractEquitySubMaster", "At least one valid equity code is required."

    ReDim outInnerCodes(1 To matchedRows.count)
    cCount = fullMaster.fieldCount
    ReDim outTable(1 To matchedRows.count + 1, 1 To cCount)
    For c = 1 To cCount
        outTable(1, c) = fullMaster.fieldName(c)
    Next c
    For r = 1 To matchedRows.count
        rowIdx = CLng(matchedRows(r))
        outInnerCodes(r) = CLng(fullMaster.ValueAt(rowIdx, "INNER_CODE"))
        For c = 1 To cCount
            outTable(r + 1, c) = fullMaster.ValueAtColumn(rowIdx, c)
        Next c
    Next r

    Set resTbl = New clsDataTable
    resTbl.Load outTable
    Set ExtractEquitySubMaster = resTbl
End Function

Private Function NormalizeCodeArray(ByVal secCodes As Variant) As Variant
    Dim arr() As Variant
    If IsArray(secCodes) Then
        NormalizeCodeArray = secCodes
    Else
        ReDim arr(1 To 1)
        arr(1) = secCodes
        NormalizeCodeArray = arr
    End If
End Function
