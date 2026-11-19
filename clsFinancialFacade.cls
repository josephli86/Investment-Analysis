VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFinancialFacade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_FinService As clsFinancialService
Private m_App As clsFinancialAnalysisApplication
Private m_Renderer As clsResultRenderer

Private Sub Class_Initialize()
    Set m_FinService = New clsFinancialService
    Set m_App = New clsFinancialAnalysisApplication
    Set m_Renderer = New clsResultRenderer
End Sub

Public Function metric(ByVal companyCode As String, ByVal periodEnd As String, ByVal metricName As String, Optional ByVal ifAdjusted As Long = 2, Optional ByVal asOfDate As String = "") As Variant
    metric = GET_FIN_METRIC(companyCode, periodEnd, metricName, 1, CInt(ifAdjusted), asOfDate)
End Function

Public Function Audit(ByVal companyCode As String, ByVal periodEnd As String, ByVal metricName As String, Optional ByVal ifAdjusted As Long = 2) As clsDataTable
    Dim stmt As clsFinancialStatement, ctx As clsAnalysisContext, res As clsAnalysisResult
    Dim pEnd As String, pBegin As String
    pEnd = FinancialPeriodEnd(periodEnd)
    pBegin = Format$(DateAdd("yyyy", -3, CDate(pEnd)), "yyyy-mm-dd")
    Set stmt = m_FinService.GetStatement(companyCode, pBegin, pEnd, 1, "", "", ifAdjusted)
    Set ctx = New clsAnalysisContext
    ctx.Initialize CDate(pEnd), CDate(pEnd), "CNY", "METRIC_AUDIT_V1"
    Set res = m_App.RunAuditMetric(stmt, ctx, pEnd, metricName, ifAdjusted)
    Set Audit = res.table("METRIC_AUDIT")
End Function

Public Function statement(ByVal companyCode As String, Optional ByVal beginDate As String = "", Optional ByVal endDate As String = "", Optional ByVal ifAdjusted As Long = 2, Optional ByVal cutoffDate As String = "") As clsFinancialStatement
    Dim dEnd As String, dBegin As String
    If Len(Trim$(endDate)) = 0 Then dEnd = Format$(Date, "yyyy-mm-dd") Else dEnd = FinancialPeriodEnd(endDate)
    If Len(Trim$(beginDate)) = 0 Then dBegin = Format$(DateAdd("yyyy", -3, CDate(dEnd)), "yyyy-mm-dd") Else dBegin = FinancialPeriodEnd(beginDate)
    Set statement = m_FinService.GetStatement(companyCode, dBegin, dEnd, 1, cutoffDate, "", CInt(ifAdjusted))
End Function

Public Function ValidateBalance(ByVal companyCode As String, ByVal periodEnd As String) As Boolean
    Dim stmt As clsFinancialStatement, pEnd As String, pBegin As String
    Dim totAssets As Variant, totLiab As Variant, totEquity As Variant
    Dim diff As Double
    pEnd = FinancialPeriodEnd(periodEnd)
    pBegin = Format$(DateAdd("yyyy", -1, CDate(pEnd)), "yyyy-mm-dd")
    Set stmt = m_FinService.GetStatement(companyCode, pBegin, pEnd, 1, "", "", 2)
    If Not stmt.TryGetRawValue(pEnd, "TOTAL_ASSETS", totAssets, 2) Then Exit Function
    If Not stmt.TryGetRawValue(pEnd, "TOTAL_LIABILITIES", totLiab, 2) Then Exit Function
    If Not stmt.TryGetRawValue(pEnd, "TOTAL_EQUITY", totEquity, 2) Then Exit Function
    diff = Abs(CDbl(totAssets) - CDbl(totLiab) - CDbl(totEquity))
    ValidateBalance = (diff < 1#)
End Function

Public Function CommonSizeSingle(ByVal companyCode As String, ByVal periodsText As String, Optional ByVal stmtType As String = "IS", Optional ByVal mode As String = "BOTH") As clsDataTable
    Dim rawParts As Variant, i As Long, pClean As String
    Dim minPeriod As String, maxPeriod As String, periods() As String, validCount As Long
    Dim stmt As clsFinancialStatement, ctx As clsAnalysisContext, res As clsAnalysisResult
    rawParts = Split(periodsText, ",")
    ReDim periods(0 To UBound(rawParts))
    minPeriod = "9999-12-31": maxPeriod = "0000-00-00": validCount = 0
    For i = 0 To UBound(rawParts)
        pClean = Trim$(rawParts(i))
        If Len(pClean) > 0 Then
            periods(validCount) = FinancialPeriodEnd(pClean)
            If periods(validCount) > maxPeriod Then maxPeriod = periods(validCount)
            If periods(validCount) < minPeriod Then minPeriod = periods(validCount)
            validCount = validCount + 1
        End If
    Next i
    If validCount = 0 Then Err.Raise 5, "clsFinancialFacade", "No valid periods provided."
    ReDim Preserve periods(0 To validCount - 1)
    Set stmt = m_FinService.GetStatement(companyCode, Format$(DateAdd("yyyy", -1, CDate(minPeriod)), "yyyy-mm-dd"), maxPeriod, 1, "", "", 2)
    Set ctx = New clsAnalysisContext
    ctx.Initialize CDate(maxPeriod), CDate(maxPeriod), "CNY", "COMMON_SIZE_V1"
    Set res = m_App.RunCommonSizeSingleCompany(stmt, ctx, periods, stmtType, mode, 2)
    Set CommonSizeSingle = res.table("MULTI_PERIOD_COMPARISON")
End Function

Public Function CommonSizePeer(ByVal companyCodesText As String, ByVal periodEnd As String, Optional ByVal stmtType As String = "IS") As clsDataTable
    Dim rawCodes As Variant, codes() As String, i As Long, cClean As String, validCount As Long
    Dim pEnd As String, pBegin As String
    Dim stmts As Collection, stmt As clsFinancialStatement
    Dim ctx As clsAnalysisContext, res As clsAnalysisResult
    pEnd = FinancialPeriodEnd(periodEnd)
    pBegin = Format$(DateAdd("yyyy", -1, CDate(pEnd)), "yyyy-mm-dd")
    rawCodes = Split(companyCodesText, ",")
    ReDim codes(0 To UBound(rawCodes))
    validCount = 0
    For i = 0 To UBound(rawCodes)
        cClean = Trim$(rawCodes(i))
        If Len(cClean) > 0 Then
            codes(validCount) = cClean
            validCount = validCount + 1
        End If
    Next i
    If validCount = 0 Then Err.Raise 5, "clsFinancialFacade", "No valid company codes provided."
    ReDim Preserve codes(0 To validCount - 1)
    Set stmts = New Collection
    For i = 0 To UBound(codes)
        Set stmt = m_FinService.GetStatement(codes(i), pBegin, pEnd, 1, "", "", 2)
        stmts.Add stmt
    Next i
    Set ctx = New clsAnalysisContext
    ctx.Initialize CDate(pEnd), CDate(pEnd), "CNY", "PEER_COMPARE_V1"
    Set res = m_App.RunCommonSizePeerComparison(stmts, ctx, pEnd, stmtType, "COMMON_SIZE", 2)
    Set CommonSizePeer = res.table("PEER_COMPARISON")
End Function

Public Function ExportReport(ByVal companyCode As String, ByVal periodsText As String, Optional ByVal stmtType As String = "IS", Optional ByVal deliveryMode As String = "TABLE") As Variant
    Dim req As clsAnalysisRequest, res As clsAnalysisResult, ctx As clsAnalysisContext
    Dim csTable As clsDataTable
    Set csTable = CommonSizeSingle(companyCode, periodsText, stmtType, "BOTH")
    Set ctx = New clsAnalysisContext
    ctx.Initialize Date, Date, "CNY", "FINANCIAL_REPORT_V1"
    Set req = New clsAnalysisRequest
    req.Initialize "FINANCIAL_REPORT", ctx
    req.SetParameter "COMPANY_CODE", companyCode
    req.SetParameter "PERIODS", periodsText
    req.SetParameter "STMT_TYPE", stmtType
    Set res = New clsAnalysisResult
    res.Initialize req
    res.AddTable "COMMON_SIZE", csTable
    Set ExportReport = m_Renderer.Render(res, deliveryMode, "FIN_" & companyCode & "_" & stmtType)
End Function
