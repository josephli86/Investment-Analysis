VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFinancialService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Config As clsConfig
Private m_SourceAlias As String

Private Sub Class_Initialize()
    Set m_Config = New clsConfig
    m_SourceAlias = UCase$(Trim$(m_Config.GetSourceAliasByModule("FINANCIAL")))
End Sub

Public Property Get sourceAlias() As String
    sourceAlias = m_SourceAlias
End Property

Public Property Let sourceAlias(ByVal Value As String)
    m_SourceAlias = UCase$(Trim$(Value))
    If m_SourceAlias = "" Then m_SourceAlias = UCase$(m_Config.GetSourceAliasByModule("FINANCIAL"))
End Property

Public Function GetStatement(ByVal companyCode As String, _
        ByVal beginDate As String, _
        ByVal endDate As String, _
        Optional ByVal ifMerged As Long = 1, _
        Optional ByVal asOfDate As String = "", _
        Optional ByVal publishedDate As String = "", _
        Optional ByVal ifAdjusted As Long = 2) As clsFinancialStatement
    Dim masterStmt As clsFinancialStatement, viewStmt As clsFinancialStatement
    Dim cached As Variant, key As String
    Dim sqlText As String, conn As Object, adapter As clsWideFactAdapter
    Dim executor As clsDbExecutor, schema As clsDataSchema, data As clsDataSet
    Dim errorNumber As Long, errorDescription As String
    Dim builder As clsSqlBuilder
    On Error GoTo Failed
    companyCode = NormalizeCompanyCode(companyCode)
    beginDate = FinancialPeriodEnd(beginDate)
    endDate = FinancialPeriodEnd(endDate)
    If ifAdjusted <> 1 And ifAdjusted <> 2 Then Err.Raise 5, , "IfAdjusted must be 1 or 2."
    If beginDate > endDate Then
        Err.Raise 5, , "Report begin date is after end date."
    End If
    If ifMerged <> 1 And ifMerged <> 2 Then
        Err.Raise 5, , "IfMerged must be 1 or 2."
    End If
    If asOfDate = "" Then asOfDate = Format$(Date, "yyyy-mm-dd")
    asOfDate = Format$(CDate(asOfDate), "yyyy-mm-dd")
    If publishedDate <> "" Then publishedDate = Format$(CDate(publishedDate), "yyyy-mm-dd")
    If publishedDate > asOfDate Then Err.Raise 5, , "Publication date exceeds the as-of date."
    If m_Config.GetUseMock(m_SourceAlias) Then
        Err.Raise 5, , "Production loading requires USE_MOCK=0. Run offline regression fixtures separately."
    End If
    ' Cache key represents the immutable underlying fact store; view filters are applied on cloned views.
    key = "FIN_V6|" & m_Config.LocalSourceAlias & "|" & m_Config.LocalDbPath & "|" & m_SourceAlias & "|" & companyCode & "|" & beginDate & "|" & endDate & "|" & CStr(ifMerged) & "|" & asOfDate
    If CacheMgr.TryGet(key, cached) Then
        Set masterStmt = cached
        Set viewStmt = masterStmt.CloneView()
        viewStmt.ConfigureView ifMerged, asOfDate, publishedDate, endDate, ifAdjusted
        Set GetStatement = viewStmt
        Exit Function
    End If
    Set masterStmt = New clsFinancialStatement
    masterStmt.companyCode = companyCode
    ' Unicode file name is assembled so VBA source remains ASCII.
    Set builder = New clsSqlBuilder
    Call builder.LoadTemplate(ChrW(&H8D22) & ChrW(&H62A5) & ".sql")
    Call builder.SetNumber("CompanyCode", companyCode)
    ' The financial template already quotes its date placeholders.
    Call builder.SetDate("ReportBegin", beginDate, False)
    Call builder.SetDate("ReportEnd", endDate, False)
    Call builder.SetDate("AsOfDate", asOfDate, False)
    Call builder.SetNumber("IfMerged", ifMerged)
    sqlText = builder.Build()
    Set conn = ConnMgr.GetConnection(m_SourceAlias)
    Set executor = New clsDbExecutor
    Set schema = NewFinancialWideSchema()
    Set data = executor.ReadDataSet(conn, sqlText, schema, m_SourceAlias, ChrW(&H8D22) & ChrW(&H62A5) & ".sql")
    Set adapter = New clsWideFactAdapter
    adapter.LoadData data.data, masterStmt
    masterStmt.AddSource data
    LoadLocalOverrides masterStmt, beginDate, endDate, ifMerged, asOfDate
    CacheMgr.SetItem key, masterStmt
    Set viewStmt = masterStmt.CloneView()
    viewStmt.ConfigureView ifMerged, asOfDate, publishedDate, endDate, ifAdjusted
    Set GetStatement = viewStmt
    Exit Function
Failed:
    errorNumber = Err.number
    errorDescription = Err.description
    Err.Raise errorNumber, "clsFinancialService.GetStatement", errorDescription
End Function

' Public business query: CompanyCode + fiscal period + merged + adjusted.
' History uses the same as-of cutoff, independently resolving each period's batch.
Public Function GetReport(ByVal companyCode As String, ByVal fiscalPeriod As String, Optional ByVal ifMerged As Long = 1, Optional ByVal ifAdjusted As Long = 2, Optional ByVal asOfDate As String = "", Optional ByVal historyYears As Long = 1) As clsFinancialStatement
    Dim periodEnd As String, beginDate As String
    If historyYears < 0 Or historyYears > 100 Then Err.Raise 5, , "HistoryYears must be between 0 and 100."
    periodEnd = FinancialPeriodEnd(fiscalPeriod)
    beginDate = Format$(DateAdd("yyyy", -historyYears, CDate(periodEnd)), "yyyy-mm-dd")
    Set GetReport = GetStatement(companyCode, beginDate, periodEnd, ifMerged, asOfDate, "", ifAdjusted)
End Function

Public Function NormalizeCompanyCode(ByVal Value As String) As String
    Dim pos As Long, ch As String
    Value = Trim$(Value)
    If Len(Value) = 0 Or Len(Value) > 18 Then Err.Raise 5, , "A numeric company code is required."
    For pos = 1 To Len(Value)
        ch = Mid$(Value, pos, 1)
        If ch < "0" Or ch > "9" Then Err.Raise 5, , "Use CompanyCode, not a security ticker."
    Next pos
    Do While Len(Value) > 1 And Left$(Value, 1) = "0"
        Value = Mid$(Value, 2)
    Loop
    If Value = "0" Then Err.Raise 5, , "CompanyCode must be positive."
    NormalizeCompanyCode = Value
End Function

Private Sub LoadLocalOverrides(ByVal stmt As clsFinancialStatement, _
        ByVal beginDate As String, _
        ByVal endDate As String, _
        ByVal ifMerged As Long, _
        ByVal asOfDate As String)
    Dim conn As Object, sqlText As String, connString As String
    Dim key As String, versionKey As String, account As String, seen As Object
    Dim errorNumber As Long, errorDescription As String, localPath As String
    Dim builder As clsSqlBuilder, executor As clsDbExecutor
    Dim schema As clsDataSchema, dataset As clsDataSet, table As clsDataTable
    Dim r As Long, rptDateText As String, publDateText As String, adjFlag As Long
    On Error GoTo Failed
    localPath = m_Config.LocalDbPath
    connString = m_Config.GetLocalConnString()
    If connString = "" Then Exit Sub
    Set conn = CreateObject("ADODB.Connection")
    conn.mode = 1
    conn.ConnectionTimeout = 15
    conn.CommandTimeout = 120
    conn.Open connString
    Set builder = New clsSqlBuilder
    Call builder.LoadRawSql("SELECT [COMPANY_CODE], [END_DATE], [IF_MERGED], [IF_ADJUSTED], [INFO_PUBL_DATE], [ACCOUNT_CODE], [AMOUNT] " & _
        "FROM [MANUAL_FINANCIALS] WHERE [COMPANY_CODE]={CompanyCode}" & _
        " AND [END_DATE]>={BeginDate} AND [END_DATE]<{EndDate}" & _
        " AND [IF_MERGED]={IfMerged} AND [INFO_PUBL_DATE]<{AsOfDate}", "MANUAL_FINANCIALS")
    Call builder.SetNumber("CompanyCode", stmt.companyCode)
    Call builder.SetAccessDate("BeginDate", beginDate)
    Call builder.SetAccessDate("EndDate", DateAdd("d", 1, CDate(endDate)))
    Call builder.SetNumber("IfMerged", ifMerged)
    Call builder.SetAccessDate("AsOfDate", DateAdd("d", 1, CDate(asOfDate)))
    sqlText = builder.Build()
    Set schema = NewFinancialOverrideSchema()
    Set executor = New clsDbExecutor
    Set dataset = executor.ReadDataSet(conn, sqlText, schema, m_Config.LocalSourceAlias, "MANUAL_FINANCIALS")
    Set table = dataset.data
    Set seen = CreateObject("Scripting.Dictionary")
    For r = 1 To table.rowCount
        account = UCase$(Trim$(CStr(table.ValueAt(r, "ACCOUNT_CODE"))))
        If Len(account) = 0 Then Err.Raise 5, , "Local override ACCOUNT_CODE must not be empty."
        adjFlag = CLng(table.ValueAt(r, "IF_ADJUSTED"))
        If adjFlag <> 1 And adjFlag <> 2 Then Err.Raise 5, , "Local override IF_ADJUSTED must be 1 or 2."
        rptDateText = Format$(CDate(table.ValueAt(r, "END_DATE")), "yyyy-mm-dd")
        publDateText = Format$(CDate(table.ValueAt(r, "INFO_PUBL_DATE")), "yyyy-mm-dd")
        versionKey = stmt.companyCode & "|" & rptDateText & "|" & CStr(CLng(table.ValueAt(r, "IF_MERGED"))) & "|" & CStr(adjFlag) & "|" & publDateText
        key = versionKey & "|" & account
        If seen.exists(key) Then Err.Raise 5, , "Duplicate local override: " & key
        seen.Add key, True
        If stmt.FindExactVersion(rptDateText, adjFlag, publDateText) = "" Then
            Err.Raise 5, , "Local override has no matching loaded report: " & versionKey
        End If
        stmt.AddReportFact versionKey, account, table.ValueAt(r, "AMOUNT"), True
    Next r
    stmt.AddSource dataset
    conn.Close
    Exit Sub
Failed:
    errorNumber = Err.number
    errorDescription = Err.description
    On Error Resume Next
    If Not conn Is Nothing Then
        If conn.State <> 0 Then conn.Close
    End If
    On Error GoTo 0
    Err.Raise errorNumber, "clsFinancialService.LoadLocalOverrides", "Local source [" & m_Config.LocalSourceAlias & "] (" & localPath & "), table MANUAL_FINANCIALS: " & errorDescription
End Sub
