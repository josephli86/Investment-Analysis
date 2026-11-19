VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSecurityMasterService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const SQL_MASTER_TEMPLATE As String = "DIM_SecurityMaster.sql"
Private Const MASTER_TABLE_NAME As String = "DIM_SECURITY_MASTER"

Private m_Config As clsConfig

Private Sub Class_Initialize()
    Set m_Config = New clsConfig
End Sub

' 从聚源 Oracle 同步至本地 Access (SEC_MARKET.accdb)
' forceFullRefresh = False 时自动按本地 MAX(UPDATE_TIME) 做增量更新；首次空表自动执行全量初始化
Public Function SyncFromOracle(Optional ByVal forceFullRefresh As Boolean = False) As clsAnalysisResult
    Dim oracleAlias As String, accessAlias As String, accessConnStr As String
    Dim oracleConn As Object, accessConn As Object
    Dim builder As clsSqlBuilder, executor As clsDbExecutor
    Dim schema As clsDataSchema, dataset As clsDataSet, table As clsDataTable
    Dim lastSyncText As String, isFull As Boolean, inTrans As Boolean
    Dim context As clsAnalysisContext, request As clsAnalysisRequest, result As clsAnalysisResult
    Dim summary(1 To 2, 1 To 5) As Variant, summaryTable As clsDataTable
    Dim errorNumber As Long, errorDescription As String

    On Error GoTo Failed
    oracleAlias = m_Config.GetSourceAliasByModule("FINANCIAL")
    accessAlias = m_Config.GetSourceAliasByModule("SEC_MARKET")
    If Len(Trim$(accessAlias)) = 0 Then Err.Raise 5, "clsSecurityMasterService", "Missing SEC_MARKET_SOURCE in db_config.ini."
    If m_Config.GetUseMock(oracleAlias) Then
        Err.Raise 5, "clsSecurityMasterService", "Oracle source [" & oracleAlias & "] is in mock mode (USE_MOCK=1)."
    End If

    accessConnStr = m_Config.GetAccessConnString(accessAlias)
    Set accessConn = CreateObject("ADODB.Connection")
    accessConn.ConnectionTimeout = 15
    accessConn.CommandTimeout = 120
    accessConn.Open accessConnStr

    EnsureMasterTableExists accessConn

    ' 1. 确定同步起点（全量 vs 增量）
    If forceFullRefresh Then
        lastSyncText = "1900-01-01 00:00:00"
        isFull = True
    Else
        lastSyncText = ResolveIncrementalCutoff(accessConn, isFull)
    End If

    ' 2. 从聚源 Oracle 拉取并经 Schema 归一化
    Set builder = New clsSqlBuilder
    Call builder.LoadTemplate(SQL_MASTER_TEMPLATE)
    ' 若你的 SQL 模板中写的是 TO_DATE({LastSyncTime}, ...)，第三个参数保持 True（自动加单引号）；
    ' 若模板中已自带单引号 TO_DATE('{LastSyncTime}', ...)，请将第三个参数改为 False。
    Call builder.SetString("LastSyncTime", lastSyncText, True)

    Set oracleConn = ConnMgr().GetConnection(oracleAlias)
    Set schema = NewSecurityMasterSchema()
    Set executor = New clsDbExecutor
    Set dataset = executor.ReadDataSet(oracleConn, builder.Build(), schema, oracleAlias, SQL_MASTER_TEMPLATE)
    Set table = dataset.data

    ' 3. 事务内原子写入本地 Access
    accessConn.BeginTrans
    inTrans = True
    If isFull Then
        accessConn.Execute "DELETE FROM [" & MASTER_TABLE_NAME & "]"
    ElseIf table.rowCount > 0 Then
        DeleteExistingByInnerCode accessConn, table
    End If
    If table.rowCount > 0 Then
        BulkInsertToAccess accessConn, table
    End If
    accessConn.CommitTrans
    inTrans = False

    accessConn.Close
    Set accessConn = Nothing

    ' 失效内存缓存，确保后续查询读到最新主数据
    CacheMgr().InvalidatePrefix "SEC_MASTER|"

    ' 4. 封装标准分析结果与血缘
    Set context = New clsAnalysisContext
    context.Initialize Date, Now, "CNY", "SECURITY_MASTER_SYNC_V1"
    Set request = New clsAnalysisRequest
    request.Initialize "SECURITY_MASTER_SYNC", context
    request.SetParameter "TARGET_SOURCE", accessAlias
    request.SetParameter "SYNC_MODE", IIf(isFull, "FULL", "INCREMENTAL")
    request.SetParameter "CUTOFF_TIME", lastSyncText

    summary(1, 1) = "TARGET_SOURCE": summary(1, 2) = "SYNC_MODE": summary(1, 3) = "CUTOFF_TIME"
    summary(1, 4) = "SYNCED_ROWS": summary(1, 5) = "STATUS"
    summary(2, 1) = accessAlias: summary(2, 2) = IIf(isFull, "FULL", "INCREMENTAL"): summary(2, 3) = lastSyncText
    summary(2, 4) = table.rowCount: summary(2, 5) = "OK"

    Set summaryTable = New clsDataTable
    summaryTable.Load summary
    Set result = New clsAnalysisResult
    result.Initialize request
    result.AddTable "SYNC_SUMMARY", summaryTable
    result.AddSource dataset
    result.Freeze
    Set SyncFromOracle = result
    Exit Function

Failed:
    errorNumber = Err.number
    errorDescription = Err.description
    On Error Resume Next
    If inTrans And Not accessConn Is Nothing Then accessConn.RollbackTrans
    If Not accessConn Is Nothing Then
        If accessConn.State <> 0 Then accessConn.Close
    End If
    On Error GoTo 0
    Err.Raise errorNumber, "clsSecurityMasterService.SyncFromOracle", errorDescription
End Function

' 读取全量主数据表（带内存缓存，供行情/财报/组合分析器做 O(1) 映射路由）
Public Function LoadMasterTable(Optional ByVal forceReload As Boolean = False) As clsDataTable
    Dim accessAlias As String, cacheKey As String, cached As Variant
    Dim accessConn As Object, executor As clsDbExecutor, schema As clsDataSchema, dataset As clsDataSet
    Dim errorNumber As Long, errorDescription As String
    On Error GoTo Failed
    accessAlias = m_Config.GetSourceAliasByModule("SEC_MARKET")
    cacheKey = "SEC_MASTER|" & accessAlias
    If Not forceReload Then
        If CacheMgr().TryGet(cacheKey, cached) Then
            Set LoadMasterTable = cached
            Exit Function
        End If
    End If

    Set accessConn = CreateObject("ADODB.Connection")
    accessConn.mode = 1 ' adModeRead
    accessConn.Open m_Config.GetAccessConnString(accessAlias)

    Set schema = NewSecurityMasterSchema()
    Set executor = New clsDbExecutor
    Set dataset = executor.ReadDataSet(accessConn, "SELECT * FROM [" & MASTER_TABLE_NAME & "]", schema, accessAlias, MASTER_TABLE_NAME)
    accessConn.Close

    CacheMgr().SetItem cacheKey, dataset.data, 3600
    Set LoadMasterTable = dataset.data
    Exit Function

Failed:
    errorNumber = Err.number
    errorDescription = Err.description
    On Error Resume Next
    If Not accessConn Is Nothing Then If accessConn.State <> 0 Then accessConn.Close
    On Error GoTo 0
    Err.Raise errorNumber, "clsSecurityMasterService.LoadMasterTable", errorDescription
End Function

Private Sub EnsureMasterTableExists(ByVal accessConn As Object)
    Dim rsSchema As Object, exists As Boolean
    Set rsSchema = accessConn.OpenSchema(20, Array(Empty, Empty, MASTER_TABLE_NAME, "TABLE")) ' adSchemaTables
    exists = Not rsSchema.EOF
    rsSchema.Close
    If exists Then Exit Sub

    accessConn.Execute "CREATE TABLE [" & MASTER_TABLE_NAME & "] (" & _
        "[INNER_CODE] LONG CONSTRAINT PK_DIM_SEC_MASTER PRIMARY KEY, " & _
        "[SEC_CODE] TEXT(32) NOT NULL, " & _
        "[MARKET_CODE] TEXT(16) NOT NULL, " & _
        "[SEC_NAME] TEXT(100), " & _
        "[ASSET_CLASS] TEXT(16) NOT NULL, " & _
        "[SUB_TYPE] TEXT(24) NOT NULL, " & _
        "[COMPANY_CODE] TEXT(32), " & _
        "[COMPANY_NAME] TEXT(200), " & _
        "[SOURCE_TABLE] TEXT(32) NOT NULL, " & _
        "[LIST_DATE] DATETIME, " & _
        "[DELIST_DATE] DATETIME, " & _
        "[UPDATE_TIME] DATETIME)"
    accessConn.Execute "CREATE INDEX IDX_DIM_SEC_CODE ON [" & MASTER_TABLE_NAME & "] ([SEC_CODE], [MARKET_CODE])"
    accessConn.Execute "CREATE INDEX IDX_DIM_SEC_COMPANY ON [" & MASTER_TABLE_NAME & "] ([COMPANY_CODE])"
End Sub

Private Function ResolveIncrementalCutoff(ByVal accessConn As Object, ByRef outIsFull As Boolean) As String
    Dim rs As Object, maxTs As Variant
    Set rs = accessConn.Execute("SELECT MAX([UPDATE_TIME]) AS MAX_TS FROM [" & MASTER_TABLE_NAME & "]")
    If Not rs.EOF Then maxTs = rs.fields(0).Value
    rs.Close
    If IsNull(maxTs) Or IsEmpty(maxTs) Or Not IsDate(maxTs) Then
        outIsFull = True
        ResolveIncrementalCutoff = "1900-01-01 00:00:00"
    Else
        outIsFull = False
        ' 回退 1 天作为安全缓冲窗口，确保跨日边界变更不遗漏
        ResolveIncrementalCutoff = Format$(DateAdd("d", -1, CDate(maxTs)), "yyyy-mm-dd hh:nn:ss")
    End If
End Function

Private Sub DeleteExistingByInnerCode(ByVal accessConn As Object, ByVal table As clsDataTable)
    Dim codes() As Variant, r As Long, builder As clsSqlBuilder
    ReDim codes(1 To table.rowCount)
    For r = 1 To table.rowCount
        codes(r) = table.ValueAt(r, "INNER_CODE")
    Next r
    Set builder = New clsSqlBuilder
    Call builder.LoadRawSql("DELETE FROM " & MASTER_TABLE_NAME & " WHERE {CodePred}", "DELETE_SEC_MASTER_DELTA")
    Call builder.SetChunkedInPredicate("CodePred", MASTER_TABLE_NAME & ".INNER_CODE", codes, True)
    accessConn.Execute builder.Build()
End Sub

Private Sub BulkInsertToAccess(ByVal accessConn As Object, ByVal table As clsDataTable)
    Dim rs As Object, r As Long, c As Long, fieldName As String, colCount As Long
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM [" & MASTER_TABLE_NAME & "] WHERE 1 = 0", accessConn, 1, 3 ' adOpenKeyset, adLockOptimistic
    colCount = table.fieldCount
    For r = 1 To table.rowCount
        rs.AddNew
        For c = 1 To colCount
            fieldName = table.fieldName(c)
            rs.fields(fieldName).Value = table.ValueAtColumn(r, c)
        Next c
        rs.Update
    Next r
    rs.Close
End Sub
