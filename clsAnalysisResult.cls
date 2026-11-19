VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAnalysisResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_RunId As String, m_Request As clsAnalysisRequest
Private m_Tables As Object, m_Diagnostics As Collection, m_Lineage As Collection, m_Frozen As Boolean
Private m_Status As String

Public Sub Initialize(ByVal request As clsAnalysisRequest)
    Static sequence As Long
    If Not m_Tables Is Nothing Then Err.Raise 5, , "Result already initialized."
    If request Is Nothing Then Err.Raise 5, , "Request required."
    request.Freeze: Set m_Request = request
    sequence = sequence + 1
    m_RunId = Format$(Now, "yyyymmdd-hhnnss") & "-" & CStr(sequence) & "-" & CStr(CLng(Timer * 100#))
    m_Status = "SUCCESS"
    Set m_Tables = CreateObject("Scripting.Dictionary"): m_Tables.CompareMode = 1
    Set m_Diagnostics = New Collection: Set m_Lineage = New Collection
End Sub

Public Sub AddTable(ByVal name As String, ByVal data As clsDataTable)
    RequireWritable
    name = UCase$(Trim$(name))
    If Len(name) = 0 Then Err.Raise 5, , "Table name required."
    If name = "DIAGNOSTICS" Or name = "LINEAGE" Or name = "PARAMETERS" Then Err.Raise 5, , "Reserved result table name: " & name
    If data Is Nothing Then Err.Raise 5, , "Result table required."
    If Not data.IsLoaded Then Err.Raise 5, , "Result table not loaded."
    m_Tables.Add name, data
End Sub

Public Function table(ByVal name As String) As clsDataTable
    If m_Tables Is Nothing Then Err.Raise 5, , "Result not initialized."
    name = UCase$(Trim$(name))
    Select Case name
        Case "DIAGNOSTICS": Set table = DiagnosticsTable()
        Case "LINEAGE": Set table = LineageTable()
        Case "PARAMETERS": Set table = m_Request.ParametersTable()
        Case Else
            If Not m_Tables.exists(name) Then Err.Raise 5, , "Missing result table: " & name
            Set table = m_Tables(name)
    End Select
End Function

Public Function TableNames() As Variant
    Dim names() As Variant, key As Variant, i As Long
    ReDim names(0 To m_Tables.count + 2)
    For Each key In m_Tables.Keys: names(i) = key: i = i + 1: Next key
    names(i) = "DIAGNOSTICS": names(i + 1) = "LINEAGE": names(i + 2) = "PARAMETERS"
    TableNames = names
End Function

Public Sub AddDiagnostic(ByVal code As String, ByVal entity As String, ByVal stage As String, ByVal Message As String, Optional ByVal severity As String = "WARNING")
    RequireWritable
    severity = UCase$(Trim$(severity))
    If severity <> "INFO" And severity <> "WARNING" And severity <> "ERROR" Then Err.Raise 5, , "Invalid diagnostic severity."
    m_Diagnostics.Add Array(code, entity, stage, Message, severity)
    If severity = "ERROR" Then m_Status = "FAILED"
    If severity = "WARNING" And m_Status = "SUCCESS" Then m_Status = "PARTIAL"
End Sub

Public Sub AddSource(ByVal dataset As clsDataSet)
    RequireWritable
    If dataset Is Nothing Then Err.Raise 5, , "Source dataset required."
    m_Lineage.Add Array(dataset.schema.datasetName, dataset.schema.SchemaVersion, dataset.sourceAlias, dataset.queryId, dataset.revisionId, dataset.LoadedAt, dataset.AvailabilityEvidence)
    If dataset.revisionId = "" Then AddDiagnostic "UNVERSIONED_SOURCE", dataset.schema.datasetName, "LOAD", "Source revision unavailable; replay requires a captured input snapshot."
    If dataset.AvailabilityEvidence = "UNKNOWN" Then AddDiagnostic "UNKNOWN_AVAILABILITY", dataset.schema.datasetName, "LOAD", "Historical point-in-time availability is not verified."
End Sub

Public Function DiagnosticsTable() As clsDataTable
    Set DiagnosticsTable = RowsTable(Array("CODE", "ENTITY", "STAGE", "MESSAGE", "SEVERITY"), m_Diagnostics)
End Function

Public Function LineageTable() As clsDataTable
    Set LineageTable = RowsTable(Array("DATASET", "SCHEMA_VERSION", "SOURCE", "QUERY", "REVISION", "LOADED_AT", "AVAILABILITY"), m_Lineage)
End Function

Private Function RowsTable(ByVal fields As Variant, ByVal rows As Collection) As clsDataTable
    Dim t() As Variant, r As Long, c As Long, item As Variant, data As clsDataTable
    ReDim t(1 To rows.count + 1, 1 To UBound(fields) + 1)
    For c = 0 To UBound(fields): t(1, c + 1) = fields(c): Next c
    For r = 1 To rows.count
        item = rows(r)
        For c = 0 To UBound(fields): t(r + 1, c + 1) = item(c): Next c
    Next r
    Set data = New clsDataTable: data.Load t: Set RowsTable = data
End Function

Public Sub Freeze()
    If m_Tables Is Nothing Then Err.Raise 5, , "Result not initialized."
    m_Frozen = True
End Sub

Public Property Get Diagnostics() As Collection
    Set Diagnostics = CopyRows(m_Diagnostics)
End Property

Public Property Get Lineage() As Collection
    Set Lineage = CopyRows(m_Lineage)
End Property

Private Function CopyRows(ByVal rows As Collection) As Collection
    Dim result As New Collection, item As Variant
    For Each item In rows: result.Add item: Next item
    Set CopyRows = result
End Function

Public Property Get RunId() As String
    RunId = m_RunId
End Property

Public Property Get status() As String
    status = m_Status
End Property

Public Property Get request() As clsAnalysisRequest
    Set request = m_Request
End Property

Public Property Get ResultSchemaVersion() As String
    ResultSchemaVersion = "1"
End Property

Private Sub RequireWritable()
    If m_Tables Is Nothing Then Err.Raise 5, , "Result not initialized."
    If m_Frozen Then Err.Raise 5, , "Result is frozen."
End Sub
