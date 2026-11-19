VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDbExecutor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' The connection is borrowed. This executor owns only its recordset.
Public Function ReadDataSet(ByVal connection As Object, ByVal sqlText As String, ByVal schema As clsDataSchema, ByVal sourceAlias As String, ByVal queryId As String, Optional ByVal revisionId As String = "", Optional ByVal availability As String = "UNKNOWN") As clsDataSet
    Dim rs As Object, reader As clsRecordsetReader, data As clsDataTable, result As clsDataSet
    Dim n As Long, description As String, detail As String, item As Object, stage As String
    On Error GoTo Failed
    stage = "VALIDATE"
    If connection Is Nothing Then Err.Raise 5, , "Connection required."
    If connection.State <> 1 Then Err.Raise 5, , "Open connection required."
    If schema Is Nothing Then Err.Raise 5, , "Schema required."
    If Len(Trim$(sqlText)) = 0 Then Err.Raise 5, , "SQL required."
    connection.Errors.Clear
    stage = "EXECUTE"
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sqlText, connection, 0, 1
    stage = "READ"
    Set reader = New clsRecordsetReader: Set data = reader.Read(rs)
    rs.Close: Set rs = Nothing
    stage = "SCHEMA"
    Set result = New clsDataSet
    result.Initialize data, schema, sourceAlias, queryId, revisionId, availability
    Set ReadDataSet = result
    Exit Function
Failed:
    n = Err.number: description = Err.description
    On Error Resume Next
    If Not connection Is Nothing Then
        For Each item In connection.Errors
            detail = detail & " | " & CStr(item.number) & ":" & CStr(item.description)
        Next item
    End If
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
    End If
    On Error GoTo 0
    Err.Raise n, "clsDbExecutor", "source=" & sourceAlias & "; query=" & queryId & "; stage=" & stage & "; " & description & detail
End Function
