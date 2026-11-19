VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDbDataSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IDataSetSource
Private m_Connection As Object, m_Schema As clsDataSchema
Private m_Sql As String, m_Source As String, m_Query As String
Private m_Revision As String, m_Availability As String

' Configure one approved dataset query; this is not an arbitrary-query UI.
Public Sub Initialize(ByVal connection As Object, ByVal sqlText As String, ByVal schema As clsDataSchema, ByVal sourceAlias As String, ByVal queryId As String, Optional ByVal revisionId As String = "", Optional ByVal availability As String = "UNKNOWN")
    If Not m_Connection Is Nothing Then Err.Raise 5, , "Source already initialized."
    If connection Is Nothing Or schema Is Nothing Then Err.Raise 5, , "Connection and schema required."
    If Len(Trim$(sourceAlias)) = 0 Or Len(Trim$(queryId)) = 0 Then Err.Raise 5, , "Source and query required."
    Set m_Connection = connection: Set m_Schema = schema
    m_Sql = sqlText: m_Source = sourceAlias: m_Query = queryId
    m_Revision = revisionId: m_Availability = availability
End Sub
Private Function IDataSetSource_LoadData() As clsDataSet
    Dim executor As clsDbExecutor
    Set executor = New clsDbExecutor
    Set IDataSetSource_LoadData = executor.ReadDataSet(m_Connection, m_Sql, m_Schema, m_Source, m_Query, m_Revision, m_Availability)
End Function
