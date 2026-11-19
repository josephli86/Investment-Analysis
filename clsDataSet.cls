VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDataSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Data As clsDataTable, m_Schema As clsDataSchema
Private m_Source As String, m_Query As String, m_Revision As String
Private m_LoadedAt As Date, m_Ready As Boolean, m_Availability As String

Public Sub Initialize(ByVal data As clsDataTable, ByVal schema As clsDataSchema, ByVal sourceAlias As String, ByVal queryId As String, Optional ByVal revisionId As String = "", Optional ByVal AvailabilityEvidence As String = "UNKNOWN")
    If m_Ready Then Err.Raise 5, , "Dataset is immutable."
    If schema Is Nothing Then Err.Raise 5, , "Schema required."
    If Len(Trim$(sourceAlias)) = 0 Or Len(Trim$(queryId)) = 0 Then Err.Raise 5, , "Source and query required."
    Set m_Data = schema.Normalize(data): Set m_Schema = schema
    m_Source = sourceAlias: m_Query = queryId: m_Revision = revisionId
    m_Availability = AvailabilityEvidence: m_LoadedAt = Now: m_Ready = True
End Sub
Public Property Get data() As clsDataTable
    If Not m_Ready Then Err.Raise 5, , "Dataset not initialized."
    Set data = m_Data
End Property
Public Property Get schema() As clsDataSchema
    Set schema = m_Schema
End Property
Public Property Get sourceAlias() As String
    sourceAlias = m_Source
End Property
Public Property Get queryId() As String
    queryId = m_Query
End Property
Public Property Get revisionId() As String
    revisionId = m_Revision
End Property
Public Property Get AvailabilityEvidence() As String
    AvailabilityEvidence = m_Availability
End Property
Public Property Get LoadedAt() As Date
    LoadedAt = m_LoadedAt
End Property
