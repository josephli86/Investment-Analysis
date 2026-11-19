VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAnalysisRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Task As String, m_Context As clsAnalysisContext, m_Params As Object, m_Frozen As Boolean

Public Sub Initialize(ByVal taskType As String, ByVal context As clsAnalysisContext)
    Dim checkedDate As Date
    If Not m_Params Is Nothing Then Err.Raise 5, , "Request already initialized."
    If context Is Nothing Then Err.Raise 5, , "Context required."
    checkedDate = context.valuationDate
    If Len(Trim$(taskType)) = 0 Then Err.Raise 5, , "Task type required."
    m_Task = UCase$(taskType): Set m_Context = context
    Set m_Params = CreateObject("Scripting.Dictionary"): m_Params.CompareMode = 1
End Sub
Public Sub SetParameter(ByVal key As String, ByVal Value As Variant)
    If m_Params Is Nothing Then Err.Raise 5, , "Request not initialized."
    If m_Frozen Then Err.Raise 5, , "Request is frozen."
    If IsObject(Value) Or IsArray(Value) Or IsError(Value) Then Err.Raise 13, , "Scalar parameter required."
    If Len(Trim$(key)) = 0 Then Err.Raise 5, , "Parameter key required."
    m_Params(UCase$(Trim$(key))) = Value
End Sub
Public Function Parameter(ByVal key As String) As Variant
    If Not m_Params.exists(key) Then Err.Raise 5, , "Missing request parameter: " & key
    Parameter = m_Params(key)
End Function
Public Sub Freeze()
    If m_Params Is Nothing Then Err.Raise 5, , "Request not initialized."
    m_Frozen = True
End Sub
Public Property Get taskType() As String
    taskType = m_Task
End Property
Public Property Get context() As clsAnalysisContext
    Set context = m_Context
End Property
Public Function ParametersTable() As clsDataTable
    Dim t() As Variant, key As Variant, r As Long, data As clsDataTable
    ReDim t(1 To m_Params.count + 1, 1 To 2)
    t(1, 1) = "PARAMETER": t(1, 2) = "VALUE": r = 1
    For Each key In m_Params.Keys
        r = r + 1: t(r, 1) = key: t(r, 2) = m_Params(key)
    Next key
    Set data = New clsDataTable: data.Load t: Set ParametersTable = data
End Function
