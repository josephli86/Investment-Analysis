VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAnalysisContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Date As Date, m_AsOf As Date, m_Currency As String
Private m_Policy As String, m_Ready As Boolean

Public Sub Initialize(ByVal valuationDate As Date, ByVal KnowledgeAsOf As Date, ByVal BaseCurrency As String, ByVal PolicyVersion As String)
    If m_Ready Then Err.Raise 5, , "Context is immutable."
    If Len(Trim$(BaseCurrency)) = 0 Or Len(Trim$(PolicyVersion)) = 0 Then Err.Raise 5, , "Currency and policy required."
    m_Date = valuationDate: m_AsOf = KnowledgeAsOf
    m_Currency = UCase$(BaseCurrency): m_Policy = PolicyVersion: m_Ready = True
End Sub
Public Property Get valuationDate() As Date
    RequireReady
    valuationDate = m_Date
End Property
Public Property Get KnowledgeAsOf() As Date
    RequireReady
    KnowledgeAsOf = m_AsOf
End Property
Public Property Get BaseCurrency() As String
    BaseCurrency = m_Currency
End Property
Public Property Get PolicyVersion() As String
    PolicyVersion = m_Policy
End Property
Private Sub RequireReady()
    If Not m_Ready Then Err.Raise 5, , "Context not initialized."
End Sub
