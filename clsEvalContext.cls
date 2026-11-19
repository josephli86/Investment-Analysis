VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEvalContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Providers As Collection
Private m_PrimaryStmtProvider As clsStatementProvider
Private m_State As Variant
Private m_Stack As Collection
Private m_Memo As Object

Private Sub Class_Initialize()
    Set m_Providers = New Collection
    Set m_Stack = New Collection
    Set m_Memo = CreateObject("Scripting.Dictionary")
End Sub

Public Sub RegisterProvider(ByVal provider As IValueProvider)
    If provider Is Nothing Then Err.Raise 5, "clsEvalContext", "Provider cannot be Nothing."
    m_Providers.Add provider
    If TypeOf provider Is clsStatementProvider Then
        Set m_PrimaryStmtProvider = provider
    End If
End Sub

Public Sub BeginPeriod(ByVal rptDate As String, Optional ByVal ifAdjusted As Long = 2, Optional ByVal asOfDate As String = "")
    Dim period As String, cutoff As String, versionKey As String, anchorPubl As String, parts As Variant
    period = FinancialPeriodEnd(rptDate)
    If Len(asOfDate) = 0 Then cutoff = Format$(Date, "yyyy-mm-dd") Else cutoff = Format$(CDate(asOfDate), "yyyy-mm-dd")

    If Not m_PrimaryStmtProvider Is Nothing Then
        versionKey = m_PrimaryStmtProvider.ResolveInitialVersionKey(period, ifAdjusted)
        If Len(versionKey) > 0 Then
            parts = Split(versionKey, "|")
            anchorPubl = parts(4)
        End If
    End If

    Set m_Stack = New Collection
    If Not m_Memo Is Nothing Then m_Memo.RemoveAll
    m_State = Array(period, ifAdjusted, anchorPubl, versionKey, cutoff)
End Sub

Public Property Get CurrentPeriod() As String
    If Not IsEmpty(m_State) Then CurrentPeriod = CStr(m_State(0))
End Property

Public Property Get CurrentIfAdjusted() As Long
    If Not IsEmpty(m_State) Then CurrentIfAdjusted = CLng(m_State(1))
End Property

Public Property Get CurrentPublicationDate() As String
    If Not IsEmpty(m_State) Then CurrentPublicationDate = CStr(m_State(2))
End Property

Public Property Get currentVersionKey() As String
    If Not IsEmpty(m_State) Then currentVersionKey = CStr(m_State(3))
End Property

Public Function ContextKey() As String
    If IsEmpty(m_State) Then Err.Raise 5, "clsEvalContext", "Evaluation context not initialized."
    ContextKey = CStr(m_State(0)) & "|" & CStr(m_State(1)) & "|" & CStr(m_State(2)) & "|" & CStr(m_State(3))
End Function

Public Sub Validate(ByVal reportDate As String)
    If IsEmpty(m_State) Then Err.Raise 5, "clsEvalContext", "No active evaluation context."
    If FinancialPeriodEnd(reportDate) <> CStr(m_State(0)) Then
        Err.Raise 5, "clsEvalContext", "Report period does not match evaluation context."
    End If
    If Not m_PrimaryStmtProvider Is Nothing And Len(CStr(m_State(3))) = 0 Then
        Err.Raise vbObjectError + 2102, "clsEvalContext", "Missing report: " & CStr(m_State(0)) & " / adjusted=" & CStr(m_State(1))
    End If
End Sub

Public Sub PushComparable(ByVal priorReportDate As String)
    Dim priorPeriod As String, priorVersionKey As String
    Validate CStr(m_State(0))
    If m_Stack.count >= 100 Then Err.Raise 5, "clsEvalContext", "Context depth exceeds 100."
    priorPeriod = FinancialPeriodEnd(priorReportDate)
    If Not m_PrimaryStmtProvider Is Nothing Then
        priorVersionKey = m_PrimaryStmtProvider.ResolveComparableVersionKey(priorPeriod, CStr(m_State(3)))
    End If
    m_Stack.Add m_State
    m_State = Array(priorPeriod, 1&, CStr(m_State(2)), priorVersionKey, CStr(m_State(4)))
End Sub

Public Sub PopContext()
    If m_Stack.count = 0 Then Err.Raise 5, "clsEvalContext", "Context stack is empty."
    m_State = m_Stack(m_Stack.count)
    m_Stack.Remove m_Stack.count
End Sub

Public Function TryGetMemo(ByVal fullKey As String, ByRef outVal As Double) As Boolean
    If m_Memo.exists(fullKey) Then
        outVal = CDbl(m_Memo(fullKey))
        TryGetMemo = True
    End If
End Function

Public Sub SetMemo(ByVal fullKey As String, ByVal val As Double)
    m_Memo(fullKey) = val
End Sub

Public Function ResolveLeaf(ByVal symbol As String, _
                            ByVal reportDate As String, _
                            ByRef outValue As Double, _
                            ByRef outStatus As String, _
                            ByRef outSourceKey As String) As Boolean
    Dim provider As IValueProvider, rawVal As Variant
    Validate reportDate
    For Each provider In m_Providers
        If provider.TryResolve(symbol, reportDate, m_State, rawVal, outStatus, outSourceKey) Then
            outValue = CDbl(rawVal)
            ResolveLeaf = True
            Exit Function
        End If
    Next provider
End Function
