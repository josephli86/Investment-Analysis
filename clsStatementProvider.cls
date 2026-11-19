VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStatementProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IValueProvider

Private m_Statement As clsFinancialStatement
Private m_KnownAccounts As Object

Public Sub Initialize(ByVal stmt As clsFinancialStatement, _
                      Optional ByVal knownAccounts As Object = Nothing, _
                      Optional ByVal isolateView As Boolean = True)
    If stmt Is Nothing Then Err.Raise 5, "clsStatementProvider", "Financial statement is required."
    If isolateView Then
        Set m_Statement = stmt.CloneView()
    Else
        Set m_Statement = stmt
    End If
    If knownAccounts Is Nothing Then
        Set m_KnownAccounts = ExtractAccountsFromStatement(m_Statement)
    Else
        Set m_KnownAccounts = knownAccounts
    End If
    ' Optional manually-entered contingent account; absence means zero.
    m_KnownAccounts("PROVISION") = True
End Sub

Public Property Get statement() As clsFinancialStatement
    Set statement = m_Statement
End Property

Public Sub PinReportVersion(ByVal rptDate As String, Optional ByVal ifAdj As Long = 2, Optional ByVal publishedDate As String = "")
    If m_Statement Is Nothing Then Err.Raise 5, "clsStatementProvider", "Provider is not initialized."
    m_Statement.SelectReportVersion rptDate, ifAdj, publishedDate
End Sub

Public Function CloneProvider() As clsStatementProvider
    Dim cloned As clsStatementProvider
    If m_Statement Is Nothing Then Err.Raise 5, "clsStatementProvider", "Provider is not initialized."
    Set cloned = New clsStatementProvider
    cloned.Initialize m_Statement, m_KnownAccounts, True
    Set CloneProvider = cloned
End Function

Public Function ResolveInitialVersionKey(ByVal period As String, ByVal ifAdjusted As Long) As String
    ResolveInitialVersionKey = m_Statement.SelectedVersionKey(period, ifAdjusted)
End Function

Public Function ResolveComparableVersionKey(ByVal priorPeriod As String, ByVal currentVersionKey As String) As String
    ResolveComparableVersionKey = m_Statement.SelectedComparableVersionKey(priorPeriod, currentVersionKey)
End Function

Private Property Get IValueProvider_ProviderName() As String
    IValueProvider_ProviderName = "STATEMENT:" & m_Statement.companyCode
End Property

Private Function IValueProvider_TryResolve(ByVal symbol As String, _
                                           ByVal evalDate As String, _
                                           ByVal contextState As Variant, _
                                           ByRef outValue As Variant, _
                                           ByRef outStatus As String, _
                                           ByRef outSourceKey As String) As Boolean
    Dim versionKey As String, amount As Variant
    symbol = UCase$(Trim$(symbol))
    versionKey = CStr(contextState(3))
    outSourceKey = versionKey

    If symbol = "TAX_RATE" Then
        If Len(versionKey) = 0 Then Exit Function
        If m_Statement.TryGetVersionRawValue(versionKey, "TAX_RATE", amount) Then
            outValue = CDbl(amount)
            outStatus = "OK"
        Else
            outValue = 0.25
            outStatus = "DEFAULT_TAX_RATE"
        End If
        IValueProvider_TryResolve = True
        Exit Function
    End If

    If Not m_KnownAccounts.exists(symbol) Then
        IValueProvider_TryResolve = False
        Exit Function
    End If

    If Len(versionKey) = 0 Then
        Err.Raise vbObjectError + 2102, "clsStatementProvider", "Missing report version for " & evalDate
    End If

    If m_Statement.TryGetVersionRawValue(versionKey, symbol, amount) Then
        outValue = CDbl(amount)
        outStatus = "OK"
    Else
        outValue = 0#
        outStatus = "DEFAULT_ZERO"
    End If
    IValueProvider_TryResolve = True
End Function

Private Function ExtractAccountsFromStatement(ByVal stmt As clsFinancialStatement) As Object
    Dim dict As Object, longTbl As Variant, r As Long, acc As String
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1
    longTbl = stmt.ExportLongTable()
    If Not IsEmpty(longTbl) Then
        For r = 1 To UBound(longTbl, 1)
            acc = UCase$(Trim$(CStr(longTbl(r, 6))))
            If Len(acc) > 0 Then dict(acc) = True
        Next r
    End If
    Set ExtractAccountsFromStatement = dict
End Function
