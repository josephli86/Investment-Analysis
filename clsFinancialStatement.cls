VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFinancialStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_CompanyCode As String
Private m_FactData As Object
Private m_Versions As Object
Private m_AsOfDate As String
Private m_IfMerged As Long
Private m_PublicationPins As Object
Private m_TaxRate As Double
Private m_IsSharedView As Boolean
Private m_Sources As Collection

Private Sub Class_Initialize()
    Set m_FactData = CreateObject("Scripting.Dictionary")
    Set m_Versions = CreateObject("Scripting.Dictionary")
    Set m_PublicationPins = CreateObject("Scripting.Dictionary")
    m_TaxRate = 0.25
    m_IfMerged = 1
    m_AsOfDate = Format$(Date, "yyyy-mm-dd")
    m_IsSharedView = False
    Set m_Sources = New Collection
End Sub

Friend Sub InitSharedView(ByVal companyCode As String, _
                          ByVal versions As Object, _
                          ByVal factData As Object, _
                          ByVal ifMerged As Long, _
                          ByVal asOfDate As String, _
                          ByVal taxRate As Double, _
                          ByVal publicationPins As Object)
    Dim k As Variant
    m_CompanyCode = companyCode
    Set m_Versions = versions
    Set m_FactData = factData
    m_IfMerged = ifMerged
    m_AsOfDate = asOfDate
    m_TaxRate = taxRate
    Set m_PublicationPins = CreateObject("Scripting.Dictionary")
    If Not publicationPins Is Nothing Then
        For Each k In publicationPins.Keys
            m_PublicationPins.Add CStr(k), CStr(publicationPins(k))
        Next k
    End If
    m_IsSharedView = True
End Sub

Public Function CloneView() As clsFinancialStatement
    Dim cloned As clsFinancialStatement, dataset As clsDataSet
    Set cloned = New clsFinancialStatement
    For Each dataset In m_Sources: cloned.AddSource dataset: Next dataset
    cloned.InitSharedView m_CompanyCode, m_Versions, m_FactData, m_IfMerged, m_AsOfDate, m_TaxRate, m_PublicationPins
    Set CloneView = cloned
End Function

Public Sub AddSource(ByVal dataset As clsDataSet)
    If m_IsSharedView Then Err.Raise 5, , "Cannot modify a shared read-only statement view."
    If dataset Is Nothing Then Err.Raise 5, , "Source dataset required."
    m_Sources.Add dataset
End Sub

Public Property Get SourceDataSets() As Collection
    Dim result As New Collection, dataset As clsDataSet
    For Each dataset In m_Sources: result.Add dataset: Next dataset
    Set SourceDataSets = result
End Property

Public Property Get ifMerged() As Long
    ifMerged = m_IfMerged
End Property

Public Property Get asOfDate() As String
    asOfDate = m_AsOfDate
End Property

Public Property Get IsSharedView() As Boolean
    IsSharedView = m_IsSharedView
End Property

Public Property Get companyCode() As String
    companyCode = m_CompanyCode
End Property

Public Property Let companyCode(ByVal Value As String)
    If m_IsSharedView Then Err.Raise 5, "clsFinancialStatement", "Cannot modify a shared read-only statement view."
    If m_Versions.count > 0 Then Err.Raise 5, "clsFinancialStatement", "Company cannot change after loading."
    m_CompanyCode = Trim$(Value)
End Property

' Compatibility alias only: the value is now a company code, not a security code.
Public Property Get SecurityID() As String
    SecurityID = m_CompanyCode
End Property

' Publication pins belong to one business key, never to every fiscal period.
Public Sub ConfigureView(ByVal ifMerged As Long, ByVal asOfDate As String, Optional ByVal publishedDate As String = "", Optional ByVal rptDate As String = "", Optional ByVal ifAdj As Long = 2)
    Dim cutoff As String, publication As String, businessKey As String
    If ifMerged <> 1 And ifMerged <> 2 Then Err.Raise 5, , "IfMerged must be 1 or 2."
    cutoff = Format$(CDate(asOfDate), "yyyy-mm-dd")
    If Len(publishedDate) > 0 Then
        If Len(rptDate) = 0 Then Err.Raise 5, , "An exact publication requires its report period."
        publication = Format$(CDate(publishedDate), "yyyy-mm-dd")
        If publication > cutoff Then Err.Raise 5, , "Publication date exceeds the as-of date."
        businessKey = reportKey(rptDate, ifMerged, ifAdj)
    End If
    m_IfMerged = ifMerged
    m_AsOfDate = cutoff
    m_PublicationPins.RemoveAll
    If Len(businessKey) > 0 Then m_PublicationPins.Add businessKey, publication
End Sub

Public Function RegisterReport(ByVal rptDate As String, _
    ByVal ifMerged As Long, _
    ByVal ifAdj As Long, _
    ByVal publishedDate As String) As String

    Dim key As String, period As String, publ As String
    If m_IsSharedView Then Err.Raise 5, "clsFinancialStatement", "Cannot register reports on a shared read-only statement view."
    period = FinancialPeriodEnd(rptDate)
    publ = Format$(CDate(publishedDate), "yyyy-mm-dd")
    If Len(m_CompanyCode) = 0 Then Err.Raise 5, , "Company code is required."
    key = reportKey(period, ifMerged, ifAdj) & "|" & publ
    If m_Versions.exists(key) Then Err.Raise vbObjectError + 2100, , "Duplicate report identity: " & key
    m_Versions.Add key, Array(period, ifMerged, ifAdj, publ)
    RegisterReport = key
End Function

Public Sub AddReportFact(ByVal reportKey As String, _
        ByVal accountCode As String, _
        ByVal amount As Variant, _
        Optional ByVal overwrite As Boolean = False)
    Dim key As String
    If m_IsSharedView Then Err.Raise 5, "clsFinancialStatement", "Cannot add facts to a shared read-only statement view."
    If Not m_Versions.exists(reportKey) Then Err.Raise 5, , "Unknown report identity."
    key = reportKey & "|" & UCase$(Trim$(accountCode))
    If m_FactData.exists(key) And Not overwrite Then Err.Raise vbObjectError + 2101, , "Duplicate fact: " & key
    If IsEmpty(amount) Then amount = Null
    If Not IsNull(amount) Then amount = CDbl(amount)
    m_FactData(key) = amount
End Sub

' Four fields identify a business report; publication identifies its version.
Public Function reportKey(ByVal rptDate As String, ByVal ifMerged As Long, ByVal ifAdj As Long) As String
    If Len(m_CompanyCode) = 0 Then Err.Raise 5, , "Company code is required."
    If ifMerged <> 1 And ifMerged <> 2 Then Err.Raise 5, , "IfMerged must be 1 or 2."
    If ifAdj <> 1 And ifAdj <> 2 Then Err.Raise 5, , "IfAdjusted must be 1 or 2."
    reportKey = m_CompanyCode & "|" & FinancialPeriodEnd(rptDate) & "|" & CStr(ifMerged) & "|" & CStr(ifAdj)
End Function

Public Function SelectedVersionKey(ByVal rptDate As String, Optional ByVal ifAdj As Long = 2) As String
    Dim k As Variant, v As Variant, latest As String, businessKey As String, exactDate As String
    businessKey = reportKey(rptDate, m_IfMerged, ifAdj)
    If m_PublicationPins.exists(businessKey) Then exactDate = m_PublicationPins(businessKey)
    For Each k In m_Versions.Keys
        v = m_Versions(k)
        If Left$(CStr(k), InStrRev(CStr(k), "|") - 1) = businessKey Then
            If v(3) <= m_AsOfDate Then
                If Len(exactDate) = 0 Or v(3) = exactDate Then
                    If v(3) > latest Then
                        latest = v(3)
                        SelectedVersionKey = CStr(k)
                    End If
                End If
            End If
        End If
    Next k
End Function

' Compatibility alias: this returns the five-field storage version identity.
Public Function SelectedReportKey(ByVal rptDate As String, Optional ByVal ifAdj As Long = 2) As String
    SelectedReportKey = SelectedVersionKey(rptDate, ifAdj)
End Function

' Exact lookup never changes publication pins or falls back to another batch.
Public Function FindExactVersion(ByVal rptDate As String, ByVal ifAdj As Long, ByVal publishedDate As String) As String
    Dim key As String, publication As String
    If Len(publishedDate) = 0 Then Exit Function
    publication = Format$(CDate(publishedDate), "yyyy-mm-dd")
    If publication > m_AsOfDate Then Exit Function
    key = reportKey(rptDate, m_IfMerged, ifAdj) & "|" & publication
    If m_Versions.exists(key) Then FindExactVersion = key
End Function

Public Function SelectedComparableVersionKey(ByVal rptDate As String, ByVal selectedVersion As String) As String
    Dim v As Variant
    If Not IsVisibleVersion(selectedVersion) Then Exit Function
    v = m_Versions(selectedVersion)
    If FinancialPeriodEnd(rptDate) <> Format$(DateAdd("yyyy", -1, CDate(v(0))), "yyyy-mm-dd") Then
        Err.Raise 5, "clsFinancialStatement", "Comparable period must be exactly one year earlier."
    End If
    SelectedComparableVersionKey = FindExactVersion(rptDate, 1, CStr(v(3)))
End Function

Private Function IsVisibleVersion(ByVal versionKey As String) As Boolean
    Dim v As Variant
    If Not m_Versions.exists(versionKey) Then Exit Function
    v = m_Versions(versionKey)
    If CLng(v(1)) <> m_IfMerged Or CStr(v(3)) > m_AsOfDate Then Exit Function
    If versionKey <> reportKey(CStr(v(0)), CLng(v(1)), CLng(v(2))) & "|" & CStr(v(3)) Then Exit Function
    IsVisibleVersion = True
End Function

Public Function TryGetVersionRawValue(ByVal versionKey As String, ByVal accCode As String, ByRef amount As Variant) As Boolean
    Dim key As String
    amount = Null
    If Not IsVisibleVersion(versionKey) Then Exit Function
    key = versionKey & "|" & UCase$(Trim$(accCode))
    If Not m_FactData.exists(key) Then Exit Function
    amount = m_FactData(key)
    TryGetVersionRawValue = Not IsNull(amount)
End Function

Public Function SelectedPublicationDate(ByVal rptDate As String, Optional ByVal ifAdj As Long = 2) As String
    Dim versionKey As String
    versionKey = SelectedVersionKey(rptDate, ifAdj)
    If Len(versionKey) > 0 Then SelectedPublicationDate = CStr(m_Versions(versionKey)(3))
End Function

' Each item: business key, publication date, version key. Only visible versions.
' Collection order is source order; callers must not assume chronological order.
Public Function ReportVersions(ByVal rptDate As String, Optional ByVal ifAdj As Long = 2) As Collection
    Dim result As New Collection, businessKey As String, k As Variant, v As Variant
    businessKey = reportKey(rptDate, m_IfMerged, ifAdj)
    For Each k In m_Versions.Keys
        v = m_Versions(k)
        If Left$(CStr(k), InStrRev(CStr(k), "|") - 1) = businessKey Then
            If v(3) <= m_AsOfDate Then result.Add Array(businessKey, v(3), CStr(k))
        End If
    Next k
    Set ReportVersions = result
End Function

' Pin an optional known batch for one period; blank restores latest-as-of selection.
Public Sub SelectReportVersion(ByVal rptDate As String, Optional ByVal ifAdj As Long = 2, Optional ByVal publishedDate As String = "")
    Dim businessKey As String, publication As String
    businessKey = reportKey(rptDate, m_IfMerged, ifAdj)
    If Len(publishedDate) > 0 Then
        publication = Format$(CDate(publishedDate), "yyyy-mm-dd")
        If publication > m_AsOfDate Then Err.Raise 5, , "Publication date exceeds the as-of date."
        m_PublicationPins(businessKey) = publication
    ElseIf m_PublicationPins.exists(businessKey) Then
        m_PublicationPins.Remove businessKey
    End If
End Sub

Public Function TryGetRawValue(ByVal rptDate As String, ByVal accCode As String, ByRef amount As Variant, Optional ByVal ifAdj As Long = 2) As Boolean
    Dim reportKey As String, key As String
    amount = Null
    reportKey = SelectedReportKey(rptDate, ifAdj)
    If Len(reportKey) = 0 Then Exit Function
    key = reportKey & "|" & UCase$(Trim$(accCode))
    If Not m_FactData.exists(key) Then Exit Function
    amount = m_FactData(key)
    TryGetRawValue = Not IsNull(amount)
End Function

Public Function GetRawValue(ByVal rptDate As String, ByVal accCode As String, Optional ByVal ifAdj As Integer = 2) As Double
    Dim amount As Variant
    If Not TryGetRawValue(rptDate, accCode, amount, ifAdj) Then
        Err.Raise vbObjectError + 2102, "clsFinancialStatement", "Missing or NULL fact: " & rptDate & " / " & accCode & " / adjusted=" & CStr(ifAdj)
    End If
    GetRawValue = CDbl(amount)
End Function

Public Function GetRestatedGrowth(ByVal curDate As String, ByVal priDate As String, ByVal accCode As String, Optional ByVal curAdjusted As Integer = 2, Optional ByVal priorAdjusted As Integer = 2) As Variant
    Dim curVal As Variant, baseVal As Variant
    If Not TryGetRawValue(curDate, accCode, curVal, curAdjusted) Then
        GetRestatedGrowth = CVErr(xlErrNA)
    ElseIf Not TryGetRawValue(priDate, accCode, baseVal, priorAdjusted) Then
        GetRestatedGrowth = CVErr(xlErrNA)
    ElseIf Abs(CDbl(baseVal)) < 0.0000001 Then
        GetRestatedGrowth = CVErr(xlErrDiv0)
    Else
        GetRestatedGrowth = (CDbl(curVal) - CDbl(baseVal)) / Abs(CDbl(baseVal))
    End If
End Function

' Columns: company, period end, merged, adjusted, publication date, account, amount.
Public Function ExportLongTable() As Variant
    Dim result() As Variant, k As Variant, parts() As String, row As Long, col As Long
    If m_FactData.count = 0 Then
        ExportLongTable = Empty
        Exit Function
    End If
    ReDim result(1 To m_FactData.count, 1 To 7)
    For Each k In m_FactData.Keys
        row = row + 1
        parts = Split(CStr(k), "|")
        For col = 0 To 5
            result(row, col + 1) = parts(col)
        Next col
        result(row, 7) = m_FactData(k)
    Next k
    ExportLongTable = result
End Function

' Business-facing series: one selected version per period, all accounts together.
Public Function ExportSelectedLongTable(Optional ByVal ifAdj As Long = 2) As Variant
    Dim rows As New Collection, k As Variant, parts() As String, result() As Variant
    Dim selected As Object, businessKey As String, row As Long, col As Long, item As Variant
    Set selected = CreateObject("Scripting.Dictionary")
    For Each k In m_Versions.Keys
        parts = Split(CStr(k), "|")
        If CLng(parts(2)) = m_IfMerged And CLng(parts(3)) = ifAdj Then
            businessKey = reportKey(parts(1), m_IfMerged, ifAdj)
            If Not selected.exists(businessKey) Then selected.Add businessKey, SelectedVersionKey(parts(1), ifAdj)
        End If
    Next k
    For Each k In m_FactData.Keys
        parts = Split(CStr(k), "|")
        businessKey = parts(0) & "|" & parts(1) & "|" & parts(2) & "|" & parts(3)
        If selected.exists(businessKey) Then
            If Left$(CStr(k), InStrRev(CStr(k), "|") - 1) = selected(businessKey) Then
                rows.Add Array(parts(0), parts(1), parts(2), parts(3), parts(4), parts(5), m_FactData(k))
            End If
        End If
    Next k
    If rows.count = 0 Then
        ExportSelectedLongTable = Empty
        Exit Function
    End If
    ReDim result(1 To rows.count, 1 To 7)
    For row = 1 To rows.count
        item = rows(row)
        For col = 1 To 7
            result(row, col) = item(col - 1)
        Next col
    Next row
    ExportSelectedLongTable = result
End Function
