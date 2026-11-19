VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEquityPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_IsLoaded As Boolean
Private m_SecurityCodes As String
Private m_StartDate As Date
Private m_EndDate As Date
Private m_LoadedStartDate As Date
Private m_SourceAlias As String
Private m_MasterTable As clsDataTable
Private m_QuoteFrame As clsTimeSeriesFrame
Private m_Sources As Collection

Public Sub Load(ByVal SecurityCodes As String, _
        ByVal startDate As Date, _
        ByVal endDate As Date, _
        ByVal LoadedStartDate As Date, _
        ByVal sourceAlias As String, _
        ByVal MasterTable As clsDataTable, _
        ByVal quoteFrame As clsTimeSeriesFrame, _
        Optional ByVal SourceDataSets As Collection = Nothing)
    Dim item As Variant
    If m_IsLoaded Then Err.Raise 5, "clsEquityPanel.Load", "Equity panel is immutable once loaded."
    If dateValue(startDate) > dateValue(endDate) Then Err.Raise 5, "clsEquityPanel.Load", "Start date is after end date."
    If dateValue(LoadedStartDate) > dateValue(startDate) Then Err.Raise 5, "clsEquityPanel.Load", "Loaded start date is after start date."
    If MasterTable Is Nothing Or Not MasterTable.IsLoaded Then Err.Raise 5, "clsEquityPanel.Load", "Master table is required."
    If quoteFrame Is Nothing Or Not quoteFrame.IsLoaded Then Err.Raise 5, "clsEquityPanel.Load", "Quote frame is required."

    m_SecurityCodes = Trim$(SecurityCodes)
    m_StartDate = dateValue(startDate)
    m_EndDate = dateValue(endDate)
    m_LoadedStartDate = dateValue(LoadedStartDate)
    m_SourceAlias = UCase$(Trim$(sourceAlias))
    Set m_MasterTable = MasterTable
    Set m_QuoteFrame = quoteFrame
    Set m_Sources = New Collection
    If Not SourceDataSets Is Nothing Then
        For Each item In SourceDataSets
            m_Sources.Add item
        Next item
    End If
    m_IsLoaded = True
End Sub

Public Property Get IsLoaded() As Boolean
    IsLoaded = m_IsLoaded
End Property

Public Property Get SecurityCodes() As String
    RequireLoaded
    SecurityCodes = m_SecurityCodes
End Property

Public Property Get startDate() As Date
    RequireLoaded
    startDate = m_StartDate
End Property

Public Property Get endDate() As Date
    RequireLoaded
    endDate = m_EndDate
End Property

Public Property Get LoadedStartDate() As Date
    RequireLoaded
    LoadedStartDate = m_LoadedStartDate
End Property

Public Property Get sourceAlias() As String
    RequireLoaded
    sourceAlias = m_SourceAlias
End Property

Public Property Get MasterTable() As clsDataTable
    RequireLoaded
    Set MasterTable = m_MasterTable
End Property

Public Property Get quoteFrame() As clsTimeSeriesFrame
    RequireLoaded
    Set quoteFrame = m_QuoteFrame
End Property

Public Property Get SourceDataSets() As Collection
    Dim copy As New Collection, item As Variant
    RequireLoaded
    For Each item In m_Sources
        copy.Add item
    Next item
    Set SourceDataSets = copy
End Property

' 按 SEC_CODE（如 "600519"）或 INNER_CODE 提取单只股票的行情序列
' includeWarmup = True 时包含预热期数据（用于计算 MA250/MACD），False 时仅返回 [StartDate, EndDate]
Public Function QuoteFrameForStock(ByVal secCodeOrInnerCode As Variant, _
        Optional ByVal includeWarmup As Boolean = False) As clsTimeSeriesFrame
    Dim innerCode As Long, filtered As clsTimeSeriesFrame
    RequireLoaded
    innerCode = ResolveInnerCode(secCodeOrInnerCode)
    Set filtered = m_QuoteFrame.FilterField("INNER_CODE", innerCode)
    If Not includeWarmup Then
        Set filtered = filtered.FilterDateRange(m_StartDate, m_EndDate)
    End If
    Set QuoteFrameForStock = filtered.SortedByDate()
End Function

Public Function ResolveInnerCode(ByVal secCodeOrInnerCode As Variant) As Long
    Dim r As Long, codeText As String, dotPos As Long, targetSec As String, targetMkt As String
    RequireLoaded
    If IsNumeric(secCodeOrInnerCode) And VarType(secCodeOrInnerCode) <> vbString Then
        ResolveInnerCode = CLng(secCodeOrInnerCode)
        Exit Function
    End If

    codeText = UCase$(Trim$(CStr(secCodeOrInnerCode)))
    dotPos = InStr(codeText, ".")
    If dotPos > 0 Then
        targetSec = Left$(codeText, dotPos - 1)
        targetMkt = Mid$(codeText, dotPos + 1)
    Else
        targetSec = codeText
        targetMkt = ""
    End If

    For r = 1 To m_MasterTable.rowCount
        If UCase$(Trim$(CStr(m_MasterTable.ValueAt(r, "SEC_CODE")))) = targetSec Then
            If Len(targetMkt) = 0 Or UCase$(Trim$(CStr(m_MasterTable.ValueAt(r, "MARKET_CODE")))) = targetMkt Then
                ResolveInnerCode = CLng(m_MasterTable.ValueAt(r, "INNER_CODE"))
                Exit Function
            End If
        End If
    Next r
    Err.Raise 5, "clsEquityPanel.ResolveInnerCode", "Security not found in loaded equity panel: " & CStr(secCodeOrInnerCode)
End Function

Private Sub RequireLoaded()
    If Not m_IsLoaded Then Err.Raise 5, "clsEquityPanel", "Equity panel is not loaded."
End Sub
