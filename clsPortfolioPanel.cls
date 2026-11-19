VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPortfolioPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_PortCodes As String
Private m_Start As Date
Private m_End As Date
Private m_Source As String
Private m_HoldingsFrame As clsTimeSeriesFrame
Private m_PerformanceFrame As clsTimeSeriesFrame
Private m_ProfitFrame As clsTimeSeriesFrame
Private m_OrdersData As clsDataTable
Private m_IndustryData As clsDataTable
Private m_Ready As Boolean
Private m_Sources As Collection

Public Sub Load(ByVal portCodes As String, ByVal startDate As Date, ByVal endDate As Date, ByVal sourceAlias As String, ByVal holdings As clsTimeSeriesFrame, ByVal performance As clsTimeSeriesFrame, ByVal profit As clsTimeSeriesFrame, ByVal orders As clsDataTable, ByVal industry As clsDataTable, Optional ByVal sources As Collection = Nothing)
    Dim dataset As clsDataSet
    If m_Ready Then Err.Raise 5, , "Portfolio panel is immutable."
    If holdings Is Nothing Or performance Is Nothing Or profit Is Nothing Or orders Is Nothing Or industry Is Nothing Then Err.Raise 5, "clsPortfolioPanel.Load", "All portfolio data tables are required."
    If Not holdings.IsLoaded Or Not performance.IsLoaded Or Not profit.IsLoaded Or Not orders.IsLoaded Or Not industry.IsLoaded Then Err.Raise 5, "clsPortfolioPanel.Load", "All portfolio data tables must be loaded."
    If holdings.DateField <> "T_DATE" Or performance.DateField <> "T_DATE" Or profit.DateField <> "T_DATE" Then Err.Raise 5, "clsPortfolioPanel.Load", "Portfolio time-series date field must be T_DATE."
    If dateValue(startDate) > dateValue(endDate) Then Err.Raise 5, "clsPortfolioPanel.Load", "Start date is after end date."
    m_PortCodes = portCodes
    m_Start = dateValue(startDate)
    m_End = dateValue(endDate)
    m_Source = sourceAlias
    Set m_HoldingsFrame = holdings
    Set m_PerformanceFrame = performance
    Set m_ProfitFrame = profit
    Set m_OrdersData = orders
    Set m_IndustryData = industry
    Set m_Sources = New Collection
    If Not sources Is Nothing Then
        For Each dataset In sources: m_Sources.Add dataset: Next dataset
    End If
    m_Ready = True
End Sub

Public Property Get SourceDataSets() As Collection
    Dim result As New Collection, dataset As clsDataSet
    If Not m_Sources Is Nothing Then
        For Each dataset In m_Sources: result.Add dataset: Next dataset
    End If
    Set SourceDataSets = result
End Property

Public Property Get PortfolioCodes() As String
    PortfolioCodes = m_PortCodes
End Property

Public Property Get startDate() As Date
    startDate = m_Start
End Property

Public Property Get endDate() As Date
    endDate = m_End
End Property

Public Property Get sourceAlias() As String
    sourceAlias = m_Source
End Property

Public Property Get HoldingsTable() As Variant
    HoldingsTable = m_HoldingsFrame.ToTable
End Property

Public Property Get PerformanceTable() As Variant
    PerformanceTable = m_PerformanceFrame.ToTable
End Property

Public Property Get ProfitTable() As Variant
    ProfitTable = m_ProfitFrame.ToTable
End Property

Public Property Get ordersTable() As Variant
    ordersTable = m_OrdersData.ToTable
End Property

Public Property Get industryTable() As Variant
    industryTable = m_IndustryData.ToTable
End Property

Public Property Get holdingsFrame() As clsTimeSeriesFrame
    Set holdingsFrame = m_HoldingsFrame
End Property

Public Property Get performanceFrame() As clsTimeSeriesFrame
    Set performanceFrame = m_PerformanceFrame
End Property

Public Property Get profitFrame() As clsTimeSeriesFrame
    Set profitFrame = m_ProfitFrame
End Property

Public Property Get ordersData() As clsDataTable
    Set ordersData = m_OrdersData
End Property

Public Property Get industryData() As clsDataTable
    Set industryData = m_IndustryData
End Property

Public Property Get HoldingRowCount() As Long
    HoldingRowCount = m_HoldingsFrame.rowCount
End Property

Public Property Get PerformanceRowCount() As Long
    PerformanceRowCount = m_PerformanceFrame.rowCount
End Property

Public Property Get ProfitRowCount() As Long
    ProfitRowCount = m_ProfitFrame.rowCount
End Property

Public Property Get OrderRowCount() As Long
    OrderRowCount = m_OrdersData.rowCount
End Property

Public Property Get IndustryRowCount() As Long
    IndustryRowCount = m_IndustryData.rowCount
End Property
