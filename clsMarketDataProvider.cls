VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMarketDataProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IValueProvider

Private m_Name As String
Private m_Series As Object ' symbol -> Dictionary(yyyy-mm-dd -> Double)
Private m_AllowAsOfFallback As Boolean

Private Sub Class_Initialize()
    m_Name = "MARKET_DATA"
    Set m_Series = CreateObject("Scripting.Dictionary")
    m_Series.CompareMode = 1
    m_AllowAsOfFallback = True
End Sub

Public Sub Configure(ByVal ProviderName As String, Optional ByVal allowAsOfFallback As Boolean = True)
    If Len(Trim$(ProviderName)) > 0 Then m_Name = UCase$(Trim$(ProviderName))
    m_AllowAsOfFallback = allowAsOfFallback
End Sub

Public Sub SetPoint(ByVal symbol As String, ByVal obsDate As String, ByVal Value As Double)
    Dim dateDict As Object, dKey As String
    symbol = UCase$(Trim$(symbol))
    dKey = Format$(CDate(obsDate), "yyyy-mm-dd")
    If Not m_Series.exists(symbol) Then
        Set dateDict = CreateObject("Scripting.Dictionary")
        m_Series.Add symbol, dateDict
    End If
    Set dateDict = m_Series(symbol)
    dateDict(dKey) = Value
End Sub

Private Property Get IValueProvider_ProviderName() As String
    IValueProvider_ProviderName = m_Name
End Property

Private Function IValueProvider_TryResolve(ByVal symbol As String, _
                                           ByVal evalDate As String, _
                                           ByVal contextState As Variant, _
                                           ByRef outValue As Variant, _
                                           ByRef outStatus As String, _
                                           ByRef outSourceKey As String) As Boolean
    Dim dateDict As Object, targetDay As String, bestDay As String, k As Variant
    symbol = UCase$(Trim$(symbol))
    If Not m_Series.exists(symbol) Then Exit Function

    Set dateDict = m_Series(symbol)
    targetDay = Format$(CDate(evalDate), "yyyy-mm-dd")

    If dateDict.exists(targetDay) Then
        outValue = CDbl(dateDict(targetDay))
        outStatus = "OK"
        outSourceKey = m_Name & "|" & symbol & "|" & targetDay
        IValueProvider_TryResolve = True
        Exit Function
    End If

    If m_AllowAsOfFallback Then
        For Each k In dateDict.Keys
            If CStr(k) <= targetDay Then
                If CStr(k) > bestDay Then bestDay = CStr(k)
            End If
        Next k
        If Len(bestDay) > 0 Then
            outValue = CDbl(dateDict(bestDay))
            outStatus = "AS_OF_FALLBACK"
            outSourceKey = m_Name & "|" & symbol & "|" & bestDay
            IValueProvider_TryResolve = True
        End If
    End If
End Function
