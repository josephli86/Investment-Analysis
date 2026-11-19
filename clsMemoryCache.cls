VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMemoryCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const DEFAULT_TTL_SECONDS As Long = 1800
Private Const MAX_CAPACITY As Long = 1000

Private m_Store As Object
Private m_DefaultTTL As Long
Private m_Hits As Long
Private m_Misses As Long

Private Sub Class_Initialize()
    Set m_Store = CreateObject("Scripting.Dictionary")
    m_DefaultTTL = DEFAULT_TTL_SECONDS
    m_Hits = 0
    m_Misses = 0
End Sub

Private Sub Class_Terminate()
    Clear
    Set m_Store = Nothing
End Sub

Public Property Get HitRate() As Double
    Dim total As Long: total = m_Hits + m_Misses
    If total > 0 Then HitRate = CDbl(m_Hits) / CDbl(total) Else HitRate = 0#
End Property

Public Function TryGet(ByVal key As String, ByRef outVal As Variant) As Boolean
    key = UCase(Trim(key))
    If Not m_Store.exists(key) Then
        m_Misses = m_Misses + 1
        TryGet = False
        Exit Function
    End If
    
    Dim node As Variant: node = m_Store(key)
    Dim expireAt As Date: expireAt = CDate(node(1))
    
    If expireAt > 0 And Now > expireAt Then
        m_Store.Remove key
        m_Misses = m_Misses + 1
        TryGet = False
        Exit Function
    End If
    
    node(2) = Now
    m_Store(key) = node
    
    If IsObject(node(0)) Then
        Set outVal = node(0)
    Else
        outVal = node(0)
    End If
    
    m_Hits = m_Hits + 1
    TryGet = True
End Function

Public Sub SetItem(ByVal key As String, ByVal val As Variant, Optional ByVal ttlSeconds As Long = -1)
    key = UCase(Trim(key))
    If m_Store.count >= MAX_CAPACITY And Not m_Store.exists(key) Then PruneLRU
    
    Dim actTTL As Long
    actTTL = IIf(ttlSeconds < 0, m_DefaultTTL, ttlSeconds)
    
    Dim expireAt As Date
    expireAt = IIf(actTTL = 0, 0, DateAdd("s", actTTL, Now))
    
    Dim node(2) As Variant
    If IsObject(val) Then Set node(0) = val Else node(0) = val
    node(1) = expireAt
    node(2) = Now
    
    m_Store(key) = node
End Sub

Public Sub InvalidatePrefix(ByVal prefix As String)
    prefix = UCase(Trim(prefix))
    If Len(prefix) = 0 Then Err.Raise 5, "clsMemoryCache.InvalidatePrefix", "Prefix cannot be empty."
    Dim k As Variant, toRemove As Object
    Set toRemove = CreateObject("Scripting.Dictionary")
    
    For Each k In m_Store.Keys
        If InStr(1, CStr(k), prefix, vbBinaryCompare) = 1 Then toRemove.Add k, True
    Next k
    
    For Each k In toRemove.Keys
        m_Store.Remove k
    Next k
End Sub

Public Sub Clear()
    m_Store.RemoveAll
    m_Hits = 0
    m_Misses = 0
End Sub

Private Sub PruneLRU()
    Dim oldestKey As Variant, oldestTime As Date
    Dim k As Variant, first As Boolean: first = True
    
    For Each k In m_Store.Keys
        Dim nodeTime As Date: nodeTime = CDate(m_Store(k)(2))
        If first Or nodeTime < oldestTime Then
            oldestTime = nodeTime
            oldestKey = k
            first = False
        End If
    Next k
    
    If Not IsEmpty(oldestKey) Then m_Store.Remove oldestKey
End Sub
