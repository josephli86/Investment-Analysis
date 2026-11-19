VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConnectionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const adStateOpen As Long = 1
Private Const PING_COOLDOWN_SEC As Double = 30#

Private m_Connections As Object
Private m_LastPing As Object
Private m_Config As clsConfig

Private Sub Class_Initialize()
    Set m_Connections = CreateObject("Scripting.Dictionary")
    Set m_LastPing = CreateObject("Scripting.Dictionary")
    Set m_Config = New clsConfig
End Sub

Private Sub Class_Terminate()
    CloseAll
    Set m_Connections = Nothing
    Set m_LastPing = Nothing
    Set m_Config = Nothing
End Sub

Public Function GetConnection(Optional ByVal sourceAlias As String = "") As Object
    Dim targetAlias As String
    If Trim$(sourceAlias) = "" Then
        targetAlias = m_Config.GetSourceAliasByModule("FINANCIAL")
    Else
        targetAlias = UCase$(Trim$(sourceAlias))
    End If

    Dim conn As Object
    If m_Connections.exists(targetAlias) Then
        Set conn = m_Connections(targetAlias)
        If Not IsConnectionAlive(targetAlias, conn) Then
            On Error Resume Next
            If conn.State = adStateOpen Then conn.Close
            On Error GoTo 0
            m_Connections.Remove targetAlias
            If m_LastPing.exists(targetAlias) Then m_LastPing.Remove targetAlias
            Set conn = CreateNewConnection(targetAlias)
            m_Connections.Add targetAlias, conn
            m_LastPing(targetAlias) = Timer
        End If
    Else
        Set conn = CreateNewConnection(targetAlias)
        m_Connections.Add targetAlias, conn
        m_LastPing(targetAlias) = Timer
    End If

    Set GetConnection = conn
End Function

Private Function CreateNewConnection(ByVal targetAlias As String) As Object
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionTimeout = 15
    conn.CommandTimeout = 120
    conn.CursorLocation = 3 ' adUseClient

    Dim connStr As String
    connStr = m_Config.GetOracleConnString(targetAlias)

    On Error GoTo ErrConn
    conn.Open connStr
    Set CreateNewConnection = conn
    Exit Function

ErrConn:
    Dim errMsg As String, providerErrors As String, item As Object
    errMsg = "Oracle connection failed. source=" & targetAlias & "; number=" & CStr(Err.number) & "; description=" & Err.description
    On Error Resume Next
    For Each item In conn.Errors
        If Len(providerErrors) > 0 Then providerErrors = providerErrors & " | "
        providerErrors = providerErrors & "number=" & CStr(item.number) & "; description=" & CStr(item.description)
    Next item
    On Error GoTo 0
    If Len(providerErrors) > 0 Then errMsg = errMsg & "; provider_errors=" & providerErrors
    Debug.Print "[clsConnectionManager] " & errMsg
    Err.Raise vbObjectError + 1001, "clsConnectionManager.CreateNewConnection", errMsg
End Function

Private Function IsConnectionAlive(ByVal targetAlias As String, ByRef conn As Object) As Boolean
    Dim elapsed As Double
    If conn Is Nothing Then Exit Function
    If conn.State <> adStateOpen Then Exit Function

    If m_LastPing.exists(targetAlias) Then
        elapsed = Timer - CDbl(m_LastPing(targetAlias))
        If elapsed >= 0# And elapsed < PING_COOLDOWN_SEC Then
            IsConnectionAlive = True
            Exit Function
        End If
    End If

    On Error GoTo HeartbeatFailed
    conn.Execute "SELECT 1 FROM DUAL"
    m_LastPing(targetAlias) = Timer
    IsConnectionAlive = True
    Exit Function

HeartbeatFailed:
    IsConnectionAlive = False
End Function

Public Sub CloseAll()
    Dim k As Variant, conn As Object
    For Each k In m_Connections.Keys
        Set conn = m_Connections(k)
        On Error Resume Next
        If Not conn Is Nothing Then
            If conn.State = adStateOpen Then conn.Close
        End If
        On Error GoTo 0
    Next k
    m_Connections.RemoveAll
    m_LastPing.RemoveAll
End Sub
