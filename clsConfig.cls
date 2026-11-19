VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_FSO As Object
Private m_BasePath As String
Private m_IniPath As String

Private Sub Class_Initialize()
    Set m_FSO = CreateObject("Scripting.FileSystemObject")
    m_BasePath = ThisWorkbook.Path
    m_IniPath = m_BasePath & "\Config\db_config.ini"
End Sub

Public Property Get BasePath() As String
    BasePath = m_BasePath
End Property

Public Property Get LocalDbPath() As String
    LocalDbPath = GetAccessDbPath(LocalSourceAlias)
End Property

Public Property Get LocalSourceAlias() As String
    LocalSourceAlias = UCase$(Trim$(GetIniVal("ROUTING", "ADJUSTMENT_LOCAL", "")))
End Property

Public Function GetLocalConnString() As String
    GetLocalConnString = GetAccessConnString(LocalSourceAlias)
End Function

Public Function GetAccessDbPath(ByVal sourceAlias As String) As String
    Dim source As String, configuredPath As String
    source = UCase$(Trim$(sourceAlias))
    If source = "" Then Exit Function
    If UCase$(GetIniVal(source, "TYPE", "")) <> "ACCESS" Then
        Err.Raise 5, "clsConfig.GetAccessDbPath", "Access source [" & source & "] requires TYPE=ACCESS."
    End If
    configuredPath = GetIniVal(source, "PATH", "")
    If Len(configuredPath) = 0 Then
        Err.Raise 5, "clsConfig.GetAccessDbPath", "Missing PATH for Access source [" & source & "]."
    End If
    configuredPath = Replace(configuredPath, "/", "\")
    If Left$(configuredPath, 2) = "\\" Or Mid$(configuredPath, 2, 2) = ":\" Then
        GetAccessDbPath = m_FSO.GetAbsolutePathName(configuredPath)
    Else
        If Left$(configuredPath, 1) = "\" Or InStr(configuredPath, ":") > 0 Then
            Err.Raise 5, "clsConfig.GetAccessDbPath", "PATH must be absolute or relative to the workbook folder."
        End If
        GetAccessDbPath = m_FSO.GetAbsolutePathName(m_FSO.BuildPath(m_BasePath, configuredPath))
    End If
End Function

Public Function GetAccessConnString(ByVal sourceAlias As String) As String
    Dim dbPath As String
    dbPath = GetAccessDbPath(sourceAlias)
    If dbPath = "" Then Exit Function
    If Not m_FSO.FileExists(dbPath) Then
        Err.Raise 53, "clsConfig.GetAccessConnString", "Configured Access database not found: " & dbPath
    End If
    GetAccessConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""" & Replace(dbPath, """", """""") & """;"
End Function

Public Function GetSourceAliasByModule(ByVal moduleName As String) As String
    Dim routingKey As String
    routingKey = UCase$(Trim$(moduleName)) & "_SOURCE"
    GetSourceAliasByModule = GetIniVal("ROUTING", routingKey, GetIniVal("ROUTING", "DEFAULT_SOURCE", "ORACLE_FIN"))
End Function

Public Function GetSourceAliasByRoutingKey(ByVal routingKey As String) As String
    Dim key As String
    key = UCase$(Trim$(routingKey))
    If Len(key) = 0 Then Err.Raise 5, "clsConfig.GetSourceAliasByRoutingKey", "Routing key is required."
    GetSourceAliasByRoutingKey = UCase$(Trim$(GetIniVal("ROUTING", key, "")))
End Function

Public Function GetUseMock(Optional ByVal sourceAlias As String = "") As Boolean
    Dim targetSource As String
    targetSource = ResolveSourceAlias(sourceAlias)
    GetUseMock = (GetIniVal(targetSource, "USE_MOCK", "1") = "1")
End Function

Public Function GetOracleConnString(Optional ByVal sourceAlias As String = "") As String
    Dim targetSource As String
    Dim h As String, p As String, s As String, u As String, pwd As String
    targetSource = ResolveSourceAlias(sourceAlias)
    h = GetIniVal(targetSource, "HOST", "127.0.0.1")
    p = GetIniVal(targetSource, "PORT", "1521")
    s = GetIniVal(targetSource, "SERVICE_NAME", "ORCL")
    u = GetIniVal(targetSource, "USER_ID", "invest_user")
    pwd = GetIniVal(targetSource, "PASSWORD", "")
    GetOracleConnString = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)" & _
                          "(HOST=" & h & ")(PORT=" & p & "))(CONNECT_DATA=(SERVICE_NAME=" & s & ")));" & _
                          "User Id=" & u & ";Password=" & pwd & ";"
End Function

Public Function ReadSQLFile(ByVal fileName As String) As String
    Dim fPath As String, stm As Object
    fPath = m_BasePath & "\SQL\" & fileName
    If Not m_FSO.FileExists(fPath) Then Exit Function
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile fPath
    ReadSQLFile = stm.ReadText
    stm.Close
End Function

Private Function ResolveSourceAlias(ByVal alias As String) As String
    If Trim$(alias) = "" Then
        ResolveSourceAlias = GetIniVal("ROUTING", "DEFAULT_SOURCE", "ORACLE_FIN")
    Else
        ResolveSourceAlias = UCase$(Trim$(alias))
    End If
End Function

Private Function GetIniVal(ByVal section As String, ByVal key As String, ByVal defVal As String) As String
    If Not m_FSO.FileExists(m_IniPath) Then
        GetIniVal = defVal
        Exit Function
    End If
    Dim ts As Object, line As String, inSec As Boolean, parts() As String
    Set ts = m_FSO.OpenTextFile(m_IniPath, 1, False)
    Do While Not ts.AtEndOfStream
        line = Trim$(ts.ReadLine)
        If Left$(line, 1) <> ";" And Left$(line, 1) <> "#" Then
            If Left$(line, 1) = "[" And Right$(line, 1) = "]" Then
                inSec = (UCase$(Mid$(line, 2, Len(line) - 2)) = UCase$(section))
            ElseIf inSec And InStr(line, "=") > 0 Then
                parts = Split(line, "=", 2)
                If UCase$(Trim$(parts(0))) = UCase$(key) Then
                    GetIniVal = Trim$(parts(1))
                    ts.Close
                    Exit Function
                End If
            End If
        End If
    Loop
    ts.Close
    GetIniVal = defVal
End Function
