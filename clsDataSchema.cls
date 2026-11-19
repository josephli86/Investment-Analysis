VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDataSchema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Name As String, m_Version As String, m_Grain As String
Private m_Fields As Object, m_Frozen As Boolean

Public Sub Initialize(ByVal datasetName As String, ByVal SchemaVersion As String, ByVal grain As String)
    If Not m_Fields Is Nothing Then Err.Raise 5, , "Schema already initialized."
    If Len(Trim$(datasetName)) = 0 Or Len(Trim$(SchemaVersion)) = 0 Or Len(Trim$(grain)) = 0 Then Err.Raise 5, , "Schema identity and grain required."
    m_Name = datasetName: m_Version = SchemaVersion: m_Grain = grain
    Set m_Fields = CreateObject("Scripting.Dictionary")
    m_Fields.CompareMode = 1
End Sub

Public Sub AddField(ByVal fieldName As String, ByVal kind As String, Optional ByVal nullable As Boolean = True, Optional ByVal unit As String = "")
    If m_Fields Is Nothing Then Err.Raise 5, , "Schema not initialized."
    If m_Frozen Then Err.Raise 5, , "Schema is frozen."
    fieldName = UCase$(Trim$(fieldName)): kind = UCase$(Trim$(kind))
    If Len(fieldName) = 0 Then Err.Raise 5, , "Field name required."
    Select Case kind
        Case "ANY", "TEXT", "DATE", "NUMBER", "DOUBLE", "DECIMAL", "LONG", "BOOLEAN"
        Case Else: Err.Raise 5, , "Unsupported field type: " & kind
    End Select
    m_Fields.Add fieldName, Array(kind, nullable, unit)
End Sub

Public Sub Freeze()
    If m_Fields Is Nothing Then Err.Raise 5, , "Schema not initialized."
    If m_Fields.count = 0 Then Err.Raise 5, , "Schema must declare fields."
    m_Frozen = True
End Sub

Public Property Get datasetName() As String
    datasetName = m_Name
End Property
Public Property Get SchemaVersion() As String
    SchemaVersion = m_Version
End Property
Public Property Get grain() As String
    grain = m_Grain
End Property
Public Function FieldUnit(ByVal fieldName As String) As String
    FieldUnit = CStr(m_Fields(fieldName)(2))
End Function

' Required fields are declared; undeclared source columns are preserved.
Public Function Normalize(ByVal data As clsDataTable) As clsDataTable
    Dim table As Variant, result As clsDataTable, key As Variant, rule As Variant
    Dim r As Long, c As Long, Value As Variant, kind As String
    Freeze
    If data Is Nothing Then Err.Raise 5, , "Data required."
    table = data.ToTable()
    For Each key In m_Fields.Keys
        c = data.FieldIndex(CStr(key)): rule = m_Fields(key): kind = CStr(rule(0))
        For r = 2 To data.rowCount + 1
            Value = table(r, c)
            If IsError(Value) Then Err.Raise 13, , "Error cell: " & CStr(key) & " row " & CStr(r - 1)
            If IsNull(Value) Or IsEmpty(Value) Then
                If Not CBool(rule(1)) Then Err.Raise 5, , "Missing required value: " & CStr(key) & " row " & CStr(r - 1)
            Else
                Select Case kind
                    Case "TEXT": Value = CStr(Value)
                    Case "DATE"
                        If VarType(Value) <> vbDate Then
                            If IsDate(Value) Then
                                Value = CDate(Value)
                            Else
                                Err.Raise 13, , "Date value required: " & CStr(key) & "; VarType=" & CStr(VarType(Value))
                            End If
                        End If
                    Case "NUMBER", "DOUBLE", "DECIMAL", "LONG"
                        If VarType(Value) = vbBoolean Or VarType(Value) = vbDate Then Err.Raise 13, , "Numeric value required: " & CStr(key)
                        If Not IsNumeric(Value) Then Err.Raise 13, , "Numeric value required: " & CStr(key)
                        If kind = "DOUBLE" Or kind = "NUMBER" Then Value = CDbl(Value)
                        If kind = "DECIMAL" Then Value = CDec(Value)
                        If kind = "LONG" Then
                            If CDec(Value) <> Fix(CDec(Value)) Then Err.Raise 13, , "Integer required: " & CStr(key)
                            Value = CLng(Value)
                        End If
                    Case "BOOLEAN"
                        If VarType(Value) <> vbBoolean Then Err.Raise 13, , "Boolean required: " & CStr(key)
                End Select
                table(r, c) = Value
            End If
        Next r
    Next key
    Set result = New clsDataTable: result.Load table
    Set Normalize = result
End Function
