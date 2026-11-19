VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDataTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Table As Variant
Private m_Columns As Object
Private m_RowCount As Long
Private m_FieldCount As Long
Private m_Loaded As Boolean

Public Sub Load(ByVal table As Variant)
    Dim r As Long, c As Long, fieldName As String

    If m_Loaded Then Err.Raise 5, "clsDataTable.Load", "Data tables are immutable."
    If Not IsArray(table) Then Err.Raise 5, "clsDataTable.Load", "A two-dimensional table is required."
    ValidateShape table

    m_RowCount = UBound(table, 1) - 1
    m_FieldCount = UBound(table, 2)
    If m_FieldCount < 1 Then Err.Raise 5, "clsDataTable.Load", "At least one table field is required."

    m_Table = table
    Set m_Columns = CreateObject("Scripting.Dictionary")
    m_Columns.CompareMode = 1
    For c = 1 To m_FieldCount
        If IsNull(m_Table(1, c)) Or IsEmpty(m_Table(1, c)) Then Err.Raise 5, "clsDataTable.Load", "Table field names cannot be blank."
        fieldName = UCase$(Trim$(CStr(m_Table(1, c))))
        If Len(fieldName) = 0 Then Err.Raise 5, "clsDataTable.Load", "Table field names cannot be blank."
        If m_Columns.exists(fieldName) Then Err.Raise 5, "clsDataTable.Load", "Duplicate table field: " & fieldName
        m_Table(1, c) = fieldName
        m_Columns.Add fieldName, c
        For r = 2 To m_RowCount + 1
            If IsObject(m_Table(r, c)) Or IsArray(m_Table(r, c)) Then Err.Raise 13, "clsDataTable.Load", "Table cells must contain scalar values."
        Next r
    Next c
    m_Loaded = True
End Sub

Public Property Get IsLoaded() As Boolean
    IsLoaded = m_Loaded
End Property

Public Property Get rowCount() As Long
    rowCount = m_RowCount
End Property

Public Property Get fieldCount() As Long
    fieldCount = m_FieldCount
End Property

Public Function HasField(ByVal fieldName As String) As Boolean
    RequireLoaded
    HasField = m_Columns.exists(UCase$(Trim$(fieldName)))
End Function

Public Function FieldIndex(ByVal fieldName As String) As Long
    RequireLoaded
    fieldName = UCase$(Trim$(fieldName))
    If Not m_Columns.exists(fieldName) Then Err.Raise 5, "clsDataTable.FieldIndex", "Missing table field: " & fieldName
    FieldIndex = CLng(m_Columns(fieldName))
End Function

Public Property Get ValueAt(ByVal rowNumber As Long, ByVal fieldName As String) As Variant
    ValidateRow rowNumber
    ValueAt = m_Table(rowNumber + 1, FieldIndex(fieldName))
End Property

Public Property Get ValueAtColumn(ByVal rowNumber As Long, ByVal columnNumber As Long) As Variant
    ValidateRow rowNumber
    If columnNumber < 1 Or columnNumber > m_FieldCount Then Err.Raise 9, "clsDataTable", "Column index is out of range."
    ValueAtColumn = m_Table(rowNumber + 1, columnNumber)
End Property

Public Property Get fieldName(ByVal columnNumber As Long) As String
    If columnNumber < 1 Or columnNumber > m_FieldCount Then Err.Raise 9, "clsDataTable", "Column index is out of range."
    fieldName = CStr(m_Table(1, columnNumber))
End Property

Public Function ToTable() As Variant
    RequireLoaded
    ToTable = m_Table
End Function

Public Function SelectRows(ByVal rows As Collection) As clsDataTable
    Dim selected() As Variant, result As clsDataTable
    Dim r As Long, c As Long, sourceRow As Long
    RequireLoaded
    ReDim selected(1 To rows.count + 1, 1 To m_FieldCount)
    For c = 1 To m_FieldCount: selected(1, c) = m_Table(1, c): Next c
    For r = 1 To rows.count
        sourceRow = CLng(rows(r))
        ValidateRow sourceRow
        For c = 1 To m_FieldCount: selected(r + 1, c) = m_Table(sourceRow + 1, c): Next c
    Next r
    Set result = New clsDataTable
    result.Load selected
    Set SelectRows = result
End Function

Public Function FilterField(ByVal fieldName As String, ByVal Value As Variant) As clsDataTable
    Dim rows As New Collection, r As Long, column As Long, token As String
    column = FieldIndex(fieldName)
    token = KeyToken(Value)
    For r = 1 To m_RowCount
        If KeyToken(m_Table(r + 1, column)) = token Then rows.Add r
    Next r
    Set FilterField = SelectRows(rows)
End Function

Public Function GroupRows(ByVal fieldNames As Variant) As Object
    Dim groups As Object, fields As Variant, columns() As Long, rows As Collection
    Dim r As Long, i As Long, groupKey As String, token As String
    If Not IsArray(fieldNames) Then Err.Raise 5, "clsDataTable.GroupRows", "Field names must be an array."
    fields = fieldNames
    If UBound(fields) < LBound(fields) Then Err.Raise 5, "clsDataTable.GroupRows", "At least one grouping field is required."
    ReDim columns(LBound(fields) To UBound(fields))
    For i = LBound(fields) To UBound(fields): columns(i) = FieldIndex(CStr(fields(i))): Next i
    Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = 0
    For r = 1 To m_RowCount
        groupKey = ""
        For i = LBound(columns) To UBound(columns)
            token = KeyToken(m_Table(r + 1, columns(i)))
            groupKey = groupKey & CStr(Len(token)) & ":" & token
        Next i
        If Not groups.exists(groupKey) Then Set rows = New Collection: groups.Add groupKey, rows
        groups(groupKey).Add r
    Next r
    Set GroupRows = groups
End Function

Public Function UniqueRowIndex(ByVal fieldNames As Variant) As Object
    Dim result As Object, groups As Object, key As Variant, rows As Collection
    Set groups = GroupRows(fieldNames)
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 0
    For Each key In groups.Keys
        Set rows = groups(key)
        If rows.count > 1 Then Err.Raise 5, "clsDataTable.UniqueRowIndex", "Duplicate row key; matching rows=" & rows.count
        result.Add CStr(key), CLng(rows(1))
    Next key
    Set UniqueRowIndex = result
End Function

Private Sub ValidateRow(ByVal rowNumber As Long)
    RequireLoaded
    If rowNumber < 1 Or rowNumber > m_RowCount Then Err.Raise 9, "clsDataTable", "Row index is out of range."
End Sub

Private Sub RequireLoaded()
    If Not m_Loaded Then Err.Raise 5, "clsDataTable", "Data table is not loaded."
End Sub

Private Sub ValidateShape(ByRef table As Variant)
    Dim rowLower As Long, columnLower As Long, rowUpper As Long, columnUpper As Long
    Dim hasThirdDimension As Boolean
    On Error GoTo InvalidShape
    rowLower = LBound(table, 1): rowUpper = UBound(table, 1)
    columnLower = LBound(table, 2): columnUpper = UBound(table, 2)
    On Error Resume Next
    Err.Clear
    hasThirdDimension = (UBound(table, 3) >= LBound(table, 3))
    If Err.number <> 0 Then hasThirdDimension = False
    Err.Clear
    On Error GoTo 0
    If hasThirdDimension Then Err.Raise 5, "clsDataTable.Load", "A two-dimensional table is required."
    If rowLower <> 1 Or columnLower <> 1 Or rowUpper < 1 Or columnUpper < 1 Then Err.Raise 5, "clsDataTable.Load", "Table must be 1-based with a header row and at least one field."
    Exit Sub
InvalidShape:
    Err.Raise 5, "clsDataTable.Load", "A two-dimensional table is required."
End Sub

Private Function KeyToken(ByVal Value As Variant) As String
    If IsNull(Value) Then
        KeyToken = "N"
    ElseIf IsEmpty(Value) Then
        KeyToken = "E"
    ElseIf IsError(Value) Then
        Err.Raise 13, "clsDataTable", "An error value cannot be used as a grouping key."
    Else
        KeyToken = "V" & CStr(VarType(Value)) & ":" & CStr(Len(CStr(Value))) & ":" & CStr(Value)
    End If
End Function
