VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSqlBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_ORACLE_IN_ITEMS As Long = 900

Private m_Config As clsConfig
Private m_Template As String
Private m_TemplateName As String
Private m_Params As Object

Private Sub Class_Initialize()
    Set m_Config = New clsConfig
    Set m_Params = CreateObject("Scripting.Dictionary")
    m_Params.CompareMode = 1 ' vbTextCompare: case-insensitive placeholder keys
End Sub

Public Function LoadTemplate(ByVal sqlFileName As String) As clsSqlBuilder
    Dim content As String
    content = m_Config.ReadSQLFile(sqlFileName)
    If Len(Trim$(content)) = 0 Then
        Err.Raise 53, "clsSqlBuilder.LoadTemplate", "Missing or empty SQL template: " & sqlFileName
    End If
    m_TemplateName = sqlFileName
    m_Template = content
    m_Params.RemoveAll
    Set LoadTemplate = Me
End Function

Public Function LoadRawSql(ByVal sqlText As String, Optional ByVal label As String = "INLINE_SQL") As clsSqlBuilder
    If Len(Trim$(sqlText)) = 0 Then
        Err.Raise 5, "clsSqlBuilder.LoadRawSql", "SQL text cannot be empty."
    End If
    m_TemplateName = label
    m_Template = sqlText
    m_Params.RemoveAll
    Set LoadRawSql = Me
End Function

Public Function SetLiteral(ByVal key As String, ByVal rawSqlFragment As String) As clsSqlBuilder
    ValidateKey key
    m_Params(UCase$(Trim$(key))) = rawSqlFragment
    Set SetLiteral = Me
End Function

Public Function SetNumber(ByVal key As String, ByVal numValue As Variant) As clsSqlBuilder
    ValidateKey key
    If IsNull(numValue) Or IsEmpty(numValue) Or Not IsNumeric(numValue) Then
        Err.Raise 5, "clsSqlBuilder.SetNumber", "Numeric value required for placeholder {" & key & "}."
    End If
    m_Params(UCase$(Trim$(key))) = FormatNumberLiteral(numValue)
    Set SetNumber = Me
End Function

Public Function SetString(ByVal key As String, ByVal textValue As String, Optional ByVal wrapQuotes As Boolean = True) As clsSqlBuilder
    Dim escaped As String
    ValidateKey key
    If InStr(textValue, vbCr) > 0 Or InStr(textValue, vbLf) > 0 Then
        Err.Raise 5, "clsSqlBuilder.SetString", "Line break is not allowed in placeholder {" & key & "}."
    End If
    escaped = Replace(textValue, "'", "''")
    If wrapQuotes Then
        m_Params(UCase$(Trim$(key))) = "'" & escaped & "'"
    Else
        m_Params(UCase$(Trim$(key))) = escaped
    End If
    Set SetString = Me
End Function

Public Function SetDate(ByVal key As String, ByVal dateValue As Variant, Optional ByVal wrapQuotes As Boolean = True) As clsSqlBuilder
    Dim formatted As String
    ValidateKey key
    If Not IsDate(dateValue) Then
        Err.Raise 5, "clsSqlBuilder.SetDate", "Valid date required for placeholder {" & key & "}."
    End If
    formatted = Format$(CDate(dateValue), "yyyy-mm-dd")
    If wrapQuotes Then
        m_Params(UCase$(Trim$(key))) = "'" & formatted & "'"
    Else
        m_Params(UCase$(Trim$(key))) = formatted
    End If
    Set SetDate = Me
End Function

Public Function SetAccessDate(ByVal key As String, ByVal dateValue As Variant) As clsSqlBuilder
    Dim formatted As String
    ValidateKey key
    If Not IsDate(dateValue) Then
        Err.Raise 5, "clsSqlBuilder.SetAccessDate", "Valid date required for placeholder {" & key & "}."
    End If
    formatted = Format$(CDate(dateValue), "mm/dd/yyyy")
    m_Params(UCase$(Trim$(key))) = "#" & formatted & "#"
    Set SetAccessDate = Me
End Function

Public Function SetInList(ByVal key As String, ByVal items As Variant, Optional ByVal isNumericList As Boolean = False) As clsSqlBuilder
    Dim uniqueList As Collection, i As Long, result As String
    ValidateKey key
    Set uniqueList = ExtractUniqueItems(items, key)
    If uniqueList.count = 0 Then
        Err.Raise 5, "clsSqlBuilder.SetInList", "At least one item is required for {" & key & "}."
    End If
    If uniqueList.count > MAX_ORACLE_IN_ITEMS Then
        Err.Raise 5, "clsSqlBuilder.SetInList", "Item count (" & CStr(uniqueList.count) & ") exceeds Oracle IN limit; use SetChunkedInPredicate instead."
    End If
    For i = 1 To uniqueList.count
        If i > 1 Then result = result & ","
        result = result & FormatListItem(uniqueList(i), isNumericList, key)
    Next i
    m_Params(UCase$(Trim$(key))) = result
    Set SetInList = Me
End Function

Public Function SetChunkedInPredicate(ByVal key As String, ByVal columnName As String, ByVal items As Variant, Optional ByVal isNumericList As Boolean = False) As clsSqlBuilder
    Dim uniqueList As Collection, i As Long, chunkCount As Long
    Dim chunk As String, predicate As String, totalChunks As Long
    ValidateKey key
    ValidateColumnName columnName, key
    Set uniqueList = ExtractUniqueItems(items, key)
    If uniqueList.count = 0 Then
        m_Params(UCase$(Trim$(key))) = "1 = 0"
        Set SetChunkedInPredicate = Me
        Exit Function
    End If

    For i = 1 To uniqueList.count
        If chunkCount = MAX_ORACLE_IN_ITEMS Then
            If Len(predicate) > 0 Then predicate = predicate & " OR "
            predicate = predicate & columnName & " IN (" & chunk & ")"
            chunk = ""
            chunkCount = 0
            totalChunks = totalChunks + 1
        End If
        If Len(chunk) > 0 Then chunk = chunk & ","
        chunk = chunk & FormatListItem(uniqueList(i), isNumericList, key)
        chunkCount = chunkCount + 1
    Next i

    If chunkCount > 0 Then
        If Len(predicate) > 0 Then predicate = predicate & " OR "
        predicate = predicate & columnName & " IN (" & chunk & ")"
        totalChunks = totalChunks + 1
    End If

    If totalChunks > 1 Then predicate = "(" & predicate & ")"
    m_Params(UCase$(Trim$(key))) = predicate
    Set SetChunkedInPredicate = Me
End Function

Private Sub ValidateColumnName(ByVal columnName As String, ByVal key As String)
    Dim rx As Object
    Set rx = CreateObject("VBScript.RegExp")
    rx.Pattern = "^[A-Za-z_][A-Za-z0-9_$#]*(\.[A-Za-z_][A-Za-z0-9_$#]*)*$"
    If Not rx.Test(Trim$(columnName)) Then
        Err.Raise 5, "clsSqlBuilder.SetChunkedInPredicate", "Qualified column name required for {" & key & "}."
    End If
End Sub

Public Function Build() As String
    Dim rx As Object, matches As Object, m As Object
    Dim lastPos As Long, tokenName As String, output As String

    If Len(m_Template) = 0 Then
        Err.Raise 5, "clsSqlBuilder.Build", "No SQL template loaded."
    End If

    Set rx = CreateObject("VBScript.RegExp")
    rx.Pattern = "\{([A-Za-z_][A-Za-z0-9_]*)\}"
    rx.Global = True

    Set matches = rx.Execute(m_Template)
    lastPos = 1
    For Each m In matches
        tokenName = UCase$(CStr(m.SubMatches(0)))
        If Not m_Params.exists(tokenName) Then
            Err.Raise 5, "clsSqlBuilder.Build", "Unresolved SQL placeholder {" & CStr(m.SubMatches(0)) & "} in " & m_TemplateName
        End If
        output = output & Mid$(m_Template, lastPos, (m.FirstIndex + 1) - lastPos) & CStr(m_Params(tokenName))
        lastPos = m.FirstIndex + m.Length + 1
    Next m
    output = output & Mid$(m_Template, lastPos)

    Build = output
End Function

Private Sub ValidateKey(ByVal key As String)
    Dim i As Long, ch As String
    key = Trim$(key)
    If Len(key) = 0 Then Err.Raise 5, "clsSqlBuilder", "Placeholder key cannot be empty."
    ch = Left$(key, 1)
    If Not IsAsciiLetter(ch) And ch <> "_" Then Err.Raise 5, "clsSqlBuilder", "Invalid placeholder key: " & key
    For i = 2 To Len(key)
        ch = Mid$(key, i, 1)
        If Not IsAsciiLetter(ch) And (ch < "0" Or ch > "9") And ch <> "_" Then
            Err.Raise 5, "clsSqlBuilder", "Invalid placeholder key: " & key
        End If
    Next i
End Sub

Private Function IsAsciiLetter(ByVal Value As String) As Boolean
    If Len(Value) <> 1 Then Exit Function
    IsAsciiLetter = ((Value >= "A" And Value <= "Z") Or (Value >= "a" And Value <= "z"))
End Function

Private Function FormatNumberLiteral(ByVal Value As Variant) As String
    Dim text As String, digits As String, i As Long, ch As String
    If VarType(Value) = vbString Then
        text = Trim$(CStr(Value))
        digits = text
        If Left$(digits, 1) = "+" Or Left$(digits, 1) = "-" Then digits = Mid$(digits, 2)
        If Len(digits) > 0 Then
            For i = 1 To Len(digits)
                ch = Mid$(digits, i, 1)
                If ch < "0" Or ch > "9" Then Exit For
            Next i
            If i > Len(digits) Then
                ' Preserve long integer identifiers without a lossy Double conversion.
                FormatNumberLiteral = text
                Exit Function
            End If
        End If
    End If
    If VarType(Value) = vbDecimal Or VarType(Value) = vbCurrency Then
        FormatNumberLiteral = Replace$(CStr(Value), ",", ".")
    Else
        ' Str$ always uses a dot, including on decimal-comma systems.
        FormatNumberLiteral = Trim$(Str$(CDbl(Value)))
    End If
End Function

Private Function ExtractUniqueItems(ByVal items As Variant, ByVal key As String) As Collection
    Dim seen As Object, result As New Collection, rawItem As Variant
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1

    If IsEmpty(items) Or IsNull(items) Then
        Set ExtractUniqueItems = result
        Exit Function
    End If

    If IsObject(items) Then
        If items Is Nothing Then
            Set ExtractUniqueItems = result
            Exit Function
        End If
        If TypeName(items) = "Dictionary" Then
            For Each rawItem In items.Keys
                AddUniqueItem rawItem, seen, result, key
            Next rawItem
        Else
            For Each rawItem In items
                AddUniqueItem rawItem, seen, result, key
            Next rawItem
        End If
    ElseIf IsArray(items) Then
        For Each rawItem In items
            AddUniqueItem rawItem, seen, result, key
        Next rawItem
    Else
        AddUniqueItem items, seen, result, key
    End If

    Set ExtractUniqueItems = result
End Function

Private Sub AddUniqueItem(ByVal rawItem As Variant, ByVal seen As Object, ByVal result As Collection, ByVal key As String)
    Dim textVal As String
    If IsNull(rawItem) Or IsEmpty(rawItem) Then Exit Sub
    textVal = Trim$(CStr(rawItem))
    If Len(textVal) = 0 Then Exit Sub
    If Len(textVal) > 128 Then Err.Raise 5, "clsSqlBuilder", "List item exceeds 128 characters in {" & key & "}."
    If InStr(textVal, vbCr) > 0 Or InStr(textVal, vbLf) > 0 Then
        Err.Raise 5, "clsSqlBuilder", "List item contains invalid line break in {" & key & "}."
    End If
    If Not seen.exists(textVal) Then
        seen.Add textVal, True
        result.Add textVal
    End If
End Sub

Private Function FormatListItem(ByVal textVal As String, ByVal isNumericList As Boolean, ByVal key As String) As String
    If isNumericList Then
        If Not IsNumeric(textVal) Then
            Err.Raise 5, "clsSqlBuilder", "Expected numeric list item for {" & key & "}, got: " & textVal
        End If
        FormatListItem = FormatNumberLiteral(textVal)
    Else
        FormatListItem = "'" & Replace(textVal, "'", "''") & "'"
    End If
End Function
