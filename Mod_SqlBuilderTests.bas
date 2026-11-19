Attribute VB_Name = "Mod_SqlBuilderTests"
Option Explicit

' Offline regression: renders real templates without opening a database connection.
Public Function RunSqlBuilderOfflineTests() As String
    Dim builder As clsSqlBuilder, config As clsConfig, count As Long
    Dim template As Variant, actual As String, expected As String, sql As String
    Dim codes As Object, items As Collection, i As Long, values() As String
    Dim errorNumber As Long
    On Error GoTo Failed
    Set builder = New clsSqlBuilder
    Set config = New clsConfig

    ' Existing portfolio templates must render exactly like the previous service.
    For Each template In Array("PORT_Holdings.sql", "PORT_Measures.sql", "PORT_PLDetail.sql", "PORT_Orders.sql")
        Call builder.LoadTemplate(CStr(template))
        Call builder.SetInList("PortList", Array("001", "O'Neil"))
        Call builder.SetDate("StartDate", DateSerial(2026, 11, 10))
        Call builder.SetDate("EndDate", DateSerial(2026, 11, 14))
        expected = config.ReadSQLFile(CStr(template))
        expected = Replace(expected, "{PortList}", "'001','O''Neil'")
        expected = Replace(expected, "{StartDate}", "'2026-11-10'")
        expected = Replace(expected, "{EndDate}", "'2026-11-14'")
        Check builder.Build() = expected, count, CStr(template)
    Next template

    ' Financial dates are already quoted in the shared SQL template.
    sql = ChrW(&H8D22) & ChrW(&H62A5) & ".sql"
    Call builder.LoadTemplate(sql)
    Call builder.SetNumber("CompanyCode", "123456789012345678")
    Call builder.SetDate("ReportBegin", DateSerial(2024, 12, 31), False)
    Call builder.SetDate("ReportEnd", DateSerial(2025, 12, 31), False)
    Call builder.SetDate("AsOfDate", DateSerial(2026, 11, 14), False)
    Call builder.SetNumber("IfMerged", 1)
    expected = config.ReadSQLFile(sql)
    expected = Replace(expected, "{CompanyCode}", "123456789012345678")
    expected = Replace(expected, "{ReportBegin}", "2024-12-31")
    expected = Replace(expected, "{ReportEnd}", "2025-12-31")
    expected = Replace(expected, "{AsOfDate}", "2026-11-14")
    expected = Replace(expected, "{IfMerged}", "1")
    Check builder.Build() = expected, count, "Financial template and integer precision"

    ' Access uses month/day/year literals and exclusive next-day upper bounds.
    Call builder.LoadRawSql("[END_DATE]>={BeginDate} AND [END_DATE]<{EndDate} AND [INFO_PUBL_DATE]<{AsOfDate}")
    Call builder.SetAccessDate("BeginDate", DateSerial(2024, 2, 1))
    Call builder.SetAccessDate("EndDate", DateAdd("d", 1, DateSerial(2024, 2, 29)))
    Call builder.SetAccessDate("AsOfDate", DateAdd("d", 1, DateSerial(2026, 12, 31)))
    Check builder.Build() = "[END_DATE]>=#02/01/2024# AND [END_DATE]<#03/01/2024# AND [INFO_PUBL_DATE]<#01/01/2027#", count, "Access date boundaries"

    ' Dictionary keys are the security identifiers, not dictionary values.
    Set codes = CreateObject("Scripting.Dictionary")
    codes.Add "000001", "ignored"
    codes.Add "O'Neil", "ignored"
    Call builder.LoadTemplate("PORT_StockIndustry.sql")
    Call builder.SetChunkedInPredicate("MainBoardPredicate", "b.AStockCode", codes)
    Call builder.SetChunkedInPredicate("StibPredicate", "b.SecuCode", Array())
    expected = config.ReadSQLFile("PORT_StockIndustry.sql")
    expected = Replace(expected, "{MainBoardPredicate}", "b.AStockCode IN ('000001','O''Neil')")
    expected = Replace(expected, "{StibPredicate}", "1 = 0")
    Check builder.Build() = expected, count, "Industry routes and empty list"

    ReDim values(0 To 900)
    For i = 0 To 900
        values(i) = "SEC" & CStr(i)
    Next i
    Call builder.LoadRawSql("WHERE {Codes} AND active = 1")
    Call builder.SetChunkedInPredicate("Codes", "b.SecuCode", values)
    actual = builder.Build()
    Check Left$(actual, 7) = "WHERE (", count, "Chunked OR grouping"
    Check InStr(actual, "'SEC899') OR b.SecuCode IN ('SEC900')) AND active = 1") > 0, count, "901 items split after item 900"
    Check UBound(Split(actual, " IN (")) = 2, count, "Exactly two chunks"
    On Error Resume Next
    Call builder.SetInList("Codes", values)
    errorNumber = Err.number
    Err.Clear
    On Error GoTo Failed
    Check errorNumber = 5, count, "Unchunked oversized list rejected"
    On Error Resume Next
    Call builder.SetChunkedInPredicate("Codes", "b.SecuCode) OR 1=1 --", Array("SEC1"))
    errorNumber = Err.number
    Err.Clear
    On Error GoTo Failed
    Check errorNumber = 5, count, "Chunked predicate rejects untrusted columns"

    Set items = New Collection
    items.Add " 001 ": items.Add "001": items.Add "002"
    Call builder.LoadRawSql("{Codes}")
    Call builder.SetInList("Codes", items)
    Check builder.Build() = "'001','002'", count, "Collection normalization and deduplication"
    Call builder.SetInList("Codes", "003")
    Check builder.Build() = "'003'", count, "Scalar input"
    On Error Resume Next
    Call builder.SetInList("Codes", Array())
    errorNumber = Err.number
    Err.Clear
    On Error GoTo Failed
    Check errorNumber = 5, count, "Empty IN list rejected"

    Call builder.LoadRawSql("{amount}|{Tiny}|{Ids}")
    Call builder.SetNumber("AMOUNT", 1.5)
    Call builder.SetNumber("Tiny", 1E-40)
    Call builder.SetInList("Ids", Array("123456789012345678", "2"), True)
    actual = builder.Build()
    Check Left$(actual, 4) = "1.5|", count, "Locale-independent decimal and case-insensitive keys"
    Check InStr(actual, "1E-40") > 0, count, "Small numbers are not rounded to zero"
    Check Right$(actual, 20) = "123456789012345678,2", count, "Numeric list integer precision"

    Call builder.LoadRawSql("{Text}|{Other}")
    Call builder.SetString("Text", "O'Neil {Other}")
    Call builder.SetNumber("Other", 2)
    Check builder.Build() = "'O''Neil {Other}'|2", count, "Values are escaped and never re-expanded"
    Call builder.LoadRawSql("{Other}")
    On Error Resume Next
    actual = builder.Build()
    errorNumber = Err.number
    Err.Clear
    On Error GoTo Failed
    Check errorNumber = 5, count, "Loading a new template clears prior parameters"
    On Error Resume Next
    Call builder.SetString("Bad-Key", "value")
    errorNumber = Err.number
    Err.Clear
    On Error GoTo Failed
    Check errorNumber = 5, count, "Keys follow placeholder grammar"

    RunSqlBuilderOfflineTests = "PASS: " & CStr(count) & " SQL assertions; no database connections."
    Exit Function
Failed:
    Err.Raise Err.number, "Mod_SqlBuilderTests", "SQL assertion " & CStr(count + 1) & " failed: " & Err.description
End Function

Private Sub Check(ByVal Condition As Boolean, ByRef count As Long, ByVal label As String)
    If Not Condition Then Err.Raise vbObjectError + 2100, "Mod_SqlBuilderTests", label
    count = count + 1
End Sub
