Attribute VB_Name = "Mod_DataSchemas"
Option Explicit

Public Function NewAnySchema(ByVal datasetName As String, ByVal grain As String) As clsDataSchema
    Dim result As clsDataSchema
    Set result = New clsDataSchema: result.Initialize datasetName, "1", grain
    Set NewAnySchema = result
End Function

Public Function NewPortfolioSchema(ByVal datasetName As String) As clsDataSchema
    Dim result As clsDataSchema, textFields As Variant, item As Variant
    datasetName = UCase$(Trim$(datasetName))
    Select Case datasetName
        Case "PORT_HOLDINGS", "PORT_PERFORMANCE", "PORT_PROFIT", "PORT_ORDERS", "PORT_INDUSTRY"
        Case Else: Err.Raise 5, "NewPortfolioSchema", "Unknown dataset: " & datasetName
    End Select
    Set result = New clsDataSchema: result.Initialize datasetName, "1", "SOURCE_QUERY_ROW"
    textFields = Array("PORT_CODE", "I_CODE", "M_TYPE", "I_NAME", "A_TYPE", "CURRENCY", "ORDER_ID", "SEC_CODE", "SW_SECTOR_CODE", "SW_SECTOR", "SW_GROUP_CODE", "SW_GROUP", "SW_INDUSTRY_CODE", "SW_INDUSTRY", "ROUTE", "SOURCE_TABLE", "ISSUE_NAME", "POSITION", "T_TYPE", "TRADE_ATTR", "ISSUER")
    For Each item In textFields
        If DataSchemaHasField(datasetName, CStr(item)) Then result.AddField CStr(item), "TEXT"
    Next item
    If UCase$(datasetName) <> "PORT_INDUSTRY" Then result.AddField "T_DATE", "DATE", False
    If datasetName = "PORT_INDUSTRY" Then
        result.AddField "START_DATE", "DATE": result.AddField "EXPIRATION_DATE", "DATE"
    End If
    result.Freeze: Set NewPortfolioSchema = result
End Function

Private Function DataSchemaHasField(ByVal datasetName As String, ByVal fieldName As String) As Boolean
    Select Case UCase$(datasetName)
        Case "PORT_HOLDINGS": DataSchemaHasField = (InStr("|PORT_CODE|I_CODE|M_TYPE|I_NAME|A_TYPE|CURRENCY|", "|" & fieldName & "|") > 0)
        Case "PORT_PERFORMANCE": DataSchemaHasField = (InStr("|PORT_CODE|", "|" & fieldName & "|") > 0)
        Case "PORT_PROFIT": DataSchemaHasField = (InStr("|PORT_CODE|I_CODE|M_TYPE|ISSUE_NAME|", "|" & fieldName & "|") > 0)
        Case "PORT_ORDERS": DataSchemaHasField = (InStr("|PORT_CODE|I_CODE|M_TYPE|I_NAME|A_TYPE|T_TYPE|TRADE_ATTR|ISSUER|", "|" & fieldName & "|") > 0)
        Case "PORT_INDUSTRY": DataSchemaHasField = (InStr("|SEC_CODE|SW_SECTOR_CODE|SW_SECTOR|SW_GROUP_CODE|SW_GROUP|SW_INDUSTRY_CODE|SW_INDUSTRY|ROUTE|SOURCE_TABLE|", "|" & fieldName & "|") > 0)
    End Select
End Function

Public Function NewFinancialWideSchema() As clsDataSchema
    Dim result As clsDataSchema
    Set result = New clsDataSchema: result.Initialize "FINANCIAL_WIDE", "1", "COMPANY_CODE+END_DATE+IF_MERGED+IF_ADJUSTED+INFO_PUBL_DATE"
    result.AddField "COMPANY_CODE", "TEXT", False
    result.AddField "END_DATE", "DATE", False
    result.AddField "IF_MERGED", "LONG", False
    result.AddField "IF_ADJUSTED", "LONG", False
    result.AddField "INFO_PUBL_DATE", "DATE", False
    result.Freeze: Set NewFinancialWideSchema = result
End Function

Public Function NewFinancialOverrideSchema() As clsDataSchema
    Dim result As clsDataSchema
    Set result = New clsDataSchema
    result.Initialize "MANUAL_FINANCIALS", "1", "COMPANY_CODE+END_DATE+IF_MERGED+IF_ADJUSTED+INFO_PUBL_DATE+ACCOUNT_CODE"
    result.AddField "COMPANY_CODE", "TEXT", False
    result.AddField "END_DATE", "DATE", False
    result.AddField "IF_MERGED", "LONG", False
    result.AddField "IF_ADJUSTED", "LONG", False
    result.AddField "INFO_PUBL_DATE", "DATE", False
    result.AddField "ACCOUNT_CODE", "TEXT", False
    result.AddField "AMOUNT", "DOUBLE", True
    result.Freeze
    Set NewFinancialOverrideSchema = result
End Function

Public Function NewSecurityMasterSchema() As clsDataSchema
    Dim result As clsDataSchema
    Set result = New clsDataSchema
    result.Initialize "DIM_SECURITY_MASTER", "1", "INNER_CODE"
    result.AddField "INNER_CODE", "LONG", False
    result.AddField "SEC_CODE", "TEXT", False
    result.AddField "MARKET_CODE", "TEXT", False
    result.AddField "SEC_NAME", "TEXT", True
    result.AddField "ASSET_CLASS", "TEXT", False
    result.AddField "SUB_TYPE", "TEXT", False
    result.AddField "COMPANY_CODE", "TEXT", True
    result.AddField "COMPANY_NAME", "TEXT", True
    result.AddField "SOURCE_TABLE", "TEXT", False
    result.AddField "LIST_DATE", "DATE", True
    result.AddField "DELIST_DATE", "DATE", True
    result.AddField "UPDATE_TIME", "DATE", True
    result.Freeze
    Set NewSecurityMasterSchema = result
End Function

Public Function NewStockQuoteSchema() As clsDataSchema
    Dim s As New clsDataSchema
    s.Initialize "STK_DAILY_QUOTE", "1", "INNER_CODE|T_DATE"
    s.AddField "INNER_CODE", "LONG", False
    s.AddField "T_DATE", "DATE", False
    s.AddField "PREV_CLOSE", "DOUBLE", True
    s.AddField "OPEN_PRICE", "DOUBLE", True
    s.AddField "HIGH_PRICE", "DOUBLE", True
    s.AddField "LOW_PRICE", "DOUBLE", True
    s.AddField "CLOSE_PRICE", "DOUBLE", True
    s.AddField "ADJ_FACTOR", "DOUBLE", False
    s.AddField "ADJ_OPEN", "DOUBLE", True
    s.AddField "ADJ_HIGH", "DOUBLE", True
    s.AddField "ADJ_LOW", "DOUBLE", True
    s.AddField "ADJ_CLOSE", "DOUBLE", True
    s.AddField "VOLUME", "DOUBLE", True
    s.AddField "AMOUNT", "DOUBLE", True
    s.AddField "DAILY_RTN", "DOUBLE", True
    s.AddField "TOTAL_MV", "DOUBLE", True
    s.AddField "FLOAT_MV", "DOUBLE", True
    s.Freeze
    Set NewStockQuoteSchema = s
End Function
