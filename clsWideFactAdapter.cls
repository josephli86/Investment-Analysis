VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWideFactAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Sub LoadRecordset(ByVal rs As Object, ByVal target As clsFinancialStatement)
    Dim reader As clsRecordsetReader, data As clsDataTable
    If target Is Nothing Then Err.Raise 5, "clsWideFactAdapter", "Target statement required."
    Set reader = New clsRecordsetReader
    Set data = reader.Read(rs)
    LoadData data, target
End Sub

Public Sub LoadData(ByVal data As clsDataTable, ByVal target As clsFinancialStatement)
    Dim c As Long, r As Long, key As String, code As String, name As String
    Dim required As Variant, item As Variant
    Dim company As Variant, period As Variant, merged As Variant, adjusted As Variant, publication As Variant
    Dim Value As Variant
    If data Is Nothing Or target Is Nothing Then Err.Raise 5, "clsWideFactAdapter.LoadData", "Data and target statement are required."
    If Not data.IsLoaded Then Err.Raise 5, "clsWideFactAdapter.LoadData", "Data table is not loaded."
    required = Array("COMPANY_CODE", "END_DATE", "IF_MERGED", "IF_ADJUSTED", "INFO_PUBL_DATE")
    ' Validate metadata even when the result set is empty.
    For Each item In required
        Call data.FieldIndex(CStr(item))
    Next item
    For r = 1 To data.rowCount
        For Each item In required
            If IsNull(data.ValueAt(r, CStr(item))) Or IsEmpty(data.ValueAt(r, CStr(item))) Then Err.Raise 5, "clsWideFactAdapter.LoadData", "NULL report identity: " & CStr(item)
        Next item
        company = data.ValueAt(r, "COMPANY_CODE"): period = data.ValueAt(r, "END_DATE")
        merged = data.ValueAt(r, "IF_MERGED"): adjusted = data.ValueAt(r, "IF_ADJUSTED")
        publication = data.ValueAt(r, "INFO_PUBL_DATE")
        code = CStr(company)
        If code <> target.companyCode Then Err.Raise 5, "clsWideFactAdapter.LoadData", "Unexpected company code: " & code
        key = target.RegisterReport(CStr(period), CLng(merged), CLng(adjusted), CStr(publication))
        For c = 1 To data.fieldCount
            name = data.fieldName(c)
            Select Case name
                Case "COMPANY_CODE", "END_DATE", "IF_MERGED", "IF_ADJUSTED", "INFO_PUBL_DATE"
                Case Else
                    Value = data.ValueAtColumn(r, c)
                    target.AddReportFact key, name, Value
            End Select
        Next c
    Next r
End Sub
