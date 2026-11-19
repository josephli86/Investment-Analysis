VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsResultTableRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Presentation-only adapter; does not calculate or alter result values.
Public Sub Render(ByVal result As clsAnalysisResult, ByVal tableName As String, ByVal sheet As Object)
    Dim data As Variant, rows As Long, columns As Long, r As Long, c As Long, hasText As Boolean
    If result Is Nothing Or sheet Is Nothing Then Err.Raise 5, , "Result and target sheet required."
    data = result.table(tableName).ToTable()
    rows = UBound(data, 1): columns = UBound(data, 2)
    If rows > sheet.rows.count Then Err.Raise 6, , "Too many result rows."
    For r = 1 To rows
        For c = 1 To columns
            If IsNull(data(r, c)) Or IsEmpty(data(r, c)) Then data(r, c) = Empty
        Next c
    Next r
    For c = 1 To columns
        hasText = False
        For r = 2 To rows
            If VarType(data(r, c)) = vbString Then hasText = True: Exit For
        Next r
        If hasText Then sheet.Cells(1, c).Resize(rows, 1).NumberFormat = "@"
    Next c
    With sheet.Cells(1, 1).Resize(rows, columns)
        .Value2 = data
        .rows(1).Font.Bold = True
        .columns.AutoFit
    End With
End Sub
