VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRecordsetReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Consumes from the current cursor position. Caller owns/rewinds/closes rs.
Public Function Read(ByVal rs As Object) As clsDataTable
    Dim raw As Variant, t() As Variant, r As Long, c As Long, n As Long, result As clsDataTable
    If rs Is Nothing Then Err.Raise 5, , "Recordset required."
    If rs.State <> 1 Then Err.Raise 5, , "Open recordset required."
    If rs.fields.count = 0 Then Err.Raise 5, , "Recordset has no columns."
    If Not rs.EOF Then
        raw = rs.GetRows(): n = UBound(raw, 2) + 1
    End If
    ReDim t(1 To n + 1, 1 To rs.fields.count)
    For c = 1 To rs.fields.count
        t(1, c) = rs.fields(c - 1).name
        For r = 1 To n: t(r + 1, c) = raw(c - 1, r - 1): Next r
    Next c
    Set result = New clsDataTable: result.Load t
    Set Read = result
End Function
