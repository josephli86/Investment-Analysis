VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFixedIncomeModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const BP_SIZE As Double = 0.0001

' Cash flow PERIODS are year fractions from settlement; CASHFLOW uses currency units.
Public Function price(ByVal schedule As clsDataTable, ByVal annualRate As Double) As Double
    Dim i As Long, period As Variant, amount As Variant, total As Double, denominator As Double
    ValidateSchedule schedule
    If annualRate <= -1# Then Err.Raise 5, "clsFixedIncomeModel.Price", "Annual rate must exceed -100%."
    For i = 1 To schedule.rowCount
        period = schedule.ValueAt(i, "PERIODS"): amount = schedule.ValueAt(i, "CASHFLOW")
        denominator = (1# + annualRate) ^ CDbl(period)
        total = total + CDbl(amount) / denominator
    Next i
    price = total
End Function

Public Function dv01(ByVal schedule As clsDataTable, ByVal annualRate As Double) As Double
    dv01 = (price(schedule, annualRate - BP_SIZE) - price(schedule, annualRate + BP_SIZE)) / 2#
End Function

Public Function RunRateScenario(ByVal request As clsAnalysisRequest, ByVal schedule As clsDataTable, ByVal quantity As Double, ByVal shockBasisPoints As Double) As clsAnalysisResult
    Dim result As clsAnalysisResult, table As clsDataTable, values() As Variant
    Dim rate As Double, basePrice As Double, shockedPrice As Double, diagnostic As String
    If request Is Nothing Then Err.Raise 5, , "Analysis request required."
    If request.taskType <> "BOND_RATE_SCENARIO" Then Err.Raise 5, , "Unsupported task type."
    rate = CDbl(request.Parameter("ANNUAL_RATE"))
    basePrice = price(schedule, rate)
    shockedPrice = price(schedule, rate + shockBasisPoints / 10000#)
    ReDim values(1 To 2, 1 To 6)
    values(1, 1) = "VALUATION_DATE": values(1, 2) = "BASE_PRICE": values(1, 3) = "SHOCKED_PRICE"
    values(1, 4) = "QUANTITY": values(1, 5) = "PNL": values(1, 6) = "DV01_PER_BOND"
    values(2, 1) = request.context.valuationDate
    values(2, 2) = basePrice: values(2, 3) = shockedPrice
    values(2, 4) = quantity: values(2, 5) = (shockedPrice - basePrice) * quantity
    values(2, 6) = dv01(schedule, rate)
    Set table = New clsDataTable: table.Load values
    Set result = New clsAnalysisResult: result.Initialize request
    result.AddTable "BOND_RATE_SCENARIO", table
    If Abs(shockBasisPoints) > 500# Then
        diagnostic = "Scenario shock exceeds the 500 bp sample guardrail; review assumptions."
        result.AddDiagnostic "LARGE_RATE_SHOCK", "BOND", "VALUATION", diagnostic
    End If
    result.Freeze: Set RunRateScenario = result
End Function

Private Sub ValidateSchedule(ByVal schedule As clsDataTable)
    Dim i As Long, period As Variant, amount As Variant
    If schedule Is Nothing Then Err.Raise 5, "clsFixedIncomeModel", "Cash-flow schedule required."
    Call schedule.FieldIndex("PERIODS"): Call schedule.FieldIndex("CASHFLOW")
    If schedule.rowCount = 0 Then Err.Raise 5, "clsFixedIncomeModel", "Cash-flow schedule is empty."
    For i = 1 To schedule.rowCount
        period = schedule.ValueAt(i, "PERIODS"): amount = schedule.ValueAt(i, "CASHFLOW")
        If IsNull(period) Or IsEmpty(period) Or IsNull(amount) Or IsEmpty(amount) Then Err.Raise 5, "clsFixedIncomeModel", "Cash-flow values cannot be missing."
        If Not IsNumeric(period) Or Not IsNumeric(amount) Then Err.Raise 13, "clsFixedIncomeModel", "Cash-flow values must be numeric."
        If CDbl(period) <= 0# Then Err.Raise 5, "clsFixedIncomeModel", "Periods must be positive year fractions."
    Next i
End Sub
