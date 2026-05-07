Attribute VB_Name = "modBI_Financial"
Option Explicit

'/**
' * Description: Returns the payback period from a series of cash flows.
' * Parameters: CashFlows - Column containing initial outflow followed by inflows.
' * Returns: Payback period as fractional period, or #N/A if not recovered.
' */
Public Function BI_PaybackPeriod(ByVal CashFlows As Variant) As Variant
    On Error GoTo ErrHandler
    Dim flows() As Double
    Dim i As Long, cumulative As Double, previous As Double
    flows = BI_ToDoubleVector(CashFlows, False)
    cumulative = flows(LBound(flows))
    For i = LBound(flows) + 1 To UBound(flows)
        previous = cumulative
        cumulative = cumulative + flows(i)
        If cumulative >= 0 Then
            BI_PaybackPeriod = (i - LBound(flows)) - 1 + Abs(previous) / flows(i)
            Exit Function
        End If
    Next i
    BI_PaybackPeriod = CVErr(xlErrNA)
    Exit Function
ErrHandler:
    BI_PaybackPeriod = CVErr(xlErrValue)
End Function

'/**
' * Description: Returns the discounted payback period for a series of cash flows.
' * Parameters: CashFlows - Column of cash flows; DiscountRate - Periodic discount rate.
' * Returns: Discounted payback period as fractional period, or #N/A if not recovered.
' */
Public Function BI_DiscountedPayback(ByVal CashFlows As Variant, ByVal DiscountRate As Double) As Variant
    On Error GoTo ErrHandler
    Dim flows() As Double, discounted As Double
    Dim i As Long, cumulative As Double, previous As Double
    flows = BI_ToDoubleVector(CashFlows, False)
    cumulative = flows(LBound(flows))
    For i = LBound(flows) + 1 To UBound(flows)
        previous = cumulative
        discounted = flows(i) / ((1 + DiscountRate) ^ (i - LBound(flows)))
        cumulative = cumulative + discounted
        If cumulative >= 0 Then
            BI_DiscountedPayback = (i - LBound(flows)) - 1 + Abs(previous) / discounted
            Exit Function
        End If
    Next i
    BI_DiscountedPayback = CVErr(xlErrNA)
    Exit Function
ErrHandler:
    BI_DiscountedPayback = CVErr(xlErrValue)
End Function

'/**
' * Description: Returns an EMI amortization schedule.
' * Parameters: Principal - Loan amount; Period - Number of periods; InterestRate - Periodic interest rate.
' * Returns: Variant array containing period, EMI, interest, principal and closing balance.
' */
Public Function BI_EMISchedule(ByVal Principal As Double, ByVal Period As Long, ByVal InterestRate As Double) As Variant
    On Error GoTo ErrHandler
    Dim result() As Variant
    Dim emi As Double, interest As Double, principalPart As Double, balance As Double
    Dim i As Long
    ReDim result(1 To Period + 1, 1 To 5)
    result(1, 1) = "Period": result(1, 2) = "EMI": result(1, 3) = "Interest": result(1, 4) = "Principal": result(1, 5) = "Balance"
    emi = -Application.WorksheetFunction.Pmt(InterestRate, Period, Principal)
    balance = Principal
    For i = 1 To Period
        interest = balance * InterestRate
        principalPart = emi - interest
        balance = balance - principalPart
        If Abs(balance) < 0.000001 Then balance = 0
        result(i + 1, 1) = i
        result(i + 1, 2) = emi
        result(i + 1, 3) = interest
        result(i + 1, 4) = principalPart
        result(i + 1, 5) = balance
    Next i
    BI_EMISchedule = result
    Exit Function
ErrHandler:
    BI_EMISchedule = CVErr(xlErrValue)
End Function

'/**
' * Description: Returns straight-line depreciation schedule.
' * Parameters: Cost - Asset cost; Life - Number of periods; Scrap - Residual value.
' * Returns: Variant schedule.
' */
Public Function BI_SLNSchedule(ByVal Cost As Double, ByVal Life As Long, ByVal Scrap As Double) As Variant
    Dim result() As Variant
    Dim dep As Double, bookValue As Double, i As Long
    ReDim result(1 To Life + 1, 1 To 4)
    result(1, 1) = "Period": result(1, 2) = "Opening WDV": result(1, 3) = "Depreciation": result(1, 4) = "Closing WDV"
    dep = (Cost - Scrap) / Life
    bookValue = Cost
    For i = 1 To Life
        result(i + 1, 1) = i
        result(i + 1, 2) = bookValue
        result(i + 1, 3) = dep
        bookValue = bookValue - dep
        result(i + 1, 4) = bookValue
    Next i
    BI_SLNSchedule = result
End Function

'/**
' * Description: Returns written-down-value depreciation schedule.
' * Parameters: Cost - Asset cost; Life - Number of periods; Scrap - Residual value.
' * Returns: Variant schedule.
' */
Public Function BI_WDVSchedule(ByVal Cost As Double, ByVal Life As Long, ByVal Scrap As Double) As Variant
    Dim result() As Variant
    Dim rate As Double, dep As Double, bookValue As Double, i As Long
    ReDim result(1 To Life + 1, 1 To 4)
    result(1, 1) = "Period": result(1, 2) = "Opening WDV": result(1, 3) = "Depreciation": result(1, 4) = "Closing WDV"
    rate = 1 - ((Scrap / Cost) ^ (1 / Life))
    bookValue = Cost
    For i = 1 To Life
        result(i + 1, 1) = i
        result(i + 1, 2) = bookValue
        dep = bookValue * rate
        If i = Life Then dep = bookValue - Scrap
        result(i + 1, 3) = dep
        bookValue = bookValue - dep
        result(i + 1, 4) = bookValue
    Next i
    BI_WDVSchedule = result
End Function

Public Sub BI_Tool_EMISchedule(control As IRibbonControl)
    On Error GoTo ErrHandler
    frmBI_EMISchedule.Show
    Exit Sub
ErrHandler:
    MsgBox "EMI Schedule failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_PaybackPeriod(control As IRibbonControl)
    On Error GoTo ErrHandler
    frmBI_FinancialTool.Tag = "Payback"
    frmBI_FinancialTool.Show
    Exit Sub
ErrHandler:
    MsgBox "Payback tool failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_DiscountedPayback(control As IRibbonControl)
    On Error GoTo ErrHandler
    frmBI_FinancialTool.Tag = "DiscountedPayback"
    frmBI_FinancialTool.Show
    Exit Sub
ErrHandler:
    MsgBox "Discounted Payback tool failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_SLNSchedule(control As IRibbonControl)
    On Error GoTo ErrHandler
    frmBI_FinancialTool.Tag = "SLN"
    frmBI_FinancialTool.Show
    Exit Sub
ErrHandler:
    MsgBox "SLN schedule tool failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_WDVSchedule(control As IRibbonControl)
    On Error GoTo ErrHandler
    frmBI_FinancialTool.Tag = "WDV"
    frmBI_FinancialTool.Show
    Exit Sub
ErrHandler:
    MsgBox "WDV schedule tool failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub
