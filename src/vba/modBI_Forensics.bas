Attribute VB_Name = "modBI_Forensics"
Option Explicit

'/**
' * Description: Lists Altman Z-Score variables in expected input order.
' * Parameters: None.
' * Returns: One-column Variant array of variable labels.
' */
Public Function BI_AltmanZScoreVariables() As Variant
    Dim result(1 To 7, 1 To 1) As Variant
    result(1, 1) = "Working Capital"
    result(2, 1) = "Total Assets"
    result(3, 1) = "Retained Earnings"
    result(4, 1) = "Earnings Before Interest and Taxes"
    result(5, 1) = "Market Value"
    result(6, 1) = "Total Liabilities"
    result(7, 1) = "Sales"
    BI_AltmanZScoreVariables = result
End Function

'/**
' * Description: Calculates Altman Z-Score using the same coefficients as the Lambda source.
' * Parameters: Data - Seven-item range in BI_AltmanZScoreVariables order.
' * Returns: Two-row Variant array containing score and verdict.
' */
Public Function BI_AltmanZScore(ByVal Data As Variant) As Variant
    On Error GoTo ErrHandler
    Dim v As Variant
    Dim nwc As Double, ta As Double, re As Double, ebit As Double
    Dim mv As Double, tl As Double, sales As Double, z As Double
    Dim result(1 To 2, 1 To 2) As Variant

    v = BI_ToColumnVector(Data)
    nwc = CDbl(v(1, 1))
    ta = CDbl(v(2, 1))
    re = CDbl(v(3, 1))
    ebit = CDbl(v(4, 1))
    mv = CDbl(v(5, 1))
    tl = CDbl(v(6, 1))
    sales = CDbl(v(7, 1))

    z = 1.2 * (nwc / ta) + 1.4 * (re / ta) + 3.3 * (ebit / ta) + 0.6 * (mv / tl) + (sales / ta)
    result(1, 1) = "Altman Z Score": result(1, 2) = z
    result(2, 1) = "Verdict"
    If z >= 3 Then
        result(2, 2) = ">= 3 - Safe Zone - Low likelyhood of Bankruptcy"
    ElseIf z < 1.81 Then
        result(2, 2) = "<1.81 - Distress Zone - High Likelihood of Bankruptcy"
    Else
        result(2, 2) = "Middle - Grey Zone - Moderate Risk of Bankruptcy"
    End If
    BI_AltmanZScore = result
    Exit Function
ErrHandler:
    BI_AltmanZScore = CVErr(xlErrValue)
End Function

'/**
' * Description: Calculates Ohlson O-Score from the source variable order.
' * Parameters: Data - Nine-item range in BI_OhlsonsOScoreVariables order.
' * Returns: Two-row Variant array containing score and probability.
' */
Public Function BI_OhlsonsOScore(ByVal Data As Variant) As Variant
    On Error GoTo ErrHandler
    Dim v As Variant, result(1 To 2, 1 To 2) As Variant
    Dim ta As Double, gnp As Double, tl As Double, wc As Double, cl As Double, ca As Double
    Dim lyni As Double, cyni As Double, ffo As Double
    Dim score As Double

    v = BI_ToColumnVector(Data)
    ta = CDbl(v(1, 1)): gnp = CDbl(v(2, 1)): tl = CDbl(v(3, 1))
    wc = CDbl(v(4, 1)): cl = CDbl(v(5, 1)): ca = CDbl(v(6, 1))
    lyni = CDbl(v(7, 1)): cyni = CDbl(v(8, 1)): ffo = CDbl(v(9, 1))

    score = -1.32 - 0.407 * Log(ta / gnp) + 6.03 * (wc / ta) + 0.0757 * (cl / ca) _
        - 1.72 * IIf(tl > ta, 1, 0) - 2.37 * (cyni / ta) - 1.83 * (ffo / tl) _
        + 0.285 * IIf(lyni < 0 And cyni < 0, 1, 0) _
        - 0.521 * ((cyni - lyni) / (Abs(lyni) + Abs(cyni)))
    result(1, 1) = "Ohlson O Score": result(1, 2) = score
    result(2, 1) = "Verdict"
    result(2, 2) = IIf(score > 0.5, "The firm will default within two years.", "The firm is not likely to default within two years.")
    BI_OhlsonsOScore = result
    Exit Function
ErrHandler:
    BI_OhlsonsOScore = CVErr(xlErrValue)
End Function

'/**
' * Description: Calculates Beneish M-Score using the source Lambda ratios and coefficients.
' * Parameters: LastYear - Twelve-item range in BI_BeneishMScoreVariables order; CurrentYear - same order.
' * Returns: Variant array containing ratios, 8-variable score, 5-variable score, and verdicts.
' */
Public Function BI_BeneishMScore(ByVal LastYear As Variant, ByVal CurrentYear As Variant) As Variant
    On Error GoTo ErrHandler
    Dim ly As Variant, cy As Variant
    Dim dsri As Double, gmi As Double, aqi As Double, sgi As Double
    Dim depi As Double, sgai As Double, lvgi As Double, tata As Double
    Dim lyolta As Double, cyolta As Double, m8 As Double, m5 As Double
    Dim result(1 To 12, 1 To 2) As Variant

    ly = BI_ToColumnVector(LastYear)
    cy = BI_ToColumnVector(CurrentYear)
    lyolta = CDbl(ly(9, 1)) - (CDbl(ly(7, 1)) + CDbl(ly(5, 1)))
    cyolta = CDbl(cy(9, 1)) - (CDbl(cy(7, 1)) + CDbl(cy(5, 1)))

    dsri = (CDbl(cy(8, 1)) / CDbl(cy(1, 1))) / (CDbl(ly(8, 1)) / CDbl(ly(1, 1)))
    gmi = ((CDbl(ly(1, 1)) - CDbl(ly(2, 1))) / CDbl(ly(1, 1))) / ((CDbl(cy(1, 1)) - CDbl(cy(2, 1))) / CDbl(cy(1, 1)))
    aqi = (cyolta / CDbl(cy(9, 1))) / (lyolta / CDbl(ly(9, 1)))
    sgi = CDbl(cy(1, 1)) / CDbl(ly(1, 1))
    depi = (CDbl(ly(6, 1)) / (CDbl(ly(6, 1)) + CDbl(ly(5, 1)))) / (CDbl(cy(6, 1)) / (CDbl(cy(6, 1)) + CDbl(cy(5, 1))))
    sgai = (CDbl(cy(3, 1)) / CDbl(cy(1, 1))) / (CDbl(ly(3, 1)) / CDbl(ly(1, 1)))
    lvgi = ((CDbl(cy(10, 1)) + CDbl(cy(11, 1))) / CDbl(cy(9, 1))) / ((CDbl(ly(10, 1)) + CDbl(ly(11, 1))) / CDbl(ly(9, 1)))
    tata = (CDbl(cy(4, 1)) - CDbl(cy(12, 1))) / CDbl(cy(9, 1))
    m8 = -4.84 + 0.92 * dsri + 0.528 * gmi + 0.404 * aqi + 0.892 * sgi + 0.115 * depi - 0.172 * sgai - 0.327 * lvgi + 4.679 * tata
    m5 = -6.065 + 0.823 * dsri + 0.906 * gmi + 0.593 * aqi + 0.717 * sgi + 0.107 * depi

    result(1, 1) = "DSRI": result(1, 2) = dsri
    result(2, 1) = "GMI": result(2, 2) = gmi
    result(3, 1) = "AQI": result(3, 2) = aqi
    result(4, 1) = "SGI": result(4, 2) = sgi
    result(5, 1) = "DEPI": result(5, 2) = depi
    result(6, 1) = "SGAI": result(6, 2) = sgai
    result(7, 1) = "LVGI": result(7, 2) = lvgi
    result(8, 1) = "TATA": result(8, 2) = tata
    result(9, 1) = "Beneish M Score 8 Variables": result(9, 2) = m8
    result(10, 1) = "Verdict": result(10, 2) = IIf(m8 < -2.22, "Since M Score (8 Variables) is less than -2.22, the company is not likely to have manipulated its earnings", "Since the M Score is more than -2.22, the company is likely to have manipulated its earnings")
    result(11, 1) = "Beneish M Score 5 Variables": result(11, 2) = m5
    result(12, 1) = "Verdict": result(12, 2) = IIf(m5 < -2.22, "Since M Score (5 Variables) is less than -2.22, the company is not likely to have manipulated its earnings", "Since the M Score is more than -2.22, the company is likely to have manipulated its earnings")
    BI_BeneishMScore = result
    Exit Function
ErrHandler:
    BI_BeneishMScore = CVErr(xlErrValue)
End Function

'/**
' * Description: Calculates Piotroski F-Score using the source Lambda tests.
' * Parameters: LastYearData - Ten-item range in BI_PiotroskiFScoreVariables order; CurrentYearData - same order.
' * Returns: Variant array containing component flags, score, and verdict.
' */
Public Function BI_PiotroskiFScore(ByVal LastYearData As Variant, ByVal CurrentYearData As Variant) As Variant
    On Error GoTo ErrHandler
    Dim ly As Variant, cy As Variant
    Dim score As Long
    Dim roa As Long, cfo As Long, roaChange As Long, cashFlowRoa As Long
    Dim leverage As Long, cr As Long, eq As Long, gpr As Long, assetTo As Long
    Dim result(1 To 11, 1 To 2) As Variant

    ly = BI_ToColumnVector(LastYearData)
    cy = BI_ToColumnVector(CurrentYearData)
    roa = IIf(CDbl(cy(1, 1)) > 0, 1, 0)
    cfo = IIf(CDbl(cy(4, 1)) > CDbl(ly(4, 1)), 1, 0)
    roaChange = IIf((CDbl(cy(1, 1)) / CDbl(cy(3, 1))) > (CDbl(ly(1, 1)) / CDbl(ly(3, 1))), 1, 0)
    cashFlowRoa = IIf(CDbl(cy(4, 1)) > CDbl(cy(1, 1)), 1, 0)
    leverage = IIf((CDbl(cy(5, 1)) / ((CDbl(cy(2, 1)) + CDbl(cy(3, 1))) / 2)) < (CDbl(ly(5, 1)) / ((CDbl(ly(2, 1)) + CDbl(ly(3, 1))) / 2)), 1, 0)
    cr = IIf(CDbl(cy(6, 1)) / CDbl(cy(7, 1)) > CDbl(ly(6, 1)) / CDbl(ly(7, 1)), 1, 0)
    eq = IIf(CDbl(cy(8, 1)) > CDbl(ly(8, 1)), 0, 1)
    gpr = IIf(CDbl(cy(10, 1)) > CDbl(ly(10, 1)), 1, 0)
    assetTo = IIf((CDbl(cy(9, 1)) / CDbl(cy(2, 1))) > (CDbl(ly(9, 1)) / CDbl(ly(2, 1))), 1, 0)
    score = roa + cfo + roaChange + cashFlowRoa + leverage + cr + eq + gpr + assetTo

    result(1, 1) = "Positive Net Income": result(1, 2) = roa
    result(2, 1) = "Positive Cash Flow": result(2, 2) = cfo
    result(3, 1) = "Change in Return on Assets": result(3, 2) = roaChange
    result(4, 1) = "Cash Flow over Net Income": result(4, 2) = cashFlowRoa
    result(5, 1) = "Leverage - Long term debt over Average Assets": result(5, 2) = leverage
    result(6, 1) = "Liquidity - Current Ratio": result(6, 2) = cr
    result(7, 1) = "Total Equity": result(7, 2) = eq
    result(8, 1) = "GP Ratio": result(8, 2) = gpr
    result(9, 1) = "Asset Turnover": result(9, 2) = assetTo
    result(10, 1) = "Total Score": result(10, 2) = score
    result(11, 1) = "Verdict": result(11, 2) = score & " Out of 9 - " & IIf(score >= 8, "Strong!", IIf(score <= 2, "Weak", "Average"))
    BI_PiotroskiFScore = result
    Exit Function
ErrHandler:
    BI_PiotroskiFScore = CVErr(xlErrValue)
End Function

Public Function BI_OhlsonsOScoreVariables() As Variant
    Dim result(1 To 9, 1 To 1) As Variant
    result(1, 1) = "Total Assets"
    result(2, 1) = "Gross National Product Price Level Index"
    result(3, 1) = "Total Liabilities"
    result(4, 1) = "Working Capital"
    result(5, 1) = "Current Liabilities"
    result(6, 1) = "Current Assets"
    result(7, 1) = "LastYear Net Income"
    result(8, 1) = "Current Year Net Income"
    result(9, 1) = "Funds From Operations"
    BI_OhlsonsOScoreVariables = result
End Function

Public Function BI_BeneishMScoreVariables() As Variant
    Dim result(1 To 12, 1 To 1) As Variant
    result(1, 1) = "Net Sales"
    result(2, 1) = "Cost of Goods Sold"
    result(3, 1) = "Selling, General and Admin Expenses"
    result(4, 1) = "Net Income"
    result(5, 1) = "Property, Plant and Equipment"
    result(6, 1) = "Depreciation"
    result(7, 1) = "Current Assets"
    result(8, 1) = "Net Receivables"
    result(9, 1) = "Total Assets"
    result(10, 1) = "Long-tem Debt"
    result(11, 1) = "Current Liabilities"
    result(12, 1) = "Cash Flow from Operations"
    BI_BeneishMScoreVariables = result
End Function

Public Function BI_PiotroskiFScoreVariables() As Variant
    Dim result(1 To 10, 1 To 1) As Variant
    result(1, 1) = "Net Income"
    result(2, 1) = "Opening Total Assets"
    result(3, 1) = "Closing Total Assets"
    result(4, 1) = "Cashflow from Operations"
    result(5, 1) = "Long Term Debt"
    result(6, 1) = "Current Assets"
    result(7, 1) = "Current Liabilities"
    result(8, 1) = "Common Equity"
    result(9, 1) = "Net Sales"
    result(10, 1) = "GP Ratio"
    BI_PiotroskiFScoreVariables = result
End Function

Public Sub BI_OutputScoreInputTemplate(ByVal modelName As String, Optional ByVal target As Range)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Dim variables As Variant
    Dim title As String
    Dim r As Long
    Dim hasTwoYears As Boolean

    If target Is Nothing Then Set target = ActiveCell
    Set ws = target.Worksheet

    Select Case UCase$(modelName)
        Case "ALTMAN"
            variables = BI_AltmanZScoreVariables()
            title = "Altman Z-Score Input Template"
        Case "BENEISH"
            variables = BI_BeneishMScoreVariables()
            title = "Beneish M-Score Input Template"
            hasTwoYears = True
        Case "OHLSON"
            variables = BI_OhlsonsOScoreVariables()
            title = "Ohlson O-Score Input Template"
        Case "PIOTROSKI"
            variables = BI_PiotroskiFScoreVariables()
            title = "Piotroski F-Score Input Template"
            hasTwoYears = True
        Case Else
            Err.Raise vbObjectError + 520, "BI_OutputScoreInputTemplate", "Unknown score model."
    End Select

    target.Value = title
    target.Font.Bold = True
    target.Offset(1, 0).Value = "Variable"
    target.Offset(1, 1).Value = IIf(hasTwoYears, "Last Year", "Value")
    If hasTwoYears Then target.Offset(1, 2).Value = "Current Year"
    target.Offset(1, 0).Resize(1, IIf(hasTwoYears, 3, 2)).Font.Bold = True

    For r = 1 To UBound(variables, 1)
        target.Offset(r + 1, 0).Value = variables(r, 1)
        target.Offset(r + 1, 1).Value = vbNullString
        If hasTwoYears Then target.Offset(r + 1, 2).Value = vbNullString
    Next r

    target.CurrentRegion.Columns.AutoFit
    Exit Sub
ErrHandler:
    MsgBox "Could not create score input template: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_ScoreInputTemplate(control As IRibbonControl)
    On Error GoTo ErrHandler
    Dim modelName As String
    modelName = Application.InputBox("Enter score model: Altman, Beneish, Ohlson, or Piotroski", "BeIndian Score Template", "Altman", Type:=2)
    If Len(modelName) = 0 Then Exit Sub
    BI_OutputScoreInputTemplate modelName, ActiveCell
    Exit Sub
ErrHandler:
    MsgBox "Score template failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_AltmanTemplate(control As IRibbonControl)
    On Error GoTo ErrHandler
    BI_OutputScoreInputTemplate "Altman", ActiveCell
    Exit Sub
ErrHandler:
    MsgBox "Altman template failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_BeneishTemplate(control As IRibbonControl)
    On Error GoTo ErrHandler
    BI_OutputScoreInputTemplate "Beneish", ActiveCell
    Exit Sub
ErrHandler:
    MsgBox "Beneish template failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_OhlsonTemplate(control As IRibbonControl)
    On Error GoTo ErrHandler
    BI_OutputScoreInputTemplate "Ohlson", ActiveCell
    Exit Sub
ErrHandler:
    MsgBox "Ohlson template failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_PiotroskiTemplate(control As IRibbonControl)
    On Error GoTo ErrHandler
    BI_OutputScoreInputTemplate "Piotroski", ActiveCell
    Exit Sub
ErrHandler:
    MsgBox "Piotroski template failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_BeneishMScore(control As IRibbonControl)
    On Error GoTo ErrHandler
    frmBI_Score.Tag = "Beneish"
    frmBI_Score.Show
    Exit Sub
ErrHandler:
    MsgBox "Beneish M-Score failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_OhlsonOScore(control As IRibbonControl)
    On Error GoTo ErrHandler
    frmBI_Score.Tag = "Ohlson"
    frmBI_Score.Show
    Exit Sub
ErrHandler:
    MsgBox "Ohlson O-Score failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_PiotroskiFScore(control As IRibbonControl)
    On Error GoTo ErrHandler
    frmBI_Score.Tag = "Piotroski"
    frmBI_Score.Show
    Exit Sub
ErrHandler:
    MsgBox "Piotroski F-Score failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_AltmanZScore(control As IRibbonControl)
    On Error GoTo ErrHandler
    frmBI_Score.Tag = "Altman"
    frmBI_Score.Show
    Exit Sub
ErrHandler:
    MsgBox "Altman Z-Score tool failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_RelativeSizeFactor(control As IRibbonControl)
    On Error GoTo ErrHandler
    Dim categoryRange As Range
    Dim valueRange As Range
    Dim firstRank As Variant
    Dim secondRank As Variant

    Set categoryRange = Application.InputBox("Select the category column.", "BeIndian Relative Size Factor", Type:=8)
    If categoryRange Is Nothing Then Exit Sub

    Set valueRange = Application.InputBox("Select the numeric value column.", "BeIndian Relative Size Factor", Type:=8)
    If valueRange Is Nothing Then Exit Sub

    firstRank = Application.InputBox("Enter first rank to compare.", "BeIndian Relative Size Factor", 1, Type:=1)
    If VarType(firstRank) = vbBoolean Then Exit Sub

    secondRank = Application.InputBox("Enter second rank to compare.", "BeIndian Relative Size Factor", 2, Type:=1)
    If VarType(secondRank) = vbBoolean Then Exit Sub

    BI_OutputToActiveCell BI_RelativeSizeFactor(categoryRange, valueRange, CLng(firstRank), CLng(secondRank))
    Exit Sub
ErrHandler:
    MsgBox "Relative Size Factor failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub
