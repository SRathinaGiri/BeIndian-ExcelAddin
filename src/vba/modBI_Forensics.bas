Attribute VB_Name = "modBI_Forensics"
Option Explicit

Private Function BI_F_SourceTo2D(ByVal source As Variant) As Variant
    Dim data As Variant, result() As Variant
    Dim r As Long, c As Long

    If TypeName(source) = "Range" Then data = source.Value2 Else data = source
    If IsArray(data) Then
        On Error GoTo OneDimensional
        ReDim result(1 To UBound(data, 1) - LBound(data, 1) + 1, 1 To UBound(data, 2) - LBound(data, 2) + 1)
        For r = 1 To UBound(result, 1)
            For c = 1 To UBound(result, 2)
                result(r, c) = data(LBound(data, 1) + r - 1, LBound(data, 2) + c - 1)
            Next c
        Next r
        BI_F_SourceTo2D = result
        Exit Function
OneDimensional:
        Err.Clear
        On Error GoTo 0
        ReDim result(1 To UBound(data) - LBound(data) + 1, 1 To 1)
        For r = LBound(data) To UBound(data)
            result(r - LBound(data) + 1, 1) = data(r)
        Next r
    Else
        ReDim result(1 To 1, 1 To 1)
        result(1, 1) = data
    End If
    BI_F_SourceTo2D = result
End Function

Private Function BI_F_ToLongArray(ByVal source As Variant) As Long()
    Dim data As Variant, cols() As Long
    Dim r As Long, c As Long, n As Long
    data = BI_F_SourceTo2D(source)
    ReDim cols(1 To UBound(data, 1) * UBound(data, 2))
    For r = 1 To UBound(data, 1)
        For c = 1 To UBound(data, 2)
            n = n + 1
            cols(n) = CLng(data(r, c))
        Next c
    Next r
    ReDim Preserve cols(1 To n)
    BI_F_ToLongArray = cols
End Function

Private Function BI_F_KeyFromColumns(ByRef data As Variant, ByVal rowIndex As Long, ByRef cols() As Long) As String
    Dim i As Long
    For i = LBound(cols) To UBound(cols)
        BI_F_KeyFromColumns = BI_F_KeyFromColumns & "<^>" & UCase$(CStr(data(rowIndex, cols(i))))
    Next i
End Function

Private Function BI_F_ColumnVector(ByVal source As Variant) As Variant
    BI_F_ColumnVector = BI_ToColumnVector(source)
End Function

Private Sub BI_F_Sort2DByColumn(ByRef data As Variant, ByVal sortColumn As Long, ByVal descending As Boolean)
    Dim r As Long, j As Long, c As Long, cmp As Long, temp As Variant
    For r = 1 To UBound(data, 1) - 1
        For j = r + 1 To UBound(data, 1)
            If IsNumeric(data(j, sortColumn)) And IsNumeric(data(r, sortColumn)) Then
                cmp = Sgn(CDbl(data(j, sortColumn)) - CDbl(data(r, sortColumn)))
            Else
                cmp = StrComp(CStr(data(j, sortColumn)), CStr(data(r, sortColumn)), vbTextCompare)
            End If
            If (descending And cmp > 0) Or ((Not descending) And cmp < 0) Then
                For c = 1 To UBound(data, 2)
                    temp = data(r, c): data(r, c) = data(j, c): data(j, c) = temp
                Next c
            End If
        Next j
    Next r
End Sub

Private Function BI_F_BinomDist(ByVal x As Long, ByVal n As Long, ByVal p As Double) As Double
    BI_F_BinomDist = Application.WorksheetFunction.Binom_Dist(x, n, p, True)
End Function

Private Function BI_F_UpperLimit(ByVal risk As Double, ByVal k As Long, ByVal sampleSize As Long) As Double
    Dim p As Double, best As Double, cumulative As Double
    For p = 1# To 0.001 Step -0.001
        cumulative = BI_F_BinomDist(k, sampleSize, p)
        If cumulative <= risk Then best = p
    Next p
    BI_F_UpperLimit = best
End Function

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

'/**
' * Description: Calculates Gestalt Element Link Test between two related columns.
' * Parameters: AnalyzeData - Primary category column; RelatedData - Related item column.
' * Returns: Table with unique pairs, pair count, total primary count, and ratio.
' */
Public Function BI_GEL1Test(ByVal AnalyzeData As Variant, ByVal RelatedData As Variant) As Variant
    On Error GoTo ErrHandler
    Dim a As Variant, b As Variant, pairCounts As Object, totalCounts As Object, key As Variant
    Dim parts() As String, result() As Variant
    Dim r As Long, i As Long
    a = BI_F_ColumnVector(AnalyzeData)
    b = BI_F_ColumnVector(RelatedData)
    Set pairCounts = CreateObject("Scripting.Dictionary")
    Set totalCounts = CreateObject("Scripting.Dictionary")
    For r = 1 To Application.Min(UBound(a, 1), UBound(b, 1))
        key = CStr(a(r, 1)) & "<^>" & CStr(b(r, 1))
        pairCounts(key) = pairCounts(key) + 1
        totalCounts(CStr(a(r, 1))) = totalCounts(CStr(a(r, 1))) + 1
    Next r
    ReDim result(1 To pairCounts.Count, 1 To 5)
    i = 0
    For Each key In pairCounts.Keys
        i = i + 1
        parts = Split(CStr(key), "<^>")
        result(i, 1) = parts(0)
        result(i, 2) = parts(1)
        result(i, 3) = pairCounts(key)
        result(i, 4) = totalCounts(parts(0))
        result(i, 5) = pairCounts(key) / totalCounts(parts(0))
    Next key
    If pairCounts.Count > 1 Then BI_F_Sort2DByColumn result, 5, True
    BI_GEL1Test = result
    Exit Function
ErrHandler:
    BI_GEL1Test = CVErr(xlErrValue)
End Function

'/**
' * Description: Calculates Gestalt Element Link Test for one selected primary item.
' * Parameters: AnalyzeData - Primary category column; RelatedData - Related item column; DataItem - selected primary item; Total - optional value column.
' * Returns: Link-strength table for the selected item.
' */
Public Function BI_GEL2Test(ByVal AnalyzeData As Variant, ByVal RelatedData As Variant, ByVal DataItem As Variant, Optional ByVal Total As Variant) As Variant
    On Error GoTo ErrHandler
    Dim a As Variant, b As Variant, t As Variant
    Dim itemCounts As Object, totalCounts As Object, itemValues As Object, totalValues As Object
    Dim key As Variant, result() As Variant, r As Long, i As Long, hasTotal As Boolean
    a = BI_F_ColumnVector(AnalyzeData)
    b = BI_F_ColumnVector(RelatedData)
    hasTotal = Not IsMissing(Total)
    If hasTotal Then t = BI_F_ColumnVector(Total)
    Set itemCounts = CreateObject("Scripting.Dictionary")
    Set totalCounts = CreateObject("Scripting.Dictionary")
    Set itemValues = CreateObject("Scripting.Dictionary")
    Set totalValues = CreateObject("Scripting.Dictionary")
    For r = 1 To Application.Min(UBound(a, 1), UBound(b, 1))
        key = CStr(b(r, 1))
        totalCounts(key) = totalCounts(key) + 1
        If hasTotal And IsNumeric(t(r, 1)) Then totalValues(key) = totalValues(key) + CDbl(t(r, 1))
        If CStr(a(r, 1)) = CStr(DataItem) Then
            itemCounts(key) = itemCounts(key) + 1
            If hasTotal And IsNumeric(t(r, 1)) Then itemValues(key) = itemValues(key) + CDbl(t(r, 1))
        End If
    Next r
    If hasTotal Then
        ReDim result(1 To itemCounts.Count + 1, 1 To 7)
        result(1, 2) = "Count": result(1, 3) = "Total Count": result(1, 4) = "Value": result(1, 5) = "TotalValue": result(1, 6) = "Count/TotalCount": result(1, 7) = "Value/TotalValue"
    Else
        ReDim result(1 To itemCounts.Count + 1, 1 To 4)
        result(1, 2) = "Count": result(1, 3) = "Total Count": result(1, 4) = "Count/TotalCount"
    End If
    i = 1
    For Each key In itemCounts.Keys
        i = i + 1
        result(i, 1) = key
        result(i, 2) = itemCounts(key)
        result(i, 3) = totalCounts(key)
        If hasTotal Then
            result(i, 4) = itemValues(key)
            result(i, 5) = totalValues(key)
            result(i, 6) = itemCounts(key) / totalCounts(key)
            If totalValues(key) <> 0 Then result(i, 7) = itemValues(key) / totalValues(key)
        Else
            result(i, 4) = itemCounts(key) / totalCounts(key)
        End If
    Next key
    If itemCounts.Count > 1 Then BI_F_Sort2DByColumn result, IIf(hasTotal, 7, 4), True
    BI_GEL2Test = result
    Exit Function
ErrHandler:
    BI_GEL2Test = CVErr(xlErrValue)
End Function

Public Function BI_MUS_SampleSize(ByVal Tolerable As Double, ByVal Expected As Double, ByVal Risk As Double, Optional ByVal Start As Long = 1) As Long
    On Error GoTo ErrHandler
    Dim n As Long, x As Long
    n = IIf(Start < 1, 1, Start)
    Do
        x = Application.WorksheetFunction.RoundUp(Expected * n, 0)
        If BI_F_BinomDist(x, n, Tolerable) <= Risk Then Exit Do
        n = n + 1
        If n > 100000 Then Err.Raise vbObjectError + 701, "BI_MUS_SampleSize", "No sample size found within limit."
    Loop
    BI_MUS_SampleSize = n
    Exit Function
ErrHandler:
    BI_MUS_SampleSize = 0
End Function

Public Function BI_MUS_ExtractSample(ByVal Data As Variant, ByVal Tolerable As Double, ByVal Expected As Double, ByVal Risk As Double, Optional ByVal OnlySamples As Boolean = False) As Variant
    On Error GoTo ErrHandler
    Dim values() As Double, sampleSize As Long, population As Double, mu As Double
    Dim running As Double, previousFloor As Double, currentFloor As Double
    Dim rows() As Long, rowCount As Long, i As Long, outCols As Long, result() As Variant, outRow As Long
    values = BI_ToDoubleVector(Data, False)
    For i = LBound(values) To UBound(values): population = population + values(i): Next i
    sampleSize = BI_MUS_SampleSize(Tolerable, Expected, Risk, 1)
    If sampleSize <= 0 Then Err.Raise vbObjectError + 702, "BI_MUS_ExtractSample", "Invalid sample size."
    mu = population / sampleSize
    ReDim rows(1 To UBound(values))
    For i = LBound(values) To UBound(values)
        running = running + values(i)
        currentFloor = Int(running / mu) * mu
        If i > LBound(values) And currentFloor <> previousFloor Then
            rowCount = rowCount + 1
            rows(rowCount) = i
        End If
        previousFloor = currentFloor
    Next i
    outCols = IIf(OnlySamples, 2, 4)
    ReDim result(1 To rowCount + 4, 1 To outCols)
    result(1, 1) = "Population Size": result(1, 2) = UBound(values) - LBound(values) + 1
    result(2, 1) = "Sample Size": result(2, 2) = sampleSize
    result(3, 1) = "Monetary Unit": result(3, 2) = mu
    result(4, 1) = "RowID": result(4, 2) = "Sample"
    If Not OnlySamples Then result(4, 3) = "RunningTotal": result(4, 4) = "MU Hit"
    running = 0
    For i = LBound(values) To UBound(values)
        running = running + values(i)
        If rowCount > 0 Then
            Dim k As Long
            For k = 1 To rowCount
                If rows(k) = i Then
                    outRow = outRow + 1
                    result(outRow + 4, 1) = i
                    result(outRow + 4, 2) = values(i)
                    If Not OnlySamples Then
                        result(outRow + 4, 3) = running
                        result(outRow + 4, 4) = Int(running / mu) * mu
                    End If
                    Exit For
                End If
            Next k
        End If
    Next i
    BI_MUS_ExtractSample = result
    Exit Function
ErrHandler:
    BI_MUS_ExtractSample = CVErr(xlErrValue)
End Function

Private Function BI_MUS_Evaluate(ByVal Data As Variant, ByVal Sample As Variant, ByVal Audited As Variant, ByVal Tolerable As Double, ByVal Risk As Double, ByVal underStatement As Boolean) As Variant
    Dim population() As Double, sampleValues() As Double, auditedValues() As Double
    Dim popTotal As Double, sampleSize As Long, i As Long, n As Long
    Dim diffs() As Double, p As Double, temp As Double, j As Long
    Dim result() As Variant, upperLimit As Double, totalMis As Double, tolerableLimit As Double
    population = BI_ToDoubleVector(Data, False)
    sampleValues = BI_ToDoubleVector(Sample, False)
    auditedValues = BI_ToDoubleVector(Audited, False)
    For i = LBound(population) To UBound(population): popTotal = popTotal + population(i): Next i
    sampleSize = UBound(sampleValues) - LBound(sampleValues) + 1
    ReDim diffs(1 To sampleSize + 1)
    diffs(1) = 1
    n = 1
    For i = LBound(sampleValues) To UBound(sampleValues)
        If sampleValues(i) <> 0 Then
            If (underStatement And sampleValues(i) - auditedValues(i) > 0) Or ((Not underStatement) And sampleValues(i) - auditedValues(i) < 0) Then
                n = n + 1
                diffs(n) = Abs((sampleValues(i) - auditedValues(i)) / sampleValues(i))
            End If
        End If
    Next i
    ReDim Preserve diffs(1 To n)
    For i = 1 To n - 1
        For j = i + 1 To n
            If diffs(j) > diffs(i) Then temp = diffs(i): diffs(i) = diffs(j): diffs(j) = temp
        Next j
    Next i
    ReDim result(1 To n + 4, 1 To 4)
    result(1, 1) = "#": result(1, 2) = "Percentage of Misstatement": result(1, 3) = "Upper Limit": result(1, 4) = IIf(underStatement, "Estimated Understatement", "Estimated Overstatement")
    For i = 1 To n
        upperLimit = BI_F_UpperLimit(Risk, i - 1, sampleSize)
        result(i + 1, 1) = i - 1
        result(i + 1, 2) = diffs(i)
        result(i + 1, 3) = upperLimit
        result(i + 1, 4) = diffs(i) * upperLimit * popTotal
        totalMis = totalMis + result(i + 1, 4)
    Next i
    tolerableLimit = popTotal * Tolerable
    result(n + 2, 1) = "Total": result(n + 2, 4) = totalMis
    result(n + 3, 1) = "Tolerable": result(n + 3, 4) = tolerableLimit
    result(n + 4, 1) = "Result": result(n + 4, 4) = IIf(totalMis < tolerableLimit, "Accept", "Reject")
    BI_MUS_Evaluate = result
End Function

Public Function BI_MUS_Evaluate_OverStatement(ByVal Data As Variant, ByVal Sample As Variant, ByVal Audited As Variant, ByVal Tolerable As Double, ByVal Expected As Double, ByVal Risk As Double) As Variant
    On Error GoTo ErrHandler
    BI_MUS_Evaluate_OverStatement = BI_MUS_Evaluate(Data, Sample, Audited, Tolerable, Risk, False)
    Exit Function
ErrHandler:
    BI_MUS_Evaluate_OverStatement = CVErr(xlErrValue)
End Function

Public Function BI_MUS_Evaluate_UnderStatement(ByVal Data As Variant, ByVal Sample As Variant, ByVal Audited As Variant, ByVal Tolerable As Double, ByVal Expected As Double, ByVal Risk As Double) As Variant
    On Error GoTo ErrHandler
    BI_MUS_Evaluate_UnderStatement = BI_MUS_Evaluate(Data, Sample, Audited, Tolerable, Risk, True)
    Exit Function
ErrHandler:
    BI_MUS_Evaluate_UnderStatement = CVErr(xlErrValue)
End Function

'/**
' * Description: Returns rows where the selected numeric column contains round numbers.
' * Parameters: Data - Source range; ColumnIndex - numeric column to test; Digits - number of trailing zero digits.
' * Returns: Filtered rows plus summary metrics.
' */
Public Function BI_RoundNumbersTest(ByVal Data As Variant, ByVal ColumnIndex As Long, ByVal Digits As Long) As Variant
    On Error GoTo ErrHandler
    Dim src As Variant, result() As Variant
    Dim r As Long, c As Long, outRow As Long, roundRows As Long, totalRows As Long
    Dim roundSum As Double, totalSum As Double, denominator As Double, minCols As Long
    src = BI_F_SourceTo2D(Data)
    totalRows = UBound(src, 1)
    minCols = Application.Max(2, UBound(src, 2))
    denominator = 10 ^ Digits
    ReDim result(1 To totalRows + 7, 1 To minCols)
    For r = 1 To UBound(src, 1)
        If IsNumeric(src(r, ColumnIndex)) Then
            totalSum = totalSum + CDbl(src(r, ColumnIndex))
            If (CDbl(src(r, ColumnIndex)) - (Fix(CDbl(src(r, ColumnIndex)) / denominator) * denominator)) = 0 Then
                outRow = outRow + 1
                For c = 1 To UBound(src, 2): result(outRow, c) = src(r, c): Next c
                roundRows = roundRows + 1
                roundSum = roundSum + CDbl(src(r, ColumnIndex))
            End If
        End If
    Next r
    outRow = outRow + 2
    result(outRow, 1) = "Round Numbers Count": result(outRow, 2) = roundRows
    result(outRow + 1, 1) = "Total Count": result(outRow + 1, 2) = totalRows
    result(outRow + 2, 1) = "Round Amount Total": result(outRow + 2, 2) = roundSum
    result(outRow + 3, 1) = "Total Amount": result(outRow + 3, 2) = totalSum
    result(outRow + 4, 1) = "Count %": If totalRows <> 0 Then result(outRow + 4, 2) = roundRows / totalRows
    result(outRow + 5, 1) = "Amount %": If totalSum <> 0 Then result(outRow + 5, 2) = roundSum / totalSum
    BI_RoundNumbersTest = result
    Exit Function
ErrHandler:
    BI_RoundNumbersTest = CVErr(xlErrValue)
End Function

'/**
' * Description: Returns rows having duplicate combinations in selected columns.
' * Parameters: Data - Source table; TestColumns - one or more column indexes.
' * Returns: Matching source rows.
' */
Public Function BI_SameSameSameTest(ByVal Data As Variant, ByVal TestColumns As Variant) As Variant
    On Error GoTo ErrHandler
    Dim src As Variant, cols() As Long, counts As Object
    Dim r As Long, c As Long, outRow As Long, key As String, result() As Variant
    src = BI_F_SourceTo2D(Data)
    cols = BI_F_ToLongArray(TestColumns)
    Set counts = CreateObject("Scripting.Dictionary")
    For r = 1 To UBound(src, 1)
        key = BI_F_KeyFromColumns(src, r, cols)
        counts(key) = counts(key) + 1
    Next r
    ReDim result(1 To UBound(src, 1), 1 To UBound(src, 2))
    For r = 1 To UBound(src, 1)
        If counts(BI_F_KeyFromColumns(src, r, cols)) > 1 Then
            outRow = outRow + 1
            For c = 1 To UBound(src, 2): result(outRow, c) = src(r, c): Next c
        End If
    Next r
    If outRow = 0 Then BI_SameSameSameTest = CVErr(xlErrNA) Else BI_SameSameSameTest = BI_TakeRows(result, outRow)
    Exit Function
ErrHandler:
    BI_SameSameSameTest = CVErr(xlErrValue)
End Function

Private Function BI_TakeRows(ByRef source As Variant, ByVal rowCount As Long) As Variant
    Dim result() As Variant, r As Long, c As Long
    ReDim result(1 To rowCount, 1 To UBound(source, 2))
    For r = 1 To rowCount
        For c = 1 To UBound(source, 2)
            result(r, c) = source(r, c)
        Next c
    Next r
    BI_TakeRows = result
End Function

'/**
' * Description: Returns combinations where selected columns are same but another column differs.
' * Parameters: Data - Source table; TestColumns - same columns; DifferentColumn - column expected to differ.
' * Returns: Unique suspicious combinations.
' */
Public Function BI_SameSameDifferentTest(ByVal Data As Variant, ByVal TestColumns As Variant, ByVal DifferentColumn As Long) As Variant
    On Error GoTo ErrHandler
    Dim src As Variant, cols() As Long, groupValues As Object, groupRows As Object
    Dim r As Long, c As Long, i As Long, key As Variant, outRow As Long, result() As Variant
    src = BI_F_SourceTo2D(Data)
    cols = BI_F_ToLongArray(TestColumns)
    Set groupValues = CreateObject("Scripting.Dictionary")
    Set groupRows = CreateObject("Scripting.Dictionary")
    For r = 1 To UBound(src, 1)
        key = BI_F_KeyFromColumns(src, r, cols)
        If Not groupValues.Exists(key) Then
            groupValues.Add key, CreateObject("Scripting.Dictionary")
            groupRows.Add key, New Collection
        End If
        groupValues(key)(UCase$(CStr(src(r, DifferentColumn)))) = True
        groupRows(key).Add r
    Next r
    ReDim result(1 To UBound(src, 1), 1 To UBound(cols) - LBound(cols) + 2)
    For Each key In groupValues.Keys
        If groupValues(key).Count > 1 Then
            For i = 1 To groupRows(key).Count
                r = groupRows(key)(i)
                outRow = outRow + 1
                For c = LBound(cols) To UBound(cols)
                    result(outRow, c - LBound(cols) + 1) = src(r, cols(c))
                Next c
                result(outRow, UBound(cols) - LBound(cols) + 2) = src(r, DifferentColumn)
            Next i
        End If
    Next key
    If outRow = 0 Then BI_SameSameDifferentTest = CVErr(xlErrNA) Else BI_SameSameDifferentTest = BI_TakeRows(result, outRow)
    Exit Function
ErrHandler:
    BI_SameSameDifferentTest = CVErr(xlErrValue)
End Function

'/**
' * Description: Calculates the Subset Number Duplication Test.
' * Parameters: CategoryData - category/group column; SubsetData - subset number column.
' * Returns: Category and frequency factor sorted descending.
' */
Public Function BI_SubsetNumberDuplicationTest(ByVal CategoryData As Variant, ByVal SubsetData As Variant) As Variant
    On Error GoTo ErrHandler
    Dim cats As Variant, subs As Variant, pairCounts As Object, categoryCounts As Object
    Dim key As Variant, parts() As String, cat As String, r As Long
    Dim sumSq As Object, result() As Variant, outRow As Long
    cats = BI_F_ColumnVector(CategoryData)
    subs = BI_F_ColumnVector(SubsetData)
    Set pairCounts = CreateObject("Scripting.Dictionary")
    Set categoryCounts = CreateObject("Scripting.Dictionary")
    Set sumSq = CreateObject("Scripting.Dictionary")
    For r = 1 To Application.Min(UBound(cats, 1), UBound(subs, 1))
        cat = CStr(cats(r, 1))
        key = cat & "<^>" & CStr(subs(r, 1))
        pairCounts(key) = pairCounts(key) + 1
        categoryCounts(cat) = categoryCounts(cat) + 1
    Next r
    For Each key In pairCounts.Keys
        If pairCounts(key) <> 1 Then
            parts = Split(CStr(key), "<^>")
            sumSq(parts(0)) = sumSq(parts(0)) + (pairCounts(key) ^ 2)
        End If
    Next key
    ReDim result(1 To sumSq.Count, 1 To 2)
    For Each key In sumSq.Keys
        outRow = outRow + 1
        result(outRow, 1) = key
        result(outRow, 2) = sumSq(key) / (categoryCounts(key) ^ 2)
    Next key
    If outRow > 1 Then BI_F_Sort2DByColumn result, 2, True
    BI_SubsetNumberDuplicationTest = result
    Exit Function
ErrHandler:
    BI_SubsetNumberDuplicationTest = CVErr(xlErrValue)
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

Private Function BI_PickScoreTemplateTarget(ByVal modelName As String) As Range
    On Error GoTo ErrHandler
    Dim picked As Range
    Dim prompt As String

    prompt = "Select the top-left cell where the " & modelName & " score template should be created."
    Set picked = Application.InputBox(prompt, "BeIndian Score Template", ActiveCell.Address(External:=True), Type:=8)
    If picked Is Nothing Then Exit Function
    Set picked = picked.Cells(1, 1)

    If Len(CStr(picked.Value)) > 0 Or picked.CurrentRegion.Cells.CountLarge > 1 Then
        If MsgBox("The selected area is not blank. Creating the template may overwrite existing data. Continue?", _
                  vbQuestion + vbYesNo, "BeIndian Score Template") <> vbYes Then
            Exit Function
        End If
    End If

    Set BI_PickScoreTemplateTarget = picked
    Exit Function
ErrHandler:
    Set BI_PickScoreTemplateTarget = Nothing
End Function

Public Sub BI_Tool_ScoreInputTemplate(control As IRibbonControl)
    On Error GoTo ErrHandler
    Dim modelName As String
    Dim target As Range
    modelName = Application.InputBox("Enter score model: Altman, Beneish, Ohlson, or Piotroski", "BeIndian Score Template", "Altman", Type:=2)
    If Len(modelName) = 0 Then Exit Sub
    Set target = BI_PickScoreTemplateTarget(modelName)
    If target Is Nothing Then Exit Sub
    BI_OutputScoreInputTemplate modelName, target
    Exit Sub
ErrHandler:
    MsgBox "Score template failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_AltmanTemplate(control As IRibbonControl)
    On Error GoTo ErrHandler
    Dim target As Range
    Set target = BI_PickScoreTemplateTarget("Altman")
    If target Is Nothing Then Exit Sub
    BI_OutputScoreInputTemplate "Altman", target
    Exit Sub
ErrHandler:
    MsgBox "Altman template failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_BeneishTemplate(control As IRibbonControl)
    On Error GoTo ErrHandler
    Dim target As Range
    Set target = BI_PickScoreTemplateTarget("Beneish")
    If target Is Nothing Then Exit Sub
    BI_OutputScoreInputTemplate "Beneish", target
    Exit Sub
ErrHandler:
    MsgBox "Beneish template failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_OhlsonTemplate(control As IRibbonControl)
    On Error GoTo ErrHandler
    Dim target As Range
    Set target = BI_PickScoreTemplateTarget("Ohlson")
    If target Is Nothing Then Exit Sub
    BI_OutputScoreInputTemplate "Ohlson", target
    Exit Sub
ErrHandler:
    MsgBox "Ohlson template failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_PiotroskiTemplate(control As IRibbonControl)
    On Error GoTo ErrHandler
    Dim target As Range
    Set target = BI_PickScoreTemplateTarget("Piotroski")
    If target Is Nothing Then Exit Sub
    BI_OutputScoreInputTemplate "Piotroski", target
    Exit Sub
ErrHandler:
    MsgBox "Piotroski template failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_BeneishMScore(control As IRibbonControl)
    On Error GoTo ErrHandler
    Unload frmBI_Score
    frmBI_Score.Tag = "Beneish"
    frmBI_Score.Show
    Exit Sub
ErrHandler:
    MsgBox "Beneish M-Score failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_OhlsonOScore(control As IRibbonControl)
    On Error GoTo ErrHandler
    Unload frmBI_Score
    frmBI_Score.Tag = "Ohlson"
    frmBI_Score.Show
    Exit Sub
ErrHandler:
    MsgBox "Ohlson O-Score failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_PiotroskiFScore(control As IRibbonControl)
    On Error GoTo ErrHandler
    Unload frmBI_Score
    frmBI_Score.Tag = "Piotroski"
    frmBI_Score.Show
    Exit Sub
ErrHandler:
    MsgBox "Piotroski F-Score failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_AltmanZScore(control As IRibbonControl)
    On Error GoTo ErrHandler
    Unload frmBI_Score
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
