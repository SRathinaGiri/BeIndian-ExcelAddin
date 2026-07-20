Attribute VB_Name = "modBI_Stats"
Option Explicit

Private Function BI_DataRangeTo2D(ByVal source As Variant) As Variant
    Dim data As Variant
    Dim result() As Variant
    Dim r As Long, c As Long
    If TypeName(source) = "Range" Then
        data = source.Value2
    Else
        data = source
    End If
    If IsArray(data) Then
        On Error GoTo OneDimensional
        ReDim result(1 To UBound(data, 1) - LBound(data, 1) + 1, 1 To UBound(data, 2) - LBound(data, 2) + 1)
        For r = 1 To UBound(result, 1)
            For c = 1 To UBound(result, 2)
                result(r, c) = data(LBound(data, 1) + r - 1, LBound(data, 2) + c - 1)
            Next c
        Next r
        BI_DataRangeTo2D = result
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
    BI_DataRangeTo2D = result
End Function

Private Function BI_StatsBody(ByVal SourceRange As Variant, Optional ByVal Labels As Boolean = False) As Variant
    Dim data As Variant
    Dim result() As Variant
    Dim r As Long, c As Long
    Dim startRow As Long

    data = BI_DataRangeTo2D(SourceRange)
    startRow = IIf(Labels, 2, 1)
    ReDim result(1 To UBound(data, 1) - startRow + 1, 1 To UBound(data, 2))
    For r = startRow To UBound(data, 1)
        For c = 1 To UBound(data, 2)
            result(r - startRow + 1, c) = data(r, c)
        Next c
    Next r
    BI_StatsBody = result
End Function

Private Function BI_StatsHeaders(ByVal SourceRange As Variant, Optional ByVal Labels As Boolean = False, Optional ByVal Prefix As String = "Column ") As Variant
    Dim data As Variant
    Dim result() As Variant
    Dim c As Long

    data = BI_DataRangeTo2D(SourceRange)
    ReDim result(1 To 1, 1 To UBound(data, 2))
    For c = 1 To UBound(data, 2)
        If Labels Then
            result(1, c) = CStr(data(1, c))
        Else
            result(1, c) = Prefix & c
        End If
    Next c
    BI_StatsHeaders = result
End Function

Private Function BI_StatsColumnToVector(ByVal data As Variant, ByVal columnIndex As Long) As Double()
    Dim result() As Double
    Dim r As Long, n As Long

    ReDim result(1 To UBound(data, 1))
    For r = 1 To UBound(data, 1)
        If IsNumeric(data(r, columnIndex)) Then
            n = n + 1
            result(n) = CDbl(data(r, columnIndex))
        Else
            Err.Raise vbObjectError + 540, "BI_StatsColumnToVector", "Non-numeric value found in statistics source range."
        End If
    Next r
    ReDim Preserve result(1 To n)
    BI_StatsColumnToVector = result
End Function

Private Function BI_StatsColumnMatrix(ByVal data As Variant, ByVal columnIndex As Long) As Variant
    Dim result() As Variant
    Dim r As Long
    ReDim result(1 To UBound(data, 1), 1 To 1)
    For r = 1 To UBound(data, 1)
        result(r, 1) = data(r, columnIndex)
    Next r
    BI_StatsColumnMatrix = result
End Function

Private Sub BI_SortSummaryRows(ByRef data As Variant, ByVal sortColumn As Long, Optional ByVal descending As Boolean = True)
    Dim i As Long, j As Long, c As Long
    Dim temp As Variant
    For i = 1 To UBound(data, 1) - 1
        For j = i + 1 To UBound(data, 1)
            If (descending And CDbl(data(j, sortColumn)) > CDbl(data(i, sortColumn))) _
                Or ((Not descending) And CDbl(data(j, sortColumn)) < CDbl(data(i, sortColumn))) Then
                For c = 1 To UBound(data, 2)
                    temp = data(i, c)
                    data(i, c) = data(j, c)
                    data(j, c) = temp
                Next c
            End If
        Next j
    Next i
End Sub

Private Function BI_ConfidenceCaption(ByVal Confidence As Double) As String
    BI_ConfidenceCaption = Format$(Confidence * 100, "0.##") & "%"
End Function

Private Function BI_PopulationCovariance(ByRef valuesA() As Double, ByRef valuesB() As Double) As Double
    Dim i As Long
    Dim meanA As Double, meanB As Double
    Dim total As Double
    For i = LBound(valuesA) To UBound(valuesA)
        meanA = meanA + valuesA(i)
        meanB = meanB + valuesB(i)
    Next i
    meanA = meanA / (UBound(valuesA) - LBound(valuesA) + 1)
    meanB = meanB / (UBound(valuesB) - LBound(valuesB) + 1)
    For i = LBound(valuesA) To UBound(valuesA)
        total = total + ((valuesA(i) - meanA) * (valuesB(i) - meanB))
    Next i
    BI_PopulationCovariance = total / (UBound(valuesA) - LBound(valuesA) + 1)
End Function

Private Function BI_CorrelationValue(ByRef valuesA() As Double, ByRef valuesB() As Double) As Double
    Dim i As Long
    Dim meanA As Double, meanB As Double
    Dim varA As Double, varB As Double, covar As Double
    For i = LBound(valuesA) To UBound(valuesA)
        meanA = meanA + valuesA(i)
        meanB = meanB + valuesB(i)
    Next i
    meanA = meanA / (UBound(valuesA) - LBound(valuesA) + 1)
    meanB = meanB / (UBound(valuesB) - LBound(valuesB) + 1)
    For i = LBound(valuesA) To UBound(valuesA)
        covar = covar + ((valuesA(i) - meanA) * (valuesB(i) - meanB))
        varA = varA + ((valuesA(i) - meanA) ^ 2)
        varB = varB + ((valuesB(i) - meanB) ^ 2)
    Next i
    If varA = 0 Or varB = 0 Then
        BI_CorrelationValue = 0
    Else
        BI_CorrelationValue = covar / Sqr(varA * varB)
    End If
End Function

Private Function BI_RowCountNumeric(ByVal data As Variant, ByVal rowIndex As Long) As Long
    Dim c As Long
    For c = 1 To UBound(data, 2)
        If IsNumeric(data(rowIndex, c)) Then BI_RowCountNumeric = BI_RowCountNumeric + 1
    Next c
End Function

Private Function BI_RowSumNumeric(ByVal data As Variant, ByVal rowIndex As Long) As Double
    Dim c As Long
    For c = 1 To UBound(data, 2)
        If IsNumeric(data(rowIndex, c)) Then BI_RowSumNumeric = BI_RowSumNumeric + CDbl(data(rowIndex, c))
    Next c
End Function

Private Function BI_RowAverageNumeric(ByVal data As Variant, ByVal rowIndex As Long) As Double
    Dim countValue As Long
    countValue = BI_RowCountNumeric(data, rowIndex)
    If countValue > 0 Then BI_RowAverageNumeric = BI_RowSumNumeric(data, rowIndex) / countValue
End Function

Private Function BI_RowVarianceSample(ByVal data As Variant, ByVal rowIndex As Long) As Double
    Dim values() As Double
    Dim c As Long, n As Long
    ReDim values(1 To UBound(data, 2))
    For c = 1 To UBound(data, 2)
        If IsNumeric(data(rowIndex, c)) Then
            n = n + 1
            values(n) = CDbl(data(rowIndex, c))
        End If
    Next c
    If n > 1 Then
        ReDim Preserve values(1 To n)
        BI_RowVarianceSample = Application.WorksheetFunction.Var_S(values)
    End If
End Function

Private Function BI_RowDevSq(ByVal data As Variant, ByVal rowIndex As Long) As Double
    Dim avgValue As Double
    Dim c As Long
    avgValue = BI_RowAverageNumeric(data, rowIndex)
    For c = 1 To UBound(data, 2)
        If IsNumeric(data(rowIndex, c)) Then
            BI_RowDevSq = BI_RowDevSq + ((CDbl(data(rowIndex, c)) - avgValue) ^ 2)
        End If
    Next c
End Function

Private Function BI_ColumnCountNumeric(ByVal data As Variant, ByVal columnIndex As Long) As Long
    Dim r As Long
    For r = 1 To UBound(data, 1)
        If IsNumeric(data(r, columnIndex)) Then BI_ColumnCountNumeric = BI_ColumnCountNumeric + 1
    Next r
End Function

Private Function BI_ColumnSumNumeric(ByVal data As Variant, ByVal columnIndex As Long) As Double
    Dim r As Long
    For r = 1 To UBound(data, 1)
        If IsNumeric(data(r, columnIndex)) Then BI_ColumnSumNumeric = BI_ColumnSumNumeric + CDbl(data(r, columnIndex))
    Next r
End Function

Private Function BI_ColumnAverageNumeric(ByVal data As Variant, ByVal columnIndex As Long) As Double
    Dim countValue As Long
    countValue = BI_ColumnCountNumeric(data, columnIndex)
    If countValue > 0 Then BI_ColumnAverageNumeric = BI_ColumnSumNumeric(data, columnIndex) / countValue
End Function

Private Function BI_ColumnVarianceSample(ByVal data As Variant, ByVal columnIndex As Long) As Double
    Dim values() As Double
    Dim r As Long, n As Long
    ReDim values(1 To UBound(data, 1))
    For r = 1 To UBound(data, 1)
        If IsNumeric(data(r, columnIndex)) Then
            n = n + 1
            values(n) = CDbl(data(r, columnIndex))
        End If
    Next r
    If n > 1 Then
        ReDim Preserve values(1 To n)
        BI_ColumnVarianceSample = Application.WorksheetFunction.Var_S(values)
    End If
End Function

'/**
' * Description: Returns descriptive statistics for a numeric range.
' * Parameters: SourceRange - Numeric source range.
' * Returns: Two-column descriptive statistics table.
' */
Public Function BI_Describe(ByVal SourceRange As Variant) As Variant
    On Error GoTo ErrHandler
    Dim values() As Double
    Dim result(1 To 23, 1 To 2) As Variant
    Dim q1 As Double, q3 As Double, iqr As Double, whisker As Double

    values = BI_ToDoubleVector(SourceRange, False)
    q1 = Application.WorksheetFunction.Quartile_Inc(values, 1)
    q3 = Application.WorksheetFunction.Quartile_Inc(values, 3)
    iqr = q3 - q1
    whisker = iqr * 1.5

    result(1, 1) = "Descriptive Summary Statistics"
    result(2, 1) = "Average": result(2, 2) = Application.WorksheetFunction.Average(values)
    result(3, 1) = "Median": result(3, 2) = Application.WorksheetFunction.Median(values)
    On Error Resume Next
    result(4, 1) = "Mode": result(4, 2) = Application.WorksheetFunction.Mode_Sngl(values)
    If Err.Number <> 0 Then result(4, 2) = "-"
    Err.Clear
    On Error GoTo ErrHandler
    result(5, 1) = "St.Dev. P.": result(5, 2) = Application.WorksheetFunction.StDev_P(values)
    result(6, 1) = "St. Dev. S": result(6, 2) = Application.WorksheetFunction.StDev_S(values)
    result(7, 1) = "Std. Error": result(7, 2) = result(6, 2) / Sqr(UBound(values) - LBound(values) + 1)
    result(8, 1) = "Population Variance": result(8, 2) = Application.WorksheetFunction.Var_P(values)
    result(9, 1) = "Sample Variance": result(9, 2) = Application.WorksheetFunction.Var_S(values)
    result(10, 1) = "Kurtosis": result(10, 2) = Application.WorksheetFunction.Kurt(values)
    result(11, 1) = "Skewness.P": result(11, 2) = Application.WorksheetFunction.Skew_P(values)
    result(12, 1) = "Skewness S": result(12, 2) = Application.WorksheetFunction.Skew(values)
    result(13, 1) = "Range": result(13, 2) = Application.WorksheetFunction.Max(values) - Application.WorksheetFunction.Min(values)
    result(14, 1) = "Maximum": result(14, 2) = Application.WorksheetFunction.Max(values)
    result(15, 1) = "Minimum": result(15, 2) = Application.WorksheetFunction.Min(values)
    result(16, 1) = "Sum": result(16, 2) = Application.WorksheetFunction.Sum(values)
    result(17, 1) = "Count": result(17, 2) = UBound(values) - LBound(values) + 1
    result(18, 1) = "Quartile1": result(18, 2) = q1
    result(19, 1) = "Quartile3": result(19, 2) = q3
    result(20, 1) = "IQR": result(20, 2) = iqr
    result(21, 1) = "1.5 Times IQR": result(21, 2) = whisker
    result(22, 1) = "Whisker Lower Bound": result(22, 2) = Application.Max(q1 - whisker, Application.WorksheetFunction.Min(values))
    result(23, 1) = "Whisker Outer Bound": result(23, 2) = Application.Min(q3 + whisker, Application.WorksheetFunction.Max(values))
    BI_Describe = result
    Exit Function
ErrHandler:
    BI_Describe = CVErr(xlErrValue)
End Function

'/**
' * Description: Calculates Anova (Analysis of Variance) Single Factor for numeric columns.
' * Parameters: TableData - Numeric table grouped by columns; Alpha - Significance level; Labels - Optional header flag.
' * Returns: Summary and ANOVA table.
' */
Public Function BI_AnovaSingleFactor(ByVal TableData As Variant, Optional ByVal Alpha As Double = 0.05, Optional ByVal Labels As Boolean = False) As Variant
    On Error GoTo ErrHandler
    Dim data As Variant, headers As Variant
    Dim colCount As Long, r As Long, c As Long, outRow As Long
    Dim counts() As Long, sums() As Double, averages() As Double, variances() As Double
    Dim totalCount As Long, totalSum As Double
    Dim betweenSS As Double, withinSS As Double, totalSS As Double
    Dim betweenDf As Long, withinDf As Long
    Dim betweenMS As Double, withinMS As Double, fValue As Double, pValue As Double, fCrit As Double
    Dim result() As Variant

    data = BI_StatsBody(TableData, Labels)
    headers = BI_StatsHeaders(TableData, Labels)
    colCount = UBound(data, 2)
    ReDim counts(1 To colCount)
    ReDim sums(1 To colCount)
    ReDim averages(1 To colCount)
    ReDim variances(1 To colCount)

    For c = 1 To colCount
        counts(c) = BI_ColumnCountNumeric(data, c)
        sums(c) = BI_ColumnSumNumeric(data, c)
        If counts(c) > 0 Then averages(c) = sums(c) / counts(c)
        If counts(c) > 1 Then variances(c) = BI_ColumnVarianceSample(data, c)
        totalCount = totalCount + counts(c)
        totalSum = totalSum + sums(c)
    Next c

    For c = 1 To colCount
        betweenSS = betweenSS + (sums(c) * averages(c))
        withinSS = withinSS + Application.WorksheetFunction.DevSq(BI_StatsColumnToVector(data, c))
    Next c
    betweenSS = betweenSS - ((totalSum ^ 2) / totalCount)
    totalSS = betweenSS + withinSS
    betweenDf = colCount - 1
    withinDf = totalCount - colCount
    betweenMS = betweenSS / betweenDf
    withinMS = withinSS / withinDf
    fValue = betweenMS / withinMS
    pValue = Application.WorksheetFunction.FDist(fValue, betweenDf, withinDf)
    fCrit = Application.WorksheetFunction.FInv(Alpha, betweenDf, withinDf)

    ReDim result(1 To 12 + colCount, 1 To 7)
    result(1, 1) = "Anova: Single Factor"
    result(3, 1) = "Summary"
    result(4, 1) = "Groups": result(4, 2) = "Count": result(4, 3) = "Sum": result(4, 4) = "Average": result(4, 5) = "Variance"
    For c = 1 To colCount
        result(4 + c, 1) = headers(1, c)
        result(4 + c, 2) = counts(c)
        result(4 + c, 3) = sums(c)
        result(4 + c, 4) = averages(c)
        result(4 + c, 5) = variances(c)
    Next c

    outRow = 7 + colCount
    result(outRow, 1) = "Anova"
    result(outRow + 1, 1) = "Source of Variation"
    result(outRow + 1, 2) = "SS"
    result(outRow + 1, 3) = "DF"
    result(outRow + 1, 4) = "MS"
    result(outRow + 1, 5) = "F"
    result(outRow + 1, 6) = "P-value"
    result(outRow + 1, 7) = "F Crit"
    result(outRow + 2, 1) = "Between Groups": result(outRow + 2, 2) = betweenSS: result(outRow + 2, 3) = betweenDf: result(outRow + 2, 4) = betweenMS: result(outRow + 2, 5) = fValue: result(outRow + 2, 6) = pValue: result(outRow + 2, 7) = fCrit
    result(outRow + 3, 1) = "Within Groups": result(outRow + 3, 2) = withinSS: result(outRow + 3, 3) = withinDf: result(outRow + 3, 4) = withinMS
    result(outRow + 5, 1) = "Total": result(outRow + 5, 2) = totalSS: result(outRow + 5, 3) = betweenDf + withinDf

    BI_AnovaSingleFactor = result
    Exit Function
ErrHandler:
    BI_AnovaSingleFactor = CVErr(xlErrValue)
End Function

'/**
' * Description: Calculates Anova (Analysis of Variance) without replication for a numeric matrix.
' * Parameters: Data - Numeric matrix; Alpha - Significance level; Labels - Optional header flag.
' * Returns: Summary and ANOVA table.
' */
Public Function BI_AnovaTwoFactorsWithoutReplication(ByVal Data As Variant, Optional ByVal Alpha As Double = 0.05, Optional ByVal Labels As Boolean = False) As Variant
    On Error GoTo ErrHandler
    Dim body As Variant, rowHeaders() As Variant, colHeaders As Variant
    Dim rowCount As Long, colCount As Long, r As Long, c As Long, outRow As Long
    Dim rowCounts() As Long, rowSums() As Double, rowAvgs() As Double, rowVars() As Double
    Dim colCounts() As Long, colSums() As Double, colAvgs() As Double, colVars() As Double
    Dim totalCount As Long, totalSum As Double, totalSq As Double
    Dim rowSS As Double, colSS As Double, errSS As Double, totalSS As Double
    Dim rowDf As Long, colDf As Long, totalDf As Long, errDf As Long
    Dim rowMS As Double, colMS As Double, errMS As Double
    Dim rowF As Double, colF As Double, rowP As Double, colP As Double, rowFCrit As Double, colFCrit As Double
    Dim result() As Variant
    Dim raw As Variant

    raw = BI_DataRangeTo2D(Data)
    If Labels Then
        rowCount = UBound(raw, 1) - 1
        colCount = UBound(raw, 2) - 1
        ReDim rowHeaders(1 To rowCount, 1 To 1)
        For r = 1 To rowCount
            rowHeaders(r, 1) = raw(r + 1, 1)
        Next r
        colHeaders = BI_StatsHeaders(Data, True)
        body = BI_StatsBody(Data, True)
    Else
        body = BI_DataRangeTo2D(Data)
        rowCount = UBound(body, 1)
        colCount = UBound(body, 2)
        ReDim rowHeaders(1 To rowCount, 1 To 1)
        ReDim colHeaders(1 To 1, 1 To colCount)
        For r = 1 To rowCount
            rowHeaders(r, 1) = "Row" & r
        Next r
        For c = 1 To colCount
            colHeaders(1, c) = "Column" & c
        Next c
    End If

    ReDim rowCounts(1 To rowCount)
    ReDim rowSums(1 To rowCount)
    ReDim rowAvgs(1 To rowCount)
    ReDim rowVars(1 To rowCount)
    ReDim colCounts(1 To colCount)
    ReDim colSums(1 To colCount)
    ReDim colAvgs(1 To colCount)
    ReDim colVars(1 To colCount)

    For r = 1 To rowCount
        rowCounts(r) = BI_RowCountNumeric(body, r)
        rowSums(r) = BI_RowSumNumeric(body, r)
        rowAvgs(r) = BI_RowAverageNumeric(body, r)
        If rowCounts(r) > 1 Then rowVars(r) = BI_RowVarianceSample(body, r)
        totalCount = totalCount + rowCounts(r)
        totalSum = totalSum + rowSums(r)
    Next r
    For c = 1 To colCount
        colCounts(c) = BI_ColumnCountNumeric(body, c)
        colSums(c) = BI_ColumnSumNumeric(body, c)
        colAvgs(c) = BI_ColumnAverageNumeric(body, c)
        If colCounts(c) > 1 Then colVars(c) = BI_ColumnVarianceSample(body, c)
    Next c
    For r = 1 To rowCount
        For c = 1 To colCount
            If IsNumeric(body(r, c)) Then totalSq = totalSq + (CDbl(body(r, c)) ^ 2)
        Next c
    Next r

    For r = 1 To rowCount
        rowSS = rowSS + (rowSums(r) * rowAvgs(r))
    Next r
    rowSS = rowSS - ((totalSum ^ 2) / totalCount)
    For c = 1 To colCount
        colSS = colSS + (colSums(c) * colAvgs(c))
    Next c
    colSS = colSS - ((totalSum ^ 2) / totalCount)
    errSS = totalSq + ((totalSum ^ 2) / totalCount) - (rowSS + colSS + ((totalSum ^ 2) / totalCount))
    totalSS = rowSS + colSS + errSS
    rowDf = rowCount - 1
    colDf = colCount - 1
    totalDf = totalCount - 1
    errDf = totalDf - rowDf - colDf
    rowMS = rowSS / rowDf
    colMS = colSS / colDf
    errMS = errSS / errDf
    rowF = rowMS / errMS
    colF = colMS / errMS
    rowP = Application.WorksheetFunction.FDist(rowF, rowDf, errDf)
    colP = Application.WorksheetFunction.FDist(colF, colDf, errDf)
    rowFCrit = Application.WorksheetFunction.FInv(Alpha, rowDf, errDf)
    colFCrit = Application.WorksheetFunction.FInv(Alpha, colDf, errDf)

    ReDim result(1 To 13 + rowCount + colCount, 1 To 7)
    result(1, 1) = "Anova: Two-Factor Without Replication"
    result(3, 1) = "Summary"
    result(4, 1) = "Rows": result(4, 2) = "Count": result(4, 3) = "Sum": result(4, 4) = "Average": result(4, 5) = "Variance"
    For r = 1 To rowCount
        result(4 + r, 1) = rowHeaders(r, 1)
        result(4 + r, 2) = rowCounts(r)
        result(4 + r, 3) = rowSums(r)
        result(4 + r, 4) = rowAvgs(r)
        result(4 + r, 5) = rowVars(r)
    Next r

    outRow = 6 + rowCount
    result(outRow, 1) = "Columns"
    result(outRow, 2) = "Count"
    result(outRow, 3) = "Sum"
    result(outRow, 4) = "Average"
    result(outRow, 5) = "Variance"
    For c = 1 To colCount
        result(outRow + c, 1) = colHeaders(1, c)
        result(outRow + c, 2) = colCounts(c)
        result(outRow + c, 3) = colSums(c)
        result(outRow + c, 4) = colAvgs(c)
        result(outRow + c, 5) = colVars(c)
    Next c

    outRow = outRow + colCount + 2
    result(outRow, 1) = "Anova"
    result(outRow + 1, 1) = "Source of Variation"
    result(outRow + 1, 2) = "SS"
    result(outRow + 1, 3) = "df"
    result(outRow + 1, 4) = "MS"
    result(outRow + 1, 5) = "F"
    result(outRow + 1, 6) = "P-value"
    result(outRow + 1, 7) = "F crit"
    result(outRow + 2, 1) = "Rows": result(outRow + 2, 2) = rowSS: result(outRow + 2, 3) = rowDf: result(outRow + 2, 4) = rowMS: result(outRow + 2, 5) = rowF: result(outRow + 2, 6) = rowP: result(outRow + 2, 7) = rowFCrit
    result(outRow + 3, 1) = "Columns": result(outRow + 3, 2) = colSS: result(outRow + 3, 3) = colDf: result(outRow + 3, 4) = colMS: result(outRow + 3, 5) = colF: result(outRow + 3, 6) = colP: result(outRow + 3, 7) = colFCrit
    result(outRow + 4, 1) = "Error": result(outRow + 4, 2) = errSS: result(outRow + 4, 3) = errDf: result(outRow + 4, 4) = errMS
    result(outRow + 6, 1) = "Total": result(outRow + 6, 2) = totalSS: result(outRow + 6, 3) = totalDf

    BI_AnovaTwoFactorsWithoutReplication = result
    Exit Function
ErrHandler:
    BI_AnovaTwoFactorsWithoutReplication = CVErr(xlErrValue)
End Function

'/**
' * Description: Returns the fixed Anscombe Quartet data and descriptive statistics.
' * Parameters: None.
' * Returns: Eight-column table.
' */
Public Function BI_AnsCombeQuartet() As Variant
    Dim result(1 To 18, 1 To 8) As Variant
    Dim xVals As Variant, y1 As Variant, y2 As Variant, y3 As Variant, x4 As Variant, y4 As Variant
    Dim i As Long

    xVals = Array(10, 8, 13, 9, 11, 14, 6, 4, 12, 7, 5)
    y1 = Array(8.04, 6.95, 7.58, 8.81, 8.33, 9.96, 7.24, 4.26, 10.84, 4.82, 5.68)
    y2 = Array(9.14, 8.14, 8.74, 8.77, 9.26, 8.1, 6.13, 3.1, 9.13, 7.26, 4.74)
    y3 = Array(7.46, 6.77, 12.74, 7.11, 7.81, 8.84, 6.08, 5.39, 8.15, 6.42, 5.73)
    x4 = Array(8, 8, 8, 8, 8, 8, 8, 19, 8, 8, 8)
    y4 = Array(6.58, 5.76, 7.71, 8.84, 8.47, 7.04, 5.25, 12.5, 5.56, 7.91, 6.89)

    result(1, 1) = "X1": result(1, 2) = "Y1": result(1, 3) = "X2": result(1, 4) = "Y2"
    result(1, 5) = "X3": result(1, 6) = "Y3": result(1, 7) = "X4": result(1, 8) = "Y4"
    For i = 0 To 10
        result(i + 2, 1) = xVals(i): result(i + 2, 2) = y1(i)
        result(i + 2, 3) = xVals(i): result(i + 2, 4) = y2(i)
        result(i + 2, 5) = xVals(i): result(i + 2, 6) = y3(i)
        result(i + 2, 7) = x4(i): result(i + 2, 8) = y4(i)
    Next i
    result(13, 1) = "Descriptive Statistics"
    result(14, 1) = "Average X": result(14, 2) = 9
    result(15, 1) = "Average Y": result(15, 2) = 7.5
    result(16, 1) = "Sample Variance": result(16, 2) = 11
    result(17, 1) = "XY Correlation": result(17, 2) = 0.816
    result(18, 1) = "R Squared": result(18, 2) = 0.67
    BI_AnsCombeQuartet = result
End Function

'/**
' * Description: Creates four scatter charts for the Anscombe Quartet table written by BI_AnsCombeQuartet.
' * Parameters: OutputCell - Top-left cell of the Anscombe output table.
' * Returns: None.
' */
Public Sub BI_CreateAnscombeQuartetChart(ByVal OutputCell As Range)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Dim i As Long
    Dim chartObj As ChartObject
    Dim xRange As Range, yRange As Range
    Dim chartLeft As Double, chartTop As Double
    Dim chartWidth As Double, chartHeight As Double

    Set ws = OutputCell.Worksheet
    chartWidth = 240
    chartHeight = 170
    chartLeft = OutputCell.Offset(0, 10).Left
    chartTop = OutputCell.Top

    For i = 1 To 4
        Set xRange = OutputCell.Offset(1, (i - 1) * 2).Resize(11, 1)
        Set yRange = OutputCell.Offset(1, (i - 1) * 2 + 1).Resize(11, 1)
        Set chartObj = ws.ChartObjects.Add( _
            chartLeft + ((i - 1) Mod 2) * (chartWidth + 18), _
            chartTop + ((i - 1) \ 2) * (chartHeight + 28), _
            chartWidth, _
            chartHeight)
        With chartObj.Chart
            .ChartType = xlXYScatter
            .HasTitle = True
            .ChartTitle.Text = "Anscombe Set " & CStr(i)
            .SeriesCollection.NewSeries
            With .SeriesCollection(1)
                .Name = "Set " & CStr(i)
                .XValues = xRange
                .Values = yRange
            End With
            .HasLegend = False
            With .Axes(xlCategory)
                .MinimumScale = 0
                .MaximumScale = 20
                .MajorUnit = 5
            End With
            With .Axes(xlValue)
                .MinimumScale = 0
                .MaximumScale = 14
                .MajorUnit = 2
            End With
        End With
    Next i
    Exit Sub
ErrHandler:
    MsgBox "Anscombe chart creation failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

'/**
' * Description: Returns a dense rank for a number within a reference range.
' * Parameters: NumberValue - Number to rank; RefRange - Reference range; Order - Optional order.
' * Returns: Dense rank.
' */
Public Function BI_Rank_Dense(ByVal NumberValue As Double, ByVal RefRange As Variant, Optional ByVal Order As Long = 0) As Variant
    On Error GoTo ErrHandler
    Dim values() As Double
    Dim uniques As Object
    Dim i As Long
    Dim sorted() As Double
    Dim n As Long, j As Long, temp As Double

    values = BI_ToDoubleVector(RefRange, False)
    Set uniques = CreateObject("Scripting.Dictionary")
    For i = LBound(values) To UBound(values)
        uniques(CStr(values(i))) = values(i)
    Next i
    ReDim sorted(1 To uniques.Count)
    For Each temp In uniques.Items
        n = n + 1
        sorted(n) = temp
    Next temp
    For i = 1 To n - 1
        For j = i + 1 To n
            If (Order = 0 And sorted(j) > sorted(i)) Or (Order <> 0 And sorted(j) < sorted(i)) Then
                temp = sorted(i): sorted(i) = sorted(j): sorted(j) = temp
            End If
        Next j
    Next i
    For i = 1 To n
        If sorted(i) = NumberValue Then
            BI_Rank_Dense = i
            Exit Function
        End If
    Next i
    BI_Rank_Dense = CVErr(xlErrNA)
    Exit Function
ErrHandler:
    BI_Rank_Dense = CVErr(xlErrValue)
End Function

'/**
' * Description: Generates repeatable pseudo-random numbers from a seed.
' * Parameters: Seed - Seed value; Size - Number of rows.
' * Returns: Single-column random number range.
' */
Public Function BI_GenerateRandom(ByVal Seed As Long, ByVal Size As Long) As Variant
    Dim result() As Variant
    Dim value As Double
    Dim i As Long
    ReDim result(1 To Application.Max(1, Size), 1 To 1)
    value = Seed / (2 ^ 31 - 1)
    For i = 1 To Application.Max(1, Size)
        value = ((value * (2 ^ 31 - 1) * 48271) Mod (2 ^ 31 - 1)) / (2 ^ 31 - 1)
        result(i, 1) = value
    Next i
    BI_GenerateRandom = result
End Function

'/**
' * Description: Generates repeatable pseudo-random numbers between two limits.
' * Parameters: Seed - Seed value; Size - Number of rows; Bottom - Minimum; Top - Maximum; IntegerOnly - Optional integer flag.
' * Returns: Single-column random number range.
' */
Public Function BI_GenerateRandomBetween(ByVal Seed As Long, ByVal Size As Long, ByVal Bottom As Double, ByVal Top As Double, Optional ByVal IntegerOnly As Boolean = False) As Variant
    Dim base As Variant
    Dim result() As Variant
    Dim i As Long
    base = BI_GenerateRandom(Seed, Size)
    ReDim result(1 To UBound(base, 1), 1 To 1)
    For i = 1 To UBound(base, 1)
        result(i, 1) = Bottom + (CDbl(base(i, 1)) * (Top - Bottom))
        If IntegerOnly Then result(i, 1) = Int(result(i, 1))
    Next i
    BI_GenerateRandomBetween = result
End Function

'/**
' * Description: Returns histogram bins and counts for a numeric range.
' * Parameters: SourceRange - Numeric source range; BinStart - Optional starting bin; BinIncrement - Optional bin increment.
' * Returns: Two-column histogram table.
' */
Public Function BI_Histogram(ByVal SourceRange As Variant, Optional ByVal BinStart As Variant, Optional ByVal BinIncrement As Variant) As Variant
    On Error GoTo ErrHandler
    Dim values() As Double
    Dim minValue As Double, maxValue As Double, startValue As Double, incrementValue As Double
    Dim binRows As Long, i As Long, idx As Long
    Dim result() As Variant

    values = BI_ToDoubleVector(SourceRange, False)
    minValue = Application.WorksheetFunction.Min(values)
    maxValue = Application.WorksheetFunction.Max(values)
    If IsMissing(BinStart) Or IsEmpty(BinStart) Then startValue = minValue Else startValue = CDbl(BinStart)
    If IsMissing(BinIncrement) Or IsEmpty(BinIncrement) Then
        incrementValue = Application.WorksheetFunction.RoundUp((maxValue - minValue) / 10, 0)
    Else
        incrementValue = CDbl(BinIncrement)
    End If
    If incrementValue = 0 Then incrementValue = 1
    binRows = Application.WorksheetFunction.RoundUp((maxValue - startValue) / incrementValue, 0) + 1
    ReDim result(1 To binRows, 1 To 2)
    For i = 1 To binRows
        result(i, 1) = startValue + ((i - 1) * incrementValue) & "-" & startValue + (i * incrementValue)
        result(i, 2) = 0
    Next i
    For i = LBound(values) To UBound(values)
        idx = Int((values(i) - startValue) / incrementValue) + 1
        If idx < 1 Then idx = 1
        If idx > binRows Then idx = binRows
        result(idx, 2) = result(idx, 2) + 1
    Next i
    BI_Histogram = result
    Exit Function
ErrHandler:
    BI_Histogram = CVErr(xlErrValue)
End Function

'/**
' * Description: Returns moving average values for a numeric range.
' * Parameters: SourceRange - Numeric source range; Interval - Optional window size; Labels - Optional header flag; StandardError - Optional standard error output.
' * Returns: Moving average output range.
' */
Public Function BI_MovingAverage(ByVal SourceRange As Variant, Optional ByVal Interval As Long = 3, Optional ByVal Labels As Boolean = False, Optional ByVal StandardError As Boolean = False) As Variant
    On Error GoTo ErrHandler
    Dim data As Variant
    Dim startRow As Long, rowCount As Long
    Dim result() As Variant
    Dim i As Long, j As Long
    Dim avg As Double, sumSq As Double

    data = BI_ToColumnVector(SourceRange)
    startRow = IIf(Labels, 2, 1)
    rowCount = UBound(data, 1) - startRow + 1
    ReDim result(1 To rowCount, 1 To IIf(StandardError, 2, 1))
    For i = 1 To rowCount
        If i < Interval Then
            result(i, 1) = CVErr(xlErrNA)
            If StandardError Then result(i, 2) = CVErr(xlErrNA)
        Else
            avg = 0
            For j = i - Interval + 1 To i
                avg = avg + CDbl(data(startRow + j - 1, 1))
            Next j
            avg = avg / Interval
            result(i, 1) = avg
            If StandardError Then
                sumSq = 0
                For j = i - Interval + 1 To i
                    sumSq = sumSq + ((CDbl(data(startRow + j - 1, 1)) - avg) ^ 2)
                Next j
                result(i, 2) = Sqr(sumSq / Interval)
            End If
        End If
    Next i
    BI_MovingAverage = result
    Exit Function
ErrHandler:
    BI_MovingAverage = CVErr(xlErrValue)
End Function

'/**
' * Description: Returns exponential smoothing values for a numeric range.
' * Parameters: SourceRange - Numeric source range; DampingFactor - Optional damping factor; Labels - Optional header flag.
' * Returns: Smoothed values.
' */
Public Function BI_ExponentialSmoothing(ByVal SourceRange As Variant, Optional ByVal DampingFactor As Double = 0.3, Optional ByVal Labels As Boolean = False) As Variant
    On Error GoTo ErrHandler
    Dim data As Variant
    Dim startRow As Long, rowCount As Long
    Dim result() As Variant
    Dim i As Long
    Dim smoothed As Double

    data = BI_ToColumnVector(SourceRange)
    startRow = IIf(Labels, 2, 1)
    rowCount = UBound(data, 1) - startRow + 1
    ReDim result(1 To rowCount, 1 To 1)
    smoothed = CDbl(data(startRow, 1))
    For i = 1 To rowCount
        smoothed = (smoothed * DampingFactor) + (CDbl(data(startRow + i - 1, 1)) * (1 - DampingFactor))
        result(i, 1) = smoothed
    Next i
    BI_ExponentialSmoothing = result
    Exit Function
ErrHandler:
    BI_ExponentialSmoothing = CVErr(xlErrValue)
End Function

'/**
' * Description: Returns a correlation table of numeric columns along with strongest pairwise relationships.
' * Parameters: SourceRange - Numeric source range; Labels - Optional header flag.
' * Returns: Correlation matrix followed by a ranked pair summary.
' */
Public Function BI_Correlation(ByVal SourceRange As Variant, Optional ByVal Labels As Boolean = False) As Variant
    On Error GoTo ErrHandler
    Dim body As Variant, headers As Variant
    Dim columnCount As Long, i As Long, j As Long
    Dim matrix() As Variant, result() As Variant
    Dim pairCount As Long, summaryRow As Long, totalRows As Long
    Dim summary() As Variant
    Dim colA() As Double, colB() As Double

    body = BI_StatsBody(SourceRange, Labels)
    headers = BI_StatsHeaders(SourceRange, Labels)
    columnCount = UBound(body, 2)

    ReDim matrix(1 To columnCount + 1, 1 To columnCount + 1)
    matrix(1, 1) = "Correlation"
    For i = 1 To columnCount
        matrix(1, i + 1) = headers(1, i)
        matrix(i + 1, 1) = headers(1, i)
    Next i

    If columnCount > 1 Then ReDim summary(1 To (columnCount * (columnCount - 1)) \ 2, 1 To 3)
    For i = 1 To columnCount
        colA = BI_StatsColumnToVector(body, i)
        For j = 1 To columnCount
            If i >= j Then
                colB = BI_StatsColumnToVector(body, j)
                matrix(i + 1, j + 1) = BI_CorrelationValue(colA, colB)
                If i > j Then
                    pairCount = pairCount + 1
                    summary(pairCount, 1) = headers(1, i)
                    summary(pairCount, 2) = headers(1, j)
                    summary(pairCount, 3) = matrix(i + 1, j + 1)
                End If
            Else
                matrix(i + 1, j + 1) = vbNullString
            End If
        Next j
    Next i

    If pairCount > 1 Then BI_SortSummaryRows summary, 3, True
    totalRows = UBound(matrix, 1) + IIf(pairCount > 0, pairCount + 3, 0)
    ReDim result(1 To totalRows, 1 To columnCount + 1)
    For i = 1 To UBound(matrix, 1)
        For j = 1 To UBound(matrix, 2)
            result(i, j) = matrix(i, j)
        Next j
    Next i
    If pairCount > 0 Then
        summaryRow = columnCount + 3
        result(summaryRow, 1) = "Strongest Relationships"
        result(summaryRow + 1, 1) = "Column A"
        result(summaryRow + 1, 2) = "Column B"
        result(summaryRow + 1, 3) = "Correlation"
        For i = 1 To pairCount
            result(summaryRow + 1 + i, 1) = summary(i, 1)
            result(summaryRow + 1 + i, 2) = summary(i, 2)
            result(summaryRow + 1 + i, 3) = summary(i, 3)
        Next i
    Else
        result = matrix
    End If
    BI_Correlation = result
    Exit Function
ErrHandler:
    BI_Correlation = CVErr(xlErrValue)
End Function

'/**
' * Description: Returns a covariance table for numeric columns.
' * Parameters: SourceRange - Numeric source range; Labels - Optional header flag.
' * Returns: Covariance matrix.
' */
Public Function BI_Covariance(ByVal SourceRange As Variant, Optional ByVal Labels As Boolean = False) As Variant
    On Error GoTo ErrHandler
    Dim body As Variant, headers As Variant
    Dim columnCount As Long, i As Long, j As Long
    Dim result() As Variant
    Dim colA() As Double, colB() As Double

    body = BI_StatsBody(SourceRange, Labels)
    headers = BI_StatsHeaders(SourceRange, Labels)
    columnCount = UBound(body, 2)
    ReDim result(1 To columnCount + 1, 1 To columnCount + 1)
    result(1, 1) = "Covariance"
    For i = 1 To columnCount
        result(1, i + 1) = headers(1, i)
        result(i + 1, 1) = headers(1, i)
    Next i

    For i = 1 To columnCount
        colA = BI_StatsColumnToVector(body, i)
        For j = 1 To columnCount
            If i >= j Then
                colB = BI_StatsColumnToVector(body, j)
                result(i + 1, j + 1) = BI_PopulationCovariance(colA, colB)
            Else
                result(i + 1, j + 1) = vbNullString
            End If
        Next j
    Next i
    BI_Covariance = result
    Exit Function
ErrHandler:
    BI_Covariance = CVErr(xlErrValue)
End Function

'/**
' * Description: Calculates linear regression statistics, coefficients, and residuals.
' * Parameters: KnownYs - Dependent variable range; KnownXs - Independent variable range; Labels - Optional header flag; Confidence - Optional confidence level.
' * Returns: Regression summary output range.
' */
Public Function BI_Regression(ByVal KnownYs As Variant, ByVal KnownXs As Variant, Optional ByVal Labels As Boolean = False, Optional ByVal Confidence As Double = 0.95) As Variant
    On Error GoTo ErrHandler
    Dim xBody As Variant, xHeaders As Variant, yBody As Variant
    Dim lineEst As Variant
    Dim xCount As Long, obsCount As Long
    Dim coeff() As Double, stdErr() As Double, pred() As Double, resid() As Double
    Dim r As Long, c As Long, outRow As Long, totalRows As Long
    Dim coeffCount As Long
    Dim result() As Variant
    Dim labelText As String
    Dim multipleR As Double, rSquare As Double, adjustedRSq As Double
    Dim standardError As Double, regDf As Double, resDf As Double, totalDf As Double
    Dim regSs As Double, resSs As Double, totalSs As Double, regMs As Double, resMs As Double
    Dim fValue As Double, sigF As Double, tValue As Double, pValue As Double
    Dim confT As Double, lowerBound As Double, upperBound As Double

    xBody = BI_StatsBody(KnownXs, Labels)
    xHeaders = BI_StatsHeaders(KnownXs, Labels, "X")
    yBody = BI_StatsBody(KnownYs, Labels)
    xCount = UBound(xBody, 2)
    obsCount = UBound(yBody, 1)

    lineEst = Application.LinEst(yBody, xBody, True, True)
    coeffCount = xCount + 1
    ReDim coeff(1 To coeffCount)
    ReDim stdErr(1 To coeffCount)
    For c = 1 To xCount
        coeff(c) = CDbl(lineEst(1, xCount - c + 1))
        stdErr(c) = CDbl(lineEst(2, xCount - c + 1))
    Next c
    coeff(coeffCount) = CDbl(lineEst(1, coeffCount))
    stdErr(coeffCount) = CDbl(lineEst(2, coeffCount))

    multipleR = Sqr(CDbl(lineEst(3, 1)))
    rSquare = CDbl(lineEst(3, 1))
    standardError = CDbl(lineEst(3, 2))
    regDf = xCount
    resDf = CDbl(lineEst(4, 2))
    totalDf = regDf + resDf
    regSs = CDbl(lineEst(5, 1))
    resSs = CDbl(lineEst(5, 2))
    totalSs = regSs + resSs
    regMs = regSs / regDf
    resMs = resSs / resDf
    fValue = CDbl(lineEst(4, 1))
    sigF = Application.WorksheetFunction.FDist(fValue, regDf, resDf)
    adjustedRSq = 1 - ((1 - rSquare) * (obsCount - 1) / (obsCount - xCount - 1))
    confT = Application.WorksheetFunction.TInv(1 - Confidence, resDf)
    labelText = BI_ConfidenceCaption(Confidence)

    ReDim pred(1 To obsCount)
    ReDim resid(1 To obsCount)
    For r = 1 To obsCount
        pred(r) = coeff(coeffCount)
        For c = 1 To xCount
            pred(r) = pred(r) + (CDbl(xBody(r, c)) * coeff(c))
        Next c
        resid(r) = CDbl(yBody(r, 1)) - pred(r)
    Next r

    totalRows = 18 + coeffCount + obsCount
    ReDim result(1 To totalRows, 1 To 9)
    result(1, 1) = "Summary Output"
    result(3, 1) = "Regression Statistics"
    result(4, 1) = "Multiple R": result(4, 2) = multipleR
    result(5, 1) = "R Square": result(5, 2) = rSquare
    result(6, 1) = "Adjusted R Square": result(6, 2) = adjustedRSq
    result(7, 1) = "Standard Error": result(7, 2) = standardError
    result(8, 1) = "Observations": result(8, 2) = obsCount

    result(10, 1) = "ANOVA"
    result(11, 2) = "df": result(11, 3) = "SS": result(11, 4) = "MS": result(11, 5) = "F": result(11, 6) = "Significance F"
    result(12, 1) = "Regression": result(12, 2) = regDf: result(12, 3) = regSs: result(12, 4) = regMs: result(12, 5) = fValue: result(12, 6) = sigF
    result(13, 1) = "Residual": result(13, 2) = resDf: result(13, 3) = resSs: result(13, 4) = resMs
    result(14, 1) = "Total": result(14, 2) = totalDf: result(14, 3) = totalSs

    outRow = 16
    result(outRow, 2) = "Coefficients"
    result(outRow, 3) = "Standard Error"
    result(outRow, 4) = "t Stat"
    result(outRow, 5) = "P-Value"
    result(outRow, 6) = "Lower " & labelText
    result(outRow, 7) = "Upper " & labelText

    For c = 1 To coeffCount
        If c <= xCount Then
            result(outRow + c, 1) = xHeaders(1, c)
        Else
            result(outRow + c, 1) = "Intercept"
        End If
        result(outRow + c, 2) = coeff(c)
        result(outRow + c, 3) = stdErr(c)
        If stdErr(c) <> 0 Then
            tValue = coeff(c) / stdErr(c)
            pValue = Application.WorksheetFunction.TDist(Abs(tValue), resDf, 2)
            lowerBound = coeff(c) - (stdErr(c) * confT)
            upperBound = coeff(c) + (stdErr(c) * confT)
            result(outRow + c, 4) = tValue
            result(outRow + c, 5) = pValue
            result(outRow + c, 6) = lowerBound
            result(outRow + c, 7) = upperBound
        End If
    Next c

    outRow = outRow + coeffCount + 2
    result(outRow, 1) = "Observation"
    result(outRow, 2) = "Predicted Y"
    result(outRow, 3) = "Residuals"
    For r = 1 To obsCount
        result(outRow + r, 1) = r
        result(outRow + r, 2) = pred(r)
        result(outRow + r, 3) = resid(r)
    Next r

    BI_Regression = result
    Exit Function
ErrHandler:
    BI_Regression = CVErr(xlErrValue)
End Function

Private Sub BI_ShowStatsTool(ByVal toolName As String)
    Unload frmBI_StatsTool
    frmBI_StatsTool.Tag = toolName
    frmBI_StatsTool.Show
End Sub

Public Sub BI_Tool_Describe(control As IRibbonControl)
    On Error GoTo ErrHandler
    BI_ShowStatsTool "Describe"
    Exit Sub
ErrHandler:
    MsgBox "Describe failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_Correlation(control As IRibbonControl)
    On Error GoTo ErrHandler
    BI_ShowStatsTool "Correlation"
    Exit Sub
ErrHandler:
    MsgBox "Correlation failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_Covariance(control As IRibbonControl)
    On Error GoTo ErrHandler
    BI_ShowStatsTool "Covariance"
    Exit Sub
ErrHandler:
    MsgBox "Covariance failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_Regression(control As IRibbonControl)
    On Error GoTo ErrHandler
    BI_ShowStatsTool "Regression"
    Exit Sub
ErrHandler:
    MsgBox "Regression failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_Histogram(control As IRibbonControl)
    On Error GoTo ErrHandler
    BI_ShowStatsTool "Histogram"
    Exit Sub
ErrHandler:
    MsgBox "Histogram failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_MovingAverage(control As IRibbonControl)
    On Error GoTo ErrHandler
    BI_ShowStatsTool "MovingAverage"
    Exit Sub
ErrHandler:
    MsgBox "Moving Average failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_ExponentialSmoothing(control As IRibbonControl)
    On Error GoTo ErrHandler
    BI_ShowStatsTool "ExponentialSmoothing"
    Exit Sub
ErrHandler:
    MsgBox "Exponential Smoothing failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_AnovaSingleFactor(control As IRibbonControl)
    On Error GoTo ErrHandler
    BI_ShowStatsTool "AnovaSingle"
    Exit Sub
ErrHandler:
    MsgBox "Anova Single Factor failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_AnovaTwoFactorsWithoutReplication(control As IRibbonControl)
    On Error GoTo ErrHandler
    BI_ShowStatsTool "AnovaTwoNoReplication"
    Exit Sub
ErrHandler:
    MsgBox "Anova Two Factors failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_AnsCombeQuartet(control As IRibbonControl)
    On Error GoTo ErrHandler
    BI_ShowStatsTool "AnsCombeQuartet"
    Exit Sub
ErrHandler:
    MsgBox "AnsCombe Quartet failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub
