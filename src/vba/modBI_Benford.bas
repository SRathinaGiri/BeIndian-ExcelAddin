Attribute VB_Name = "modBI_Benford"
Option Explicit

Private Function BI_FirstDigits(ByVal value As Double, ByVal digitCount As Long) As Long
    Dim s As String
    value = Abs(value)
    If value = 0 Then Exit Function
    s = Replace$(CStr(value), ".", vbNullString)
    Do While Left$(s, 1) = "0" And Len(s) > 1
        s = Mid$(s, 2)
    Loop
    If Len(s) < digitCount Then Exit Function
    BI_FirstDigits = CLng(Left$(s, digitCount))
End Function

Private Function BI_NthDigit(ByVal value As Double, ByVal digitPosition As Long) As Long
    Dim s As String
    value = Abs(value)
    If value = 0 Then BI_NthDigit = -1: Exit Function
    s = Replace$(CStr(value), ".", vbNullString)
    Do While Left$(s, 1) = "0" And Len(s) > 1
        s = Mid$(s, 2)
    Loop
    If Len(s) < digitPosition Then BI_NthDigit = -1 Else BI_NthDigit = CLng(Mid$(s, digitPosition, 1))
End Function

Private Function BI_BenfordDigitTable(ByVal BenfordRange As Variant, ByVal title As String, ByVal firstDigit As Long, ByVal lastDigit As Long, ByVal digitCount As Long, ByVal mode As String) As Variant
    Dim numbers() As Double
    Dim result() As Variant
    Dim d As Long, i As Long, rowIndex As Long, total As Long
    Dim digitValue As Long, freq As Long
    Dim actual As Double, expected As Double, absDevTotal As Double

    numbers = BI_ToDoubleVector(BenfordRange, True)
    ReDim result(1 To (lastDigit - firstDigit + 1) + 4, 1 To 4)
    result(1, 1) = title
    result(2, 1) = "Digit": result(2, 2) = "Frequency": result(2, 3) = "Actual": result(2, 4) = "Expected"

    For d = firstDigit To lastDigit
        rowIndex = d - firstDigit + 3
        freq = 0
        For i = LBound(numbers) To UBound(numbers)
            Select Case mode
                Case "FIRST": digitValue = BI_FirstDigits(numbers(i), digitCount)
                Case "NTH": digitValue = BI_NthDigit(numbers(i), digitCount)
                Case "LAST2": digitValue = CLng(Abs(numbers(i))) Mod 100
            End Select
            If digitValue = d Then freq = freq + 1
        Next i
        total = total + freq
        result(rowIndex, 1) = d
        result(rowIndex, 2) = freq
    Next d

    For d = firstDigit To lastDigit
        rowIndex = d - firstDigit + 3
        If total = 0 Then actual = 0 Else actual = CDbl(result(rowIndex, 2)) / total
        Select Case mode
            Case "FIRST": expected = Log((d + 1) / d) / Log(10#)
            Case "NTH"
                expected = BI_BenfordNthDigitExpected(d, digitCount)
            Case "LAST2"
                expected = 0.01
        End Select
        result(rowIndex, 3) = actual
        result(rowIndex, 4) = expected
        absDevTotal = absDevTotal + Abs(actual - expected)
    Next d

    rowIndex = UBound(result, 1) - 1
    result(rowIndex, 1) = "Total": result(rowIndex, 2) = total: result(rowIndex, 3) = 1: result(rowIndex, 4) = 1
    result(rowIndex + 1, 1) = "Mean Absolute Deviation"
    result(rowIndex + 1, 4) = absDevTotal / (lastDigit - firstDigit + 1)
    BI_BenfordDigitTable = result
End Function

Private Function BI_BenfordNthDigitExpected(ByVal digitValue As Long, ByVal digitPosition As Long) As Double
    Dim n As Long
    Select Case digitPosition
        Case 2
            For n = 1 To 9
                BI_BenfordNthDigitExpected = BI_BenfordNthDigitExpected + Log(1 + 1 / ((n * 10) + digitValue)) / Log(10#)
            Next n
        Case 3
            For n = 10 To 99
                BI_BenfordNthDigitExpected = BI_BenfordNthDigitExpected + Log(1 + 1 / ((n * 10) + digitValue)) / Log(10#)
            Next n
        Case Else
            BI_BenfordNthDigitExpected = 1# / 10#
    End Select
End Function

'/**
' * Description: Calculates Benford Law First Digit frequency, actual and expected values.
' * Parameters: BenfordRange - Column of numbers.
' * Returns: Variant array with frequency table and mean absolute deviation.
' */
Public Function BI_BenfordLaw_FirstDigit(ByVal BenfordRange As Variant) As Variant
    BI_BenfordLaw_FirstDigit = BI_BenfordDigitTable(BenfordRange, "Benford Law First Digit", 1, 9, 1, "FIRST")
End Function

Public Function BI_BenfordLaw_FirstTwoDigits(ByVal BenfordRange As Variant) As Variant
    BI_BenfordLaw_FirstTwoDigits = BI_BenfordDigitTable(BenfordRange, "Benford Law First Two Digits", 10, 99, 2, "FIRST")
End Function

Public Function BI_BenfordLaw_FirstThreeDigits(ByVal BenfordRange As Variant) As Variant
    BI_BenfordLaw_FirstThreeDigits = BI_BenfordDigitTable(BenfordRange, "Benford Law First Three Digits", 100, 999, 3, "FIRST")
End Function

Public Function BI_BenfordLaw_SecondDigit(ByVal BenfordRange As Variant) As Variant
    BI_BenfordLaw_SecondDigit = BI_BenfordDigitTable(BenfordRange, "Benford Law Second Digit", 0, 9, 2, "NTH")
End Function

Public Function BI_BenfordLaw_ThirdDigit(ByVal BenfordRange As Variant) As Variant
    BI_BenfordLaw_ThirdDigit = BI_BenfordDigitTable(BenfordRange, "Benford Law Third Digit", 0, 9, 3, "NTH")
End Function

Public Function BI_BenfordLaw_LastTwoDigits(ByVal BenfordRange As Variant) As Variant
    BI_BenfordLaw_LastTwoDigits = BI_BenfordDigitTable(BenfordRange, "Benford Law Last Two Digits", 1, 99, 2, "LAST2")
End Function

Public Function BI_BenfordLaw_SecondOrder(ByVal BenfordRange As Variant) As Variant
    On Error GoTo ErrHandler
    Dim numbers() As Double
    Dim i As Long, j As Long, n As Long
    Dim temp As Double
    Dim result(1 To 94, 1 To 4) As Variant
    Dim d As Long, rowIndex As Long, prefix As Long
    Dim total As Long, actual As Double, expected As Double, absDevTotal As Double

    numbers = BI_ToDoubleVector(BenfordRange, True)
    If UBound(numbers) <= LBound(numbers) Then
        BI_BenfordLaw_SecondOrder = CVErr(xlErrNA)
        Exit Function
    End If

    For i = LBound(numbers) To UBound(numbers) - 1
        For j = i + 1 To UBound(numbers)
            If numbers(j) < numbers(i) Then
                temp = numbers(i)
                numbers(i) = numbers(j)
                numbers(j) = temp
            End If
        Next j
    Next i

    result(1, 1) = "Benford Law Second Order"
    result(2, 1) = "Digit": result(2, 2) = "Frequency": result(2, 3) = "Actual": result(2, 4) = "Expected"

    For d = 10 To 99
        result(d - 7, 1) = d
    Next d

    For i = LBound(numbers) + 1 To UBound(numbers)
        prefix = BI_FirstDigits(Round((numbers(i) - numbers(i - 1)) * 100000#, 0), 2)
        If prefix >= 10 And prefix <= 99 Then
            rowIndex = prefix - 7
            result(rowIndex, 2) = CLng(result(rowIndex, 2)) + 1
            total = total + 1
        End If
    Next i

    For d = 10 To 99
        rowIndex = d - 7
        If total = 0 Then actual = 0 Else actual = CDbl(result(rowIndex, 2)) / total
        expected = Log(1 + 1 / d) / Log(10#)
        result(rowIndex, 3) = actual
        result(rowIndex, 4) = expected
        absDevTotal = absDevTotal + Abs(actual - expected)
    Next d

    result(93, 1) = "Total": result(93, 2) = total: result(93, 3) = 1: result(93, 4) = 1
    result(94, 1) = "Mean Absolute Deviation": result(94, 4) = absDevTotal / 90#
    BI_BenfordLaw_SecondOrder = result
    Exit Function
ErrHandler:
    BI_BenfordLaw_SecondOrder = CVErr(xlErrValue)
End Function

Public Function BI_BenfordLaw_SummaryTest(ByVal BenfordRange As Variant) As Variant
    On Error GoTo ErrHandler
    Dim numbers() As Double
    Dim result(1 To 94, 1 To 4) As Variant
    Dim d As Long, i As Long, rowIndex As Long
    Dim prefix As Long, total As Double, actual As Double, absDevTotal As Double

    numbers = BI_ToDoubleVector(BenfordRange, True)
    result(1, 1) = "Benford Law Summary Test"
    result(2, 1) = "Digit": result(2, 2) = "Frequency": result(2, 3) = "Actual": result(2, 4) = "Expected"

    For d = 10 To 99
        rowIndex = d - 7
        result(rowIndex, 1) = d
        For i = LBound(numbers) To UBound(numbers)
            prefix = BI_FirstDigits(numbers(i), 2)
            If prefix = d Then
                result(rowIndex, 2) = CDbl(result(rowIndex, 2)) + numbers(i)
            End If
        Next i
        total = total + CDbl(result(rowIndex, 2))
    Next d

    For d = 10 To 99
        rowIndex = d - 7
        If total = 0 Then actual = 0 Else actual = CDbl(result(rowIndex, 2)) / total
        result(rowIndex, 3) = actual
        result(rowIndex, 4) = 1# / 90#
        absDevTotal = absDevTotal + Abs(actual - (1# / 90#))
    Next d

    result(93, 1) = "Total": result(93, 2) = total: result(93, 3) = 1: result(93, 4) = 1
    result(94, 1) = "Mean Absolute Deviation": result(94, 4) = absDevTotal / 90#
    BI_BenfordLaw_SummaryTest = result
    Exit Function
ErrHandler:
    BI_BenfordLaw_SummaryTest = CVErr(xlErrValue)
End Function

'/**
' * Description: Creates a clustered column chart comparing Benford Actual and Expected percentages.
' * Parameters: TableTopLeft - Top-left cell of a Benford output table; ChartTitle - Chart title.
' * Returns: Nothing. A chart is inserted near the output table.
' */
Public Sub BI_CreateBenfordActualExpectedChart(ByVal TableTopLeft As Range, Optional ByVal ChartTitle As String = "Benford Law - Actual vs Expected")
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Dim firstDataRow As Long, lastDataRow As Long
    Dim currentRow As Long
    Dim chartLeft As Double, chartTop As Double
    Dim chartObject As ChartObject
    Dim chartName As String

    Set ws = TableTopLeft.Worksheet
    firstDataRow = TableTopLeft.Row + 2
    currentRow = firstDataRow
    Do While Len(CStr(ws.Cells(currentRow, TableTopLeft.Column).Value)) > 0
        If CStr(ws.Cells(currentRow, TableTopLeft.Column).Value) = "Total" Then Exit Do
        If CStr(ws.Cells(currentRow, TableTopLeft.Column).Value) = "Mean Absolute Deviation" Then Exit Do
        currentRow = currentRow + 1
    Loop
    lastDataRow = currentRow - 1
    If lastDataRow < firstDataRow Then Exit Sub

    chartName = "BI_BenfordChart_" & Format(Now, "hhmmss")
    chartLeft = TableTopLeft.Offset(0, 5).Left
    chartTop = TableTopLeft.Top
    Set chartObject = ws.ChartObjects.Add(chartLeft, chartTop, 480, 300)
    chartObject.Name = chartName

    With chartObject.Chart
        .ChartType = xlColumnClustered
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        With .SeriesCollection.NewSeries
            .Name = "Actual"
            .XValues = ws.Range(ws.Cells(firstDataRow, TableTopLeft.Column), ws.Cells(lastDataRow, TableTopLeft.Column))
            .Values = ws.Range(ws.Cells(firstDataRow, TableTopLeft.Column + 2), ws.Cells(lastDataRow, TableTopLeft.Column + 2))
        End With
        With .SeriesCollection.NewSeries
            .Name = "Expected"
            .XValues = ws.Range(ws.Cells(firstDataRow, TableTopLeft.Column), ws.Cells(lastDataRow, TableTopLeft.Column))
            .Values = ws.Range(ws.Cells(firstDataRow, TableTopLeft.Column + 3), ws.Cells(lastDataRow, TableTopLeft.Column + 3))
        End With
        .HasTitle = True
        .ChartTitle.Text = ChartTitle
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Axes(xlValue).TickLabels.NumberFormat = "0.00%"
        .Axes(xlCategory).HasTitle = False
        .Axes(xlValue).HasMajorGridlines = True
    End With
    Exit Sub
ErrHandler:
    MsgBox "Benford chart creation failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_BenfordFirstDigit(control As IRibbonControl)
    On Error GoTo ErrHandler
    frmBI_Benford.Show
    Exit Sub
ErrHandler:
    BI_Tool_BenfordFirstDigitFallback
End Sub

Private Sub BI_Tool_BenfordFirstDigitFallback()
    On Error GoTo ErrHandler
    Dim target As Range
    Set target = Application.InputBox("Select a column of numbers for Benford first digit analysis.", "BeIndian Benford", Type:=8)
    If target Is Nothing Then Exit Sub
    BI_OutputToActiveCell BI_BenfordLaw_FirstDigit(target)
    BI_CreateBenfordActualExpectedChart ActiveCell, "Benford Law First Digit - Actual vs Expected"
    Exit Sub
ErrHandler:
    MsgBox "Benford analysis failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub
