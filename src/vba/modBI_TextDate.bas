Attribute VB_Name = "modBI_TextDate"
Option Explicit

'/**
' * Description: Returns whether one text begins with another text.
' * Parameters: FindText - Prefix to find; WithinText - Text to test; CaseSensitive - Optional case flag.
' * Returns: TRUE if WithinText begins with FindText.
' */
Public Function BI_BeginsWith(ByVal FindText As Variant, ByVal WithinText As Variant, Optional ByVal CaseSensitive As Boolean = False) As Boolean
    Dim findValue As String, withinValue As String
    findValue = BI_NormalizeText(FindText, CaseSensitive)
    withinValue = BI_NormalizeText(WithinText, CaseSensitive)
    BI_BeginsWith = (Left$(withinValue, Len(findValue)) = findValue)
End Function

'/**
' * Description: Returns TRUE if FindText is found in WithinText.
' * Parameters: FindText - Text to find; WithinText - Text to search; CaseSensitive - Optional case flag.
' * Returns: TRUE if found.
' */
Public Function BI_Contains(ByVal FindText As Variant, ByVal WithinText As Variant, Optional ByVal CaseSensitive As Boolean = False) As Boolean
    Dim compareMode As VbCompareMethod
    compareMode = IIf(CaseSensitive, vbBinaryCompare, vbTextCompare)
    BI_Contains = (InStr(1, CStr(WithinText), CStr(FindText), compareMode) > 0)
End Function

'/**
' * Description: Returns whether one text ends with another text.
' * Parameters: FindText - Suffix to find; WithinText - Text to test; CaseSensitive - Optional case flag.
' * Returns: TRUE if WithinText ends with FindText.
' */
Public Function BI_EndsWith(ByVal FindText As Variant, ByVal WithinText As Variant, Optional ByVal CaseSensitive As Boolean = False) As Boolean
    Dim findValue As String, withinValue As String
    findValue = BI_NormalizeText(FindText, CaseSensitive)
    withinValue = BI_NormalizeText(WithinText, CaseSensitive)
    BI_EndsWith = (Right$(withinValue, Len(findValue)) = findValue)
End Function

'/**
' * Description: Extracts all numeric characters from text.
' * Parameters: Text - Input text.
' * Returns: Text containing only digits.
' */
Public Function BI_ExtractNumbers(ByVal Text As Variant) As String
    Dim i As Long, ch As String
    For i = 1 To Len(CStr(Text))
        ch = Mid$(CStr(Text), i, 1)
        If ch >= "0" And ch <= "9" Then BI_ExtractNumbers = BI_ExtractNumbers & ch
    Next i
End Function

'/**
' * Description: Extracts non-numeric characters from text.
' * Parameters: Text - Input text.
' * Returns: Text excluding digits.
' */
Public Function BI_ExtractText(ByVal Text As Variant) As String
    Dim i As Long, ch As String
    For i = 1 To Len(CStr(Text))
        ch = Mid$(CStr(Text), i, 1)
        If ch < "0" Or ch > "9" Then BI_ExtractText = BI_ExtractText & ch
    Next i
End Function

'/**
' * Description: Inserts a substring before a specified position.
' * Parameters: Text - Source text; Position - 1-based insert position; InsertText - Text to insert.
' * Returns: Combined text.
' */
Public Function BI_Insert(ByVal Text As Variant, ByVal Position As Long, ByVal InsertText As Variant) As String
    BI_Insert = Left$(CStr(Text), Position - 1) & CStr(InsertText) & Mid$(CStr(Text), Position)
End Function

'/**
' * Description: Returns text in reverse order.
' * Parameters: Text - Input text.
' * Returns: Reversed text.
' */
Public Function BI_ReverseText(ByVal Text As Variant) As String
    Dim i As Long
    For i = Len(CStr(Text)) To 1 Step -1
        BI_ReverseText = BI_ReverseText & Mid$(CStr(Text), i, 1)
    Next i
End Function

'/**
' * Description: Returns the number of words in text.
' * Parameters: Text - Input text.
' * Returns: Word count.
' */
Public Function BI_WordCount(ByVal Text As Variant) As Long
    Dim cleaned As String
    cleaned = Application.WorksheetFunction.Trim(CStr(Text))
    If Len(cleaned) = 0 Then
        BI_WordCount = 0
    Else
        BI_WordCount = UBound(Split(cleaned, " ")) + 1
    End If
End Function

'/**
' * Description: Returns the position of the first character from a pattern found in text.
' * Parameters: Text - Source text; Pattern - Characters to search for.
' * Returns: Position of the first match, or 0 if not found.
' */
Public Function BI_FindOneOf(ByVal Text As Variant, ByVal Pattern As Variant) As Long
    Dim i As Long
    Dim charPos As Long
    Dim minPos As Long
    For i = 1 To Len(CStr(Pattern))
        charPos = InStr(1, CStr(Text), Mid$(CStr(Pattern), i, 1), vbTextCompare)
        If charPos > 0 Then
            If minPos = 0 Or charPos < minPos Then minPos = charPos
        End If
    Next i
    BI_FindOneOf = minPos
End Function

'/**
' * Description: Returns the Levenshtein distance between two texts.
' * Parameters: Text1 - First text; Text2 - Second text.
' * Returns: Count of edits required to transform one text into the other.
' */
Public Function BI_Fuzzy(ByVal Text1 As Variant, ByVal Text2 As Variant) As Long
    Dim a As String, b As String
    Dim d() As Long
    Dim i As Long, j As Long
    Dim cost As Long

    a = CStr(Text1)
    b = CStr(Text2)
    ReDim d(0 To Len(a), 0 To Len(b))

    For i = 0 To Len(a)
        d(i, 0) = i
    Next i
    For j = 0 To Len(b)
        d(0, j) = j
    Next j

    For i = 1 To Len(a)
        For j = 1 To Len(b)
            cost = IIf(Mid$(a, i, 1) = Mid$(b, j, 1), 0, 1)
            d(i, j) = Application.Min(d(i - 1, j) + 1, d(i, j - 1) + 1, d(i - 1, j - 1) + cost)
        Next j
    Next i

    BI_Fuzzy = d(Len(a), Len(b))
End Function

'/**
' * Description: Fills down blank cells using the nearest non-blank value above.
' * Parameters: SourceRange - Source range.
' * Returns: Filled range values.
' */
Public Function BI_FillDown(ByVal SourceRange As Variant) As Variant
    Dim data As Variant
    Dim result() As Variant
    Dim r As Long, c As Long
    Dim lastValue As Variant

    data = BI_SourceTo2D_Local(SourceRange)
    ReDim result(1 To UBound(data, 1), 1 To UBound(data, 2))

    For c = 1 To UBound(data, 2)
        lastValue = vbNullString
        For r = 1 To UBound(data, 1)
            If Len(CStr(data(r, c))) > 0 Then
                lastValue = data(r, c)
                result(r, c) = data(r, c)
            Else
                result(r, c) = lastValue
            End If
        Next r
    Next c
    BI_FillDown = result
End Function

'/**
' * Description: Returns formulas found in a range with their addresses.
' * Parameters: ReferenceRange - Range to inspect.
' * Returns: Single-column list of address and formula text.
' */
Public Function BI_FormulaList(ByVal ReferenceRange As Range) As Variant
    Dim items As Collection
    Dim result() As Variant
    Dim cell As Range
    Dim i As Long

    Set items = New Collection
    For Each cell In ReferenceRange.Cells
        If cell.HasFormula Then items.Add cell.Address(False, False) & " " & cell.Formula
    Next cell

    If items.Count = 0 Then
        BI_FormulaList = CVErr(xlErrNA)
        Exit Function
    End If

    ReDim result(1 To items.Count, 1 To 1)
    For i = 1 To items.Count
        result(i, 1) = items(i)
    Next i
    BI_FormulaList = result
End Function

'/**
' * Description: Splits text at specified positions.
' * Parameters: Text - Source text; Positions - Column or row of split positions.
' * Returns: One-row split result.
' */
Public Function BI_TextSplitByPositions(ByVal Text As Variant, ByVal Positions As Variant) As Variant
    Dim posData As Variant
    Dim result() As Variant
    Dim i As Long
    Dim startPos As Long, endPos As Long

    posData = BI_ToColumnVector(Positions)
    ReDim result(1 To 1, 1 To UBound(posData, 1))
    startPos = 1
    For i = 1 To UBound(posData, 1)
        endPos = CLng(posData(i, 1))
        result(1, i) = Mid$(CStr(Text), startPos, endPos - startPos)
        startPos = endPos
    Next i
    BI_TextSplitByPositions = result
End Function

Private Function BI_SourceTo2D_Local(ByVal source As Variant) As Variant
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
        BI_SourceTo2D_Local = result
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
    BI_SourceTo2D_Local = result
End Function

Private Function BI_TwoDigitWords(ByVal value As Long) As String
    Dim units As Variant, tens As Variant
    units = Array("", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", _
                  "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen")
    tens = Array("", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety")
    If value < 20 Then
        BI_TwoDigitWords = units(value)
    Else
        BI_TwoDigitWords = tens(value \ 10)
        If value Mod 10 > 0 Then BI_TwoDigitWords = BI_TwoDigitWords & " " & units(value Mod 10)
    End If
End Function

Private Function BI_ThreeDigitWords(ByVal value As Long) As String
    Dim words As String
    If value >= 100 Then
        words = BI_TwoDigitWords(value \ 100) & " Hundred"
        If value Mod 100 > 0 Then words = words & " and " & BI_TwoDigitWords(value Mod 100)
    Else
        words = BI_TwoDigitWords(value)
    End If
    BI_ThreeDigitWords = Trim$(words)
End Function

'/**
' * Description: Returns the English spelling of a number.
' * Parameters: NumberValue - Number to convert.
' * Returns: Number in words.
' */
Public Function BI_Num2Word(ByVal NumberValue As Variant) As String
    On Error GoTo ErrHandler
    Dim n As Double
    Dim wholePart As Double
    Dim parts(0 To 4) As Long
    Dim labels As Variant
    Dim words As String
    Dim i As Long

    labels = Array("", "Thousand", "Million", "Billion", "Trillion")
    n = Abs(CDbl(NumberValue))
    wholePart = Fix(n)
    If wholePart = 0 Then
        BI_Num2Word = "Zero"
        Exit Function
    End If

    For i = 0 To 4
        parts(i) = wholePart Mod 1000
        wholePart = Int(wholePart / 1000)
    Next i
    For i = 4 To 0 Step -1
        If parts(i) > 0 Then
            words = words & " " & BI_ThreeDigitWords(parts(i))
            If Len(labels(i)) > 0 Then words = words & " " & labels(i)
        End If
    Next i
    BI_Num2Word = Trim$(words)
    Exit Function
ErrHandler:
    BI_Num2Word = "Overflow!"
End Function

'/**
' * Description: Returns the Indian spelling of a number.
' * Parameters: NumberValue - Number to convert.
' * Returns: Number in Indian words.
' */
Public Function BI_IndianNum2Word(ByVal NumberValue As Variant) As String
    On Error GoTo ErrHandler
    Dim n As Double
    Dim crore As Long, lakh As Long, thousand As Long, hundred As Long, remainder As Long
    Dim words As String

    n = Abs(CDbl(NumberValue))
    If Fix(n) = 0 Then
        BI_IndianNum2Word = "Zero"
        Exit Function
    End If
    If Len(CStr(Fix(n))) > 16 Then
        BI_IndianNum2Word = "Overflow!"
        Exit Function
    End If

    crore = Int(n / 10000000)
    n = n Mod 10000000
    lakh = Int(n / 100000)
    n = n Mod 100000
    thousand = Int(n / 1000)
    n = n Mod 1000
    hundred = Int(n / 100)
    remainder = n Mod 100

    If crore > 0 Then words = words & " " & BI_TwoDigitWords(crore) & IIf(crore = 1, " Crore", " Crores")
    If lakh > 0 Then words = words & " " & BI_TwoDigitWords(lakh) & IIf(lakh = 1, " Lakh", " Lakhs")
    If thousand > 0 Then words = words & " " & BI_TwoDigitWords(thousand) & " Thousand"
    If hundred > 0 Then words = words & " " & BI_TwoDigitWords(hundred) & " Hundred"
    If remainder > 0 Then
        If Len(Trim$(words)) > 0 Then
            words = words & " and " & BI_TwoDigitWords(remainder)
        Else
            words = BI_TwoDigitWords(remainder)
        End If
    End If

    BI_IndianNum2Word = Trim$(words)
    Exit Function
ErrHandler:
    BI_IndianNum2Word = "Overflow!"
End Function

'/**
' * Description: Returns a monthly calendar table.
' * Parameters: MonthValue - Optional month; YearValue - Optional year; FormatText - Optional day format; Vertical - Optional layout flag.
' * Returns: Calendar table.
' */
Public Function BI_MonthlyCalendar(Optional ByVal MonthValue As Variant, Optional ByVal YearValue As Variant, Optional ByVal FormatText As String = "DDD", Optional ByVal Vertical As Boolean = False) As Variant
    Dim m As Long, y As Long
    Dim firstDate As Date, dayCount As Long, startCol As Long
    Dim dayHeaders(1 To 1, 1 To 7) As Variant
    Dim result() As Variant
    Dim r As Long, c As Long, d As Long

    If IsMissing(MonthValue) Or IsEmpty(MonthValue) Then m = Month(Date) Else m = CLng(MonthValue)
    If IsMissing(YearValue) Or IsEmpty(YearValue) Then y = Year(Date) Else y = CLng(YearValue)

    firstDate = DateSerial(y, m, 1)
    dayCount = Day(DateSerial(y, m + 1, 0))
    startCol = Weekday(firstDate, vbSunday)

    For c = 1 To 7
        dayHeaders(1, c) = Format$(DateSerial(1900, 1, c), FormatText)
    Next c

    ReDim result(1 To 7, 1 To 7)
    For c = 1 To 7
        result(1, c) = dayHeaders(1, c)
    Next c
    d = 1
    For r = 2 To 7
        For c = 1 To 7
            If (r = 2 And c < startCol) Or d > dayCount Then
                result(r, c) = vbNullString
            Else
                result(r, c) = d
                d = d + 1
            End If
        Next c
    Next r

    If Vertical Then
        Dim v() As Variant
        ReDim v(1 To 7, 1 To 7)
        For r = 1 To 7
            For c = 1 To 7
                v(r, c) = result(c, r)
            Next c
        Next r
        BI_MonthlyCalendar = v
    Else
        BI_MonthlyCalendar = result
    End If
End Function

'/**
' * Description: Returns a yearly calendar table.
' * Parameters: YearValue - Optional year; FormatText - Optional day format; Vertical - Optional layout flag.
' * Returns: Yearly calendar output.
' */
Public Function BI_YearlyCalendar(Optional ByVal YearValue As Variant, Optional ByVal FormatText As String = "DDD", Optional ByVal Vertical As Boolean = False) As Variant
    Dim y As Long, m As Long
    Dim monthData As Variant
    Dim result() As Variant
    Dim offsetRow As Long, r As Long, c As Long

    If IsMissing(YearValue) Or IsEmpty(YearValue) Then y = Year(Date) Else y = CLng(YearValue)
    ReDim result(1 To 95, 1 To 7)
    offsetRow = 1
    For m = 1 To 12
        result(offsetRow, 1) = Format$(DateSerial(y, m, 1), "MMMM") & " " & y
        monthData = BI_MonthlyCalendar(m, y, FormatText, Vertical)
        For r = 1 To UBound(monthData, 1)
            For c = 1 To UBound(monthData, 2)
                result(offsetRow + r, c) = monthData(r, c)
            Next c
        Next r
        offsetRow = offsetRow + 8
    Next m
    BI_YearlyCalendar = result
End Function

'/**
' * Description: Returns the Indian rupee symbol.
' * Parameters: None.
' * Returns: Rupee symbol.
' */
Public Function BI_RupeeSymbol() As String
    BI_RupeeSymbol = ChrW$(&H20B9)
End Function

'/**
' * Description: Returns whether the year of a date is a leap year.
' * Parameters: Date - Optional date; defaults to today.
' * Returns: TRUE if the year is leap year.
' */
Public Function BI_IsLeap(Optional ByVal DateValue As Variant) As Boolean
    Dim d As Date
    If IsMissing(DateValue) Or IsEmpty(DateValue) Then d = Date Else d = CDate(DateValue)
    BI_IsLeap = (Day(DateSerial(Year(d), 3, 0)) = 29)
End Function

'/**
' * Description: Returns the financial year label for a date.
' * Parameters: Date - Optional date; YearEndMonth - Optional year-end month; Format - Optional output format.
' * Returns: Financial year label or year.
' */
Public Function BI_FinancialYear(Optional ByVal DateValue As Variant, Optional ByVal YearEndMonth As Long = 3, Optional ByVal Format As Long = 1) As Variant
    Dim d As Date, endYear As Long, startYear As Long
    If IsMissing(DateValue) Or IsEmpty(DateValue) Then d = Date Else d = CDate(DateValue)
    endYear = Year(d)
    If Month(d) > YearEndMonth Then endYear = endYear + 1
    startYear = endYear - 1
    Select Case Format
        Case 2: BI_FinancialYear = startYear & "-" & Right$(CStr(endYear), 2)
        Case 3: BI_FinancialYear = endYear
        Case Else: BI_FinancialYear = "FY " & startYear & "-" & Right$(CStr(endYear), 2)
    End Select
End Function

Public Function BI_FinancialYearStart(Optional ByVal DateValue As Variant, Optional ByVal YearEndMonth As Long = 3) As Date
    Dim d As Date, endYear As Long
    If IsMissing(DateValue) Or IsEmpty(DateValue) Then d = Date Else d = CDate(DateValue)
    endYear = Year(d)
    If Month(d) > YearEndMonth Then endYear = endYear + 1
    BI_FinancialYearStart = DateSerial(endYear - 1, YearEndMonth + 1, 1)
End Function

Public Function BI_FinancialYearEnd(Optional ByVal DateValue As Variant, Optional ByVal YearEndMonth As Long = 3) As Date
    Dim d As Date, endYear As Long
    If IsMissing(DateValue) Or IsEmpty(DateValue) Then d = Date Else d = CDate(DateValue)
    endYear = Year(d)
    If Month(d) > YearEndMonth Then endYear = endYear + 1
    BI_FinancialYearEnd = DateSerial(endYear, YearEndMonth + 1, 0)
End Function

Public Function BI_Quarter(Optional ByVal DateValue As Variant, Optional ByVal YearEndMonth As Long = 3) As Long
    Dim d As Date, shiftedMonth As Long
    If IsMissing(DateValue) Or IsEmpty(DateValue) Then d = Date Else d = CDate(DateValue)
    shiftedMonth = ((Month(d) - YearEndMonth - 1 + 12) Mod 12) + 1
    BI_Quarter = ((shiftedMonth - 1) \ 3) + 1
End Function

Public Function BI_QuarterStart(Optional ByVal DateValue As Variant, Optional ByVal YearEndMonth As Long = 3) As Date
    Dim fyStart As Date, q As Long
    fyStart = BI_FinancialYearStart(DateValue, YearEndMonth)
    q = BI_Quarter(DateValue, YearEndMonth)
    BI_QuarterStart = DateAdd("m", (q - 1) * 3, fyStart)
End Function

Public Function BI_QuarterEnd(Optional ByVal DateValue As Variant, Optional ByVal YearEndMonth As Long = 3) As Date
    BI_QuarterEnd = DateAdd("d", -1, DateAdd("m", 3, BI_QuarterStart(DateValue, YearEndMonth)))
End Function

'/**
' * Description: Returns age and datedif-style intervals between two dates.
' * Parameters: StartDate - Starting date; EndDate - Ending date.
' * Returns: Three-column table containing Y, YM, MD, M, D and YD outputs.
' */
Public Function BI_DateDif(ByVal StartDate As Date, ByVal EndDate As Date) As Variant
    On Error GoTo ErrHandler
    Dim yearsPart As Long, monthsPart As Long, daysPart As Long
    Dim totalMonths As Long, totalDays As Long, daysAfterYear As Long
    Dim anniversary As Date, monthAnchor As Date
    Dim result(1 To 6, 1 To 3) As Variant

    If EndDate < StartDate Then
        BI_DateDif = CVErr(xlErrNum)
        Exit Function
    End If

    yearsPart = DateDiff("yyyy", StartDate, EndDate)
    anniversary = DateAdd("yyyy", yearsPart, StartDate)
    If anniversary > EndDate Then
        yearsPart = yearsPart - 1
        anniversary = DateAdd("yyyy", yearsPart, StartDate)
    End If

    monthsPart = DateDiff("m", anniversary, EndDate)
    monthAnchor = DateAdd("m", monthsPart, anniversary)
    If monthAnchor > EndDate Then
        monthsPart = monthsPart - 1
        monthAnchor = DateAdd("m", monthsPart, anniversary)
    End If

    daysPart = DateDiff("d", monthAnchor, EndDate)
    totalMonths = yearsPart * 12 + monthsPart
    totalDays = DateDiff("d", StartDate, EndDate)
    daysAfterYear = DateDiff("d", anniversary, EndDate)

    result(1, 1) = "Years": result(1, 2) = "Y": result(1, 3) = yearsPart
    result(2, 1) = "Months": result(2, 2) = "YM": result(2, 3) = monthsPart
    result(3, 1) = "Days": result(3, 2) = "MD": result(3, 3) = daysPart
    result(4, 1) = "Total Months": result(4, 2) = "M": result(4, 3) = totalMonths
    result(5, 1) = "Total Days": result(5, 2) = "D": result(5, 3) = totalDays
    result(6, 1) = "Days after year": result(6, 2) = "YD": result(6, 3) = daysAfterYear
    BI_DateDif = result
    Exit Function
ErrHandler:
    BI_DateDif = CVErr(xlErrValue)
End Function

Public Function BI_NetworkDays_Indian(ByVal StartDate As Date, ByVal EndDate As Date, Optional ByVal Holidays As Variant) As Long
    Dim d As Date, countValue As Long
    For d = StartDate To EndDate
        If BI_IsIndianWorkday(d, Holidays) Then countValue = countValue + 1
    Next d
    BI_NetworkDays_Indian = countValue
End Function

Public Function BI_WorkDay_Indian(ByVal StartDate As Date, ByVal Days As Long, Optional ByVal Holidays As Variant) As Date
    Dim d As Date, remaining As Long
    d = StartDate
    remaining = Days
    Do While remaining > 0
        d = DateAdd("d", 1, d)
        If BI_IsIndianWorkday(d, Holidays) Then remaining = remaining - 1
    Loop
    BI_WorkDay_Indian = d
End Function

Public Sub BI_Tool_WorkDayIndian(control As IRibbonControl)
    On Error GoTo ErrHandler
    Unload frmBI_DateTool
    frmBI_DateTool.Tag = "WorkDay"
    frmBI_DateTool.Show
    Exit Sub
ErrHandler:
    MsgBox "Indian Workday tool failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_NetworkDaysIndian(control As IRibbonControl)
    On Error GoTo ErrHandler
    Unload frmBI_DateTool
    frmBI_DateTool.Tag = "NetworkDays"
    frmBI_DateTool.Show
    Exit Sub
ErrHandler:
    MsgBox "Indian NetworkDays tool failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_TextTools(control As IRibbonControl)
    On Error GoTo ErrHandler
    frmBI_TextTool.Show
    Exit Sub
ErrHandler:
    MsgBox "Text tool failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_MonthlyCalendar(control As IRibbonControl)
    On Error GoTo ErrHandler
    Dim m As Variant, y As Variant, outputCell As Range
    m = Application.InputBox("Enter month number (1-12). Leave blank for current month.", "BeIndian Monthly Calendar", Month(Date), Type:=1)
    If VarType(m) = vbBoolean Then Exit Sub
    y = Application.InputBox("Enter year. Leave blank for current year.", "BeIndian Monthly Calendar", Year(Date), Type:=1)
    If VarType(y) = vbBoolean Then Exit Sub
    Set outputCell = Application.InputBox("Select output top-left cell.", "BeIndian Monthly Calendar", ActiveCell.Address(External:=True), Type:=8)
    If outputCell Is Nothing Then Exit Sub
    outputCell.Resize(UBound(BI_MonthlyCalendar(CLng(m), CLng(y)), 1), UBound(BI_MonthlyCalendar(CLng(m), CLng(y)), 2)).Value = BI_MonthlyCalendar(CLng(m), CLng(y))
    outputCell.CurrentRegion.Columns.AutoFit
    Exit Sub
ErrHandler:
    MsgBox "Monthly Calendar failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_YearlyCalendar(control As IRibbonControl)
    On Error GoTo ErrHandler
    Dim y As Variant, outputCell As Range, result As Variant
    y = Application.InputBox("Enter year. Leave blank for current year.", "BeIndian Yearly Calendar", Year(Date), Type:=1)
    If VarType(y) = vbBoolean Then Exit Sub
    Set outputCell = Application.InputBox("Select output top-left cell.", "BeIndian Yearly Calendar", ActiveCell.Address(External:=True), Type:=8)
    If outputCell Is Nothing Then Exit Sub
    result = BI_YearlyCalendar(CLng(y))
    outputCell.Resize(UBound(result, 1), UBound(result, 2)).Value = result
    outputCell.CurrentRegion.Columns.AutoFit
    Exit Sub
ErrHandler:
    MsgBox "Yearly Calendar failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub
