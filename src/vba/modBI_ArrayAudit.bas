Attribute VB_Name = "modBI_ArrayAudit"
Option Explicit

Private Function BI_SourceTo2D(ByVal source As Variant) As Variant
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
        r = UBound(data, 1) - LBound(data, 1) + 1
        c = UBound(data, 2) - LBound(data, 2) + 1
        ReDim result(1 To r, 1 To c)
        For r = 1 To UBound(result, 1)
            For c = 1 To UBound(result, 2)
                result(r, c) = data(LBound(data, 1) + r - 1, LBound(data, 2) + c - 1)
            Next c
        Next r
        BI_SourceTo2D = result
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

    BI_SourceTo2D = result
End Function

Private Function BI_CriteriaToRow(ByVal source As Variant) As Variant
    Dim data As Variant
    Dim result() As Variant
    Dim r As Long, c As Long, n As Long

    data = BI_SourceTo2D(source)
    ReDim result(1 To 1, 1 To UBound(data, 1) * UBound(data, 2))
    For r = 1 To UBound(data, 1)
        For c = 1 To UBound(data, 2)
            n = n + 1
            result(1, n) = data(r, c)
        Next c
    Next r
    BI_CriteriaToRow = result
End Function

Public Function BI_SelectedColumns(Optional ByVal selectedColumns As Variant, Optional ByVal defaultCount As Long = 0) As Long()
    Dim cols() As Long
    Dim data As Variant
    Dim r As Long, c As Long, n As Long

    If IsMissing(selectedColumns) Or IsEmpty(selectedColumns) Then
        ReDim cols(1 To defaultCount)
        For n = 1 To defaultCount
            cols(n) = n
        Next n
        BI_SelectedColumns = cols
        Exit Function
    End If

    data = BI_SourceTo2D(selectedColumns)
    ReDim cols(1 To UBound(data, 1) * UBound(data, 2))
    For r = 1 To UBound(data, 1)
        For c = 1 To UBound(data, 2)
            n = n + 1
            cols(n) = CLng(data(r, c))
        Next c
    Next r
    ReDim Preserve cols(1 To n)
    BI_SelectedColumns = cols
End Function

Private Function BI_RowKey(ByRef data As Variant, ByVal rowIndex As Long, ByRef cols() As Long, Optional ByVal caseSensitive As Boolean = False) As String
    Dim i As Long
    Dim valueText As String

    For i = LBound(cols) To UBound(cols)
        valueText = CStr(data(rowIndex, cols(i)))
        If Not caseSensitive Then valueText = UCase$(valueText)
        BI_RowKey = BI_RowKey & "<^>" & valueText
    Next i
End Function

Private Function BI_ScalarKey(ByVal value As Variant) As String
    If IsNumeric(value) Then
        BI_ScalarKey = "N|" & Format$(CDbl(value), "0.############################")
    Else
        BI_ScalarKey = "T|" & UCase$(CStr(value))
    End If
End Function

Private Function BI_CopyRows(ByRef data As Variant, ByRef rowIndexes() As Long, ByVal rowCount As Long) As Variant
    Dim result() As Variant
    Dim r As Long, c As Long

    ReDim result(1 To rowCount, 1 To UBound(data, 2))
    For r = 1 To rowCount
        For c = 1 To UBound(data, 2)
            result(r, c) = data(rowIndexes(r), c)
        Next c
    Next r
    BI_CopyRows = result
End Function

Private Function BI_CopyFirstRows(ByRef data As Variant, ByVal rowCount As Long) As Variant
    Dim result() As Variant
    Dim r As Long, c As Long

    ReDim result(1 To rowCount, 1 To UBound(data, 2))
    For r = 1 To rowCount
        For c = 1 To UBound(data, 2)
            result(r, c) = data(r, c)
        Next c
    Next r
    BI_CopyFirstRows = result
End Function

Private Function BI_CompareValues(ByVal leftValue As Variant, ByVal rightValue As Variant) As Long
    If IsNumeric(leftValue) And IsNumeric(rightValue) Then
        If CDbl(leftValue) < CDbl(rightValue) Then
            BI_CompareValues = -1
        ElseIf CDbl(leftValue) > CDbl(rightValue) Then
            BI_CompareValues = 1
        Else
            BI_CompareValues = 0
        End If
    Else
        BI_CompareValues = StrComp(CStr(leftValue), CStr(rightValue), vbTextCompare)
    End If
End Function

Private Sub BI_Sort2DByColumn(ByRef data As Variant, ByVal sortColumn As Long, ByVal descending As Boolean)
    Dim r As Long, j As Long, c As Long
    Dim swapNeeded As Boolean
    Dim temp As Variant

    For r = 1 To UBound(data, 1) - 1
        For j = r + 1 To UBound(data, 1)
            If descending Then
                swapNeeded = (BI_CompareValues(data(j, sortColumn), data(r, sortColumn)) > 0)
            Else
                swapNeeded = (BI_CompareValues(data(j, sortColumn), data(r, sortColumn)) < 0)
            End If
            If swapNeeded Then
                For c = 1 To UBound(data, 2)
                    temp = data(r, c)
                    data(r, c) = data(j, c)
                    data(j, c) = temp
                Next c
            End If
        Next j
    Next r
End Sub

Private Function BI_CollectionToColumn(ByVal items As Collection) As Variant
    Dim result() As Variant
    Dim i As Long

    If items.Count = 0 Then
        BI_CollectionToColumn = CVErr(xlErrNA)
        Exit Function
    End If

    ReDim result(1 To items.Count, 1 To 1)
    For i = 1 To items.Count
        result(i, 1) = items(i)
    Next i
    BI_CollectionToColumn = result
End Function

Private Function BI_RowSubsetKey(ByRef data As Variant, ByVal rowIndex As Long, ByVal firstCol As Long, ByVal lastCol As Long) As String
    Dim c As Long
    For c = firstCol To lastCol
        BI_RowSubsetKey = BI_RowSubsetKey & "<^>" & CStr(data(rowIndex, c))
    Next c
End Function

Private Function BI_AggregateCollection(ByVal items As Collection, ByVal functionNum As Long) As Variant
    Dim i As Long
    Dim sumValue As Double, sumSq As Double
    Dim productValue As Double
    Dim maxValue As Double, minValue As Double
    Dim hasNumeric As Boolean

    If items.Count = 0 Then
        BI_AggregateCollection = 0
        Exit Function
    End If

    productValue = 1
    For i = 1 To items.Count
        If Len(Trim$(CStr(items(i)))) > 0 Then
            If IsNumeric(items(i)) Then
                If Not hasNumeric Then
                    maxValue = CDbl(items(i))
                    minValue = CDbl(items(i))
                    hasNumeric = True
                End If
                sumValue = sumValue + CDbl(items(i))
                sumSq = sumSq + (CDbl(items(i)) ^ 2)
                If CDbl(items(i)) > maxValue Then maxValue = CDbl(items(i))
                If CDbl(items(i)) < minValue Then minValue = CDbl(items(i))
                productValue = productValue * CDbl(items(i))
            End If
        End If
    Next i

    Select Case functionNum
        Case 1
            If hasNumeric Then BI_AggregateCollection = sumValue / items.Count Else BI_AggregateCollection = 0
        Case 2
            BI_AggregateCollection = items.Count
        Case 3
            BI_AggregateCollection = items.Count
        Case 4
            BI_AggregateCollection = maxValue
        Case 5
            BI_AggregateCollection = minValue
        Case 6
            BI_AggregateCollection = productValue
        Case 7
            If items.Count <= 1 Then
                BI_AggregateCollection = 0
            Else
                BI_AggregateCollection = Sqr((sumSq - ((sumValue * sumValue) / items.Count)) / (items.Count - 1))
            End If
        Case 8
            If items.Count = 0 Then
                BI_AggregateCollection = 0
            Else
                BI_AggregateCollection = Sqr((sumSq - ((sumValue * sumValue) / items.Count)) / items.Count)
            End If
        Case 10
            If items.Count <= 1 Then
                BI_AggregateCollection = 0
            Else
                BI_AggregateCollection = (sumSq - ((sumValue * sumValue) / items.Count)) / (items.Count - 1)
            End If
        Case 11
            If items.Count = 0 Then
                BI_AggregateCollection = 0
            Else
                BI_AggregateCollection = (sumSq - ((sumValue * sumValue) / items.Count)) / items.Count
            End If
        Case Else
            BI_AggregateCollection = sumValue
    End Select
End Function

Private Function BI_HasValue(ByVal value As Variant) As Boolean
    BI_HasValue = (Not IsEmpty(value)) And (Len(Trim$(CStr(value))) > 0)
End Function

'/**
' * Description: Compares two ranges and returns whether both ranges are equal.
' * Parameters: LeftRange - First range; RightRange - Second range; CaseSensitive - Optional case comparison flag.
' * Returns: TRUE when all values match in the same order.
' */
Public Function BI_ExactArray(ByVal LeftRange As Variant, ByVal RightRange As Variant, Optional ByVal CaseSensitive As Boolean = False) As Boolean
    On Error GoTo ErrHandler
    Dim a As Variant, b As Variant
    Dim r As Long, c As Long

    a = BI_SourceTo2D(LeftRange)
    b = BI_SourceTo2D(RightRange)
    If UBound(a, 1) <> UBound(b, 1) Or UBound(a, 2) <> UBound(b, 2) Then Exit Function

    For r = 1 To UBound(a, 1)
        For c = 1 To UBound(a, 2)
            If CaseSensitive Then
                If CStr(a(r, c)) <> CStr(b(r, c)) Then Exit Function
            Else
                If UCase$(CStr(a(r, c))) <> UCase$(CStr(b(r, c))) Then Exit Function
            End If
        Next c
    Next r

    BI_ExactArray = True
    Exit Function
ErrHandler:
    BI_ExactArray = False
End Function

'/**
' * Description: Verifies whether one record exists anywhere inside a source range.
' * Parameters: SourceRange - Source table/range; CriteriaRange - Single-row criteria range.
' * Returns: TRUE when a matching row is found.
' */
Public Function BI_ArrayContains(ByVal SourceRange As Variant, ByVal CriteriaRange As Variant) As Boolean
    On Error GoTo ErrHandler
    BI_ArrayContains = Not IsError(BI_ArrayMatch(CriteriaRange, SourceRange))
    Exit Function
ErrHandler:
    BI_ArrayContains = False
End Function

'/**
' * Description: Counts the number of times a criteria row appears in a source range.
' * Parameters: SourceRange - Source table/range; CriteriaRange - Single-row criteria range.
' * Returns: Match count.
' */
Public Function BI_ArrayCountif(ByVal SourceRange As Variant, ByVal CriteriaRange As Variant) As Long
    On Error GoTo ErrHandler
    Dim data As Variant, criteria As Variant
    Dim cols() As Long
    Dim criteriaKey As String
    Dim r As Long

    data = BI_SourceTo2D(SourceRange)
    criteria = BI_CriteriaToRow(CriteriaRange)
    cols = BI_SelectedColumns(Empty, UBound(criteria, 2))
    criteriaKey = BI_RowKey(criteria, 1, cols)

    For r = 1 To UBound(data, 1)
        If BI_RowKey(data, r, cols) = criteriaKey Then BI_ArrayCountif = BI_ArrayCountif + 1
    Next r
    Exit Function
ErrHandler:
    BI_ArrayCountif = 0
End Function

'/**
' * Description: Filters a source range for a multi-column criteria row.
' * Parameters: SourceRange - Source table/range; CriteriaRange - Single-row criteria range; SelectedColumns - Optional column indexes to compare.
' * Returns: Matching rows from the source table.
' */
Public Function BI_ArrayFilter(ByVal SourceRange As Variant, ByVal CriteriaRange As Variant, Optional ByVal SelectedColumns As Variant) As Variant
    On Error GoTo ErrHandler
    Dim data As Variant, criteria As Variant
    Dim cols() As Long, criteriaCols() As Long
    Dim criteriaKey As String
    Dim rowIndexes() As Long
    Dim r As Long, matches As Long

    data = BI_SourceTo2D(SourceRange)
    criteria = BI_CriteriaToRow(CriteriaRange)
    cols = BI_SelectedColumns(SelectedColumns, UBound(criteria, 2))
    criteriaCols = BI_SelectedColumns(Empty, UBound(criteria, 2))
    criteriaKey = BI_RowKey(criteria, 1, criteriaCols)
    ReDim rowIndexes(1 To UBound(data, 1))

    For r = 1 To UBound(data, 1)
        If BI_RowKey(data, r, cols) = criteriaKey Then
            matches = matches + 1
            rowIndexes(matches) = r
        End If
    Next r

    If matches = 0 Then
        BI_ArrayFilter = CVErr(xlErrNA)
    Else
        BI_ArrayFilter = BI_CopyRows(data, rowIndexes, matches)
    End If
    Exit Function
ErrHandler:
    BI_ArrayFilter = CVErr(xlErrValue)
End Function

'/**
' * Description: Returns only the rows from one range that are available in another range.
' * Parameters: SourceRange - Table/range to filter; LookupRange - Lookup table/range.
' * Returns: Matching rows from SourceRange.
' */
Public Function BI_ArrayInArray(ByVal SourceRange As Variant, ByVal LookupRange As Variant) As Variant
    On Error GoTo ErrHandler
    Dim data As Variant, lookupData As Variant
    Dim cols1() As Long, cols2() As Long
    Dim lookupMap As Object
    Dim rowIndexes() As Long
    Dim r As Long, matches As Long
    Dim key As String

    data = BI_SourceTo2D(SourceRange)
    lookupData = BI_SourceTo2D(LookupRange)
    cols1 = BI_SelectedColumns(Empty, UBound(data, 2))
    cols2 = BI_SelectedColumns(Empty, UBound(lookupData, 2))
    Set lookupMap = CreateObject("Scripting.Dictionary")

    For r = 1 To UBound(lookupData, 1)
        lookupMap(BI_RowKey(lookupData, r, cols2)) = True
    Next r

    ReDim rowIndexes(1 To UBound(data, 1))
    For r = 1 To UBound(data, 1)
        key = BI_RowKey(data, r, cols1)
        If lookupMap.Exists(key) Then
            matches = matches + 1
            rowIndexes(matches) = r
        End If
    Next r

    If matches = 0 Then
        BI_ArrayInArray = CVErr(xlErrNA)
    Else
        BI_ArrayInArray = BI_CopyRows(data, rowIndexes, matches)
    End If
    Exit Function
ErrHandler:
    BI_ArrayInArray = CVErr(xlErrValue)
End Function

'/**
' * Description: Returns the row number of a range where a criteria row is first found.
' * Parameters: CriteriaRange - Single-row criteria range; LookupRange - Source table/range; SelectedColumns - Optional columns to compare.
' * Returns: 1-based row index or #N/A.
' */
Public Function BI_ArrayMatch(ByVal CriteriaRange As Variant, ByVal LookupRange As Variant, Optional ByVal SelectedColumns As Variant) As Variant
    On Error GoTo ErrHandler
    Dim data As Variant, criteria As Variant
    Dim cols() As Long, criteriaCols() As Long
    Dim criteriaKey As String
    Dim r As Long

    data = BI_SourceTo2D(LookupRange)
    criteria = BI_CriteriaToRow(CriteriaRange)
    cols = BI_SelectedColumns(SelectedColumns, UBound(criteria, 2))
    criteriaCols = BI_SelectedColumns(Empty, UBound(criteria, 2))
    criteriaKey = BI_RowKey(criteria, 1, criteriaCols)

    For r = 1 To UBound(data, 1)
        If BI_RowKey(data, r, cols) = criteriaKey Then
            BI_ArrayMatch = r
            Exit Function
        End If
    Next r

    BI_ArrayMatch = CVErr(xlErrNA)
    Exit Function
ErrHandler:
    BI_ArrayMatch = CVErr(xlErrValue)
End Function

'/**
' * Description: Returns how many times the same complete row or selected columns occur in a range.
' * Parameters: DataRange - Source table/range; SelectedColumns - Optional selected columns.
' * Returns: Count per source row.
' */
Public Function BI_RowOccurence(ByVal DataRange As Variant, Optional ByVal SelectedColumns As Variant) As Variant
    On Error GoTo ErrHandler
    Dim data2D As Variant
    Dim cols() As Long
    Dim counts As Object
    Dim result() As Variant
    Dim key As String
    Dim r As Long

    data2D = BI_SourceTo2D(DataRange)
    cols = BI_SelectedColumns(SelectedColumns, UBound(data2D, 2))
    Set counts = CreateObject("Scripting.Dictionary")

    For r = 1 To UBound(data2D, 1)
        key = BI_RowKey(data2D, r, cols)
        If counts.Exists(key) Then
            counts(key) = counts(key) + 1
        Else
            counts.Add key, 1
        End If
    Next r

    ReDim result(1 To UBound(data2D, 1), 1 To 1)
    For r = 1 To UBound(data2D, 1)
        result(r, 1) = counts(BI_RowKey(data2D, r, cols))
    Next r
    BI_RowOccurence = result
    Exit Function
ErrHandler:
    BI_RowOccurence = CVErr(xlErrValue)
End Function

'/**
' * Description: Returns a running total for a numeric range.
' * Parameters: SourceRange - Numeric source values.
' * Returns: Column vector of cumulative totals.
' */
Public Function BI_RunningTotal(ByVal SourceRange As Variant) As Variant
    On Error GoTo ErrHandler
    Dim values() As Double
    Dim result() As Variant
    Dim i As Long
    Dim running As Double

    values = BI_ToDoubleVector(SourceRange, False)
    ReDim result(1 To UBound(values) - LBound(values) + 1, 1 To 1)
    For i = LBound(values) To UBound(values)
        running = running + values(i)
        result(i - LBound(values) + 1, 1) = running
    Next i
    BI_RunningTotal = result
    Exit Function
ErrHandler:
    BI_RunningTotal = CVErr(xlErrValue)
End Function

'/**
' * Description: Returns a summary by group using a subtotal-style aggregation code.
' * Parameters: DataRange - Source table/range; ValueColumn - Numeric/value column; SummaryColumn - Grouping column; FunctionNum - Optional aggregation code.
' * Returns: Two-column group summary.
' */
Public Function BI_Summary(ByVal DataRange As Variant, ByVal ValueColumn As Long, ByVal SummaryColumn As Long, Optional ByVal FunctionNum As Long = 9) As Variant
    On Error GoTo ErrHandler
    Dim data2D As Variant
    Dim stats As Object, orders As Object, numerics As Object
    Dim orderList As Collection
    Dim groupKey As String, displayKey As Variant
    Dim arr As Variant, result() As Variant
    Dim r As Long, i As Long
    Dim value As Variant

    data2D = BI_SourceTo2D(DataRange)
    Set stats = CreateObject("Scripting.Dictionary")
    Set orders = CreateObject("Scripting.Dictionary")
    Set numerics = CreateObject("Scripting.Dictionary")
    Set orderList = New Collection

    For r = 1 To UBound(data2D, 1)
        displayKey = data2D(r, SummaryColumn)
        groupKey = BI_ScalarKey(data2D(r, SummaryColumn))
        value = data2D(r, ValueColumn)

        If Not orders.Exists(groupKey) Then
            orders.Add groupKey, displayKey
            orderList.Add groupKey
            ReDim arr(1 To 7)
            stats.Add groupKey, arr
        End If

        arr = stats(groupKey)
        arr(1) = arr(1) + 1
        If Len(Trim$(CStr(value))) > 0 Then arr(2) = arr(2) + 1
        If IsNumeric(value) Then
            arr(3) = arr(3) + CDbl(value)
            If (Not numerics.Exists(groupKey & "|hasnum")) Or numerics(groupKey & "|hasnum") = False Then
                arr(4) = CDbl(value)
                arr(5) = CDbl(value)
                arr(6) = CDbl(value)
                numerics(groupKey & "|hasnum") = True
            Else
                If CDbl(value) > arr(4) Then arr(4) = CDbl(value)
                If CDbl(value) < arr(5) Then arr(5) = CDbl(value)
                arr(6) = arr(6) * CDbl(value)
            End If
            arr(7) = arr(7) + (CDbl(value) ^ 2)
        End If
        stats(groupKey) = arr
        If Not numerics.Exists(groupKey & "|hasnum") Then numerics(groupKey & "|hasnum") = IsNumeric(value)
    Next r

    ReDim result(1 To orderList.Count, 1 To 2)
    For i = 1 To orderList.Count
        groupKey = CStr(orderList(i))
        arr = stats(groupKey)
        result(i, 1) = orders(groupKey)
        Select Case FunctionNum
            Case 1
                If arr(2) = 0 Then
                    result(i, 2) = 0
                Else
                    result(i, 2) = arr(3) / arr(2)
                End If
            Case 2
                result(i, 2) = arr(2)
            Case 3
                result(i, 2) = arr(1)
            Case 4
                result(i, 2) = arr(4)
            Case 5
                result(i, 2) = arr(5)
            Case 6
                result(i, 2) = arr(6)
            Case 7
                If arr(2) <= 1 Then
                    result(i, 2) = 0
                Else
                    result(i, 2) = Sqr((arr(7) - ((arr(3) * arr(3)) / arr(2))) / (arr(2) - 1))
                End If
            Case 8
                If arr(2) = 0 Then
                    result(i, 2) = 0
                Else
                    result(i, 2) = Sqr((arr(7) - ((arr(3) * arr(3)) / arr(2))) / arr(2))
                End If
            Case 10
                If arr(2) <= 1 Then
                    result(i, 2) = 0
                Else
                    result(i, 2) = (arr(7) - ((arr(3) * arr(3)) / arr(2))) / (arr(2) - 1)
                End If
            Case 11
                If arr(2) = 0 Then
                    result(i, 2) = 0
                Else
                    result(i, 2) = (arr(7) - ((arr(3) * arr(3)) / arr(2))) / arr(2)
                End If
            Case Else
                result(i, 2) = arr(3)
        End Select
    Next i
    BI_Summary = result
    Exit Function
ErrHandler:
    BI_Summary = CVErr(xlErrValue)
End Function

'/**
' * Description: Returns duplicate values from a selected range.
' * Parameters: SourceRange - Source list/range.
' * Returns: Single-column array of duplicate values.
' */
Public Function BI_FindDuplicates(ByVal SourceRange As Variant) As Variant
    On Error GoTo ErrHandler
    Dim data As Variant
    Dim counts As Object, seen As Object
    Dim items As Collection
    Dim key As String
    Dim r As Long

    data = BI_ToColumnVector(SourceRange)
    Set counts = CreateObject("Scripting.Dictionary")
    Set seen = CreateObject("Scripting.Dictionary")
    Set items = New Collection

    For r = 1 To UBound(data, 1)
        key = BI_ScalarKey(data(r, 1))
        If counts.Exists(key) Then
            counts(key) = counts(key) + 1
        Else
            counts.Add key, 1
        End If
    Next r

    For r = 1 To UBound(data, 1)
        key = BI_ScalarKey(data(r, 1))
        If counts(key) > 1 Then
            If Not seen.Exists(key) Then
                seen.Add key, True
                items.Add data(r, 1)
            End If
        End If
    Next r

    BI_FindDuplicates = BI_CollectionToColumn(items)
    Exit Function
ErrHandler:
    BI_FindDuplicates = CVErr(xlErrValue)
End Function

'/**
' * Description: Returns the values that occur exactly once in a selected range.
' * Parameters: SourceRange - Source list/range.
' * Returns: Single-column array of unique values.
' */
Public Function BI_FindUnique(ByVal SourceRange As Variant) As Variant
    On Error GoTo ErrHandler
    Dim data As Variant
    Dim counts As Object
    Dim items As Collection
    Dim key As String
    Dim r As Long

    data = BI_ToColumnVector(SourceRange)
    Set counts = CreateObject("Scripting.Dictionary")
    Set items = New Collection

    For r = 1 To UBound(data, 1)
        key = BI_ScalarKey(data(r, 1))
        If counts.Exists(key) Then
            counts(key) = counts(key) + 1
        Else
            counts.Add key, 1
        End If
    Next r

    For r = 1 To UBound(data, 1)
        key = BI_ScalarKey(data(r, 1))
        If counts(key) = 1 Then items.Add data(r, 1)
    Next r

    BI_FindUnique = BI_CollectionToColumn(items)
    Exit Function
ErrHandler:
    BI_FindUnique = CVErr(xlErrValue)
End Function

'/**
' * Description: Returns missing integers between the minimum and maximum values in a selected range.
' * Parameters: SourceRange - Numeric source list/range.
' * Returns: Single-column array of missing integers.
' */
Public Function BI_FindMissingNumbers(ByVal SourceRange As Variant) As Variant
    On Error GoTo ErrHandler
    Dim data As Variant
    Dim present As Object
    Dim items As Collection
    Dim minValue As Long, maxValue As Long
    Dim current As Long, r As Long

    data = BI_ToColumnVector(SourceRange)
    Set present = CreateObject("Scripting.Dictionary")
    Set items = New Collection

    minValue = CLng(data(1, 1))
    maxValue = CLng(data(1, 1))
    For r = 1 To UBound(data, 1)
        If IsNumeric(data(r, 1)) Then
            current = CLng(data(r, 1))
            If current < minValue Then minValue = current
            If current > maxValue Then maxValue = current
            present(CStr(current)) = True
        End If
    Next r

    For current = minValue To maxValue
        If Not present.Exists(CStr(current)) Then items.Add current
    Next current

    BI_FindMissingNumbers = BI_CollectionToColumn(items)
    Exit Function
ErrHandler:
    BI_FindMissingNumbers = CVErr(xlErrValue)
End Function

'/**
' * Description: Returns top N rows sorted by a selected column.
' * Parameters: SourceRange - Source table/range; IndexColumn - Sort column; N - Number of rows.
' * Returns: Sorted top rows.
' */
Public Function BI_TopN(ByVal SourceRange As Variant, Optional ByVal IndexColumn As Long = 1, Optional ByVal N As Long = 1) As Variant
    On Error GoTo ErrHandler
    Dim data As Variant

    data = BI_SourceTo2D(SourceRange)
    BI_Sort2DByColumn data, IndexColumn, True
    If N > UBound(data, 1) Then N = UBound(data, 1)
    BI_TopN = BI_CopyFirstRows(data, N)
    Exit Function
ErrHandler:
    BI_TopN = CVErr(xlErrValue)
End Function

'/**
' * Description: Returns bottom N rows sorted by a selected column.
' * Parameters: SourceRange - Source table/range; IndexColumn - Sort column; N - Number of rows.
' * Returns: Sorted bottom rows.
' */
Public Function BI_BottomN(ByVal SourceRange As Variant, Optional ByVal IndexColumn As Long = 1, Optional ByVal N As Long = 1) As Variant
    On Error GoTo ErrHandler
    Dim data As Variant

    data = BI_SourceTo2D(SourceRange)
    BI_Sort2DByColumn data, IndexColumn, False
    If N > UBound(data, 1) Then N = UBound(data, 1)
    BI_BottomN = BI_CopyFirstRows(data, N)
    Exit Function
ErrHandler:
    BI_BottomN = CVErr(xlErrValue)
End Function

'/**
' * Description: Returns top rows that together account for a target cumulative percentage.
' * Parameters: SourceRange - Source table/range; IndexColumn - Numeric column; Percent - Target percentage as decimal.
' * Returns: Sorted rows up to the cumulative threshold.
' */
Public Function BI_TopPercent(ByVal SourceRange As Variant, ByVal IndexColumn As Long, ByVal Percent As Double) As Variant
    On Error GoTo ErrHandler
    Dim data2D As Variant
    Dim totalValue As Double, running As Double, rowsNeeded As Long, r As Long

    data2D = BI_SourceTo2D(SourceRange)
    BI_Sort2DByColumn data2D, IndexColumn, True
    For r = 1 To UBound(data2D, 1)
        totalValue = totalValue + CDbl(data2D(r, IndexColumn))
    Next r

    For r = 1 To UBound(data2D, 1)
        running = running + CDbl(data2D(r, IndexColumn))
        rowsNeeded = rowsNeeded + 1
        If running >= totalValue * Percent Then Exit For
    Next r

    BI_TopPercent = BI_CopyFirstRows(data2D, rowsNeeded)
    Exit Function
ErrHandler:
    BI_TopPercent = CVErr(xlErrValue)
End Function

'/**
' * Description: Returns bottom rows that together account for a target cumulative percentage.
' * Parameters: SourceRange - Source table/range; IndexColumn - Numeric column; Percent - Target percentage as decimal.
' * Returns: Sorted rows up to the cumulative threshold.
' */
Public Function BI_BottomPercent(ByVal SourceRange As Variant, ByVal IndexColumn As Long, ByVal Percent As Double) As Variant
    On Error GoTo ErrHandler
    Dim data2D As Variant
    Dim totalValue As Double, running As Double, rowsNeeded As Long, r As Long

    data2D = BI_SourceTo2D(SourceRange)
    BI_Sort2DByColumn data2D, IndexColumn, False
    For r = 1 To UBound(data2D, 1)
        totalValue = totalValue + CDbl(data2D(r, IndexColumn))
    Next r

    For r = 1 To UBound(data2D, 1)
        running = running + CDbl(data2D(r, IndexColumn))
        rowsNeeded = rowsNeeded + 1
        If running >= totalValue * Percent Then Exit For
    Next r

    BI_BottomPercent = BI_CopyFirstRows(data2D, rowsNeeded)
    Exit Function
ErrHandler:
    BI_BottomPercent = CVErr(xlErrValue)
End Function

'/**
' * Description: Returns Z-scores for a numeric range.
' * Parameters: SourceRange - Numeric source values; Population - Optional TRUE for population stdev.
' * Returns: Single-column array of z-scores.
' */
Public Function BI_ZScore(ByVal SourceRange As Variant, Optional ByVal Population As Boolean = False) As Variant
    On Error GoTo ErrHandler
    Dim values() As Double
    Dim result() As Variant
    Dim avg As Double, stdev As Double
    Dim i As Long

    values = BI_ToDoubleVector(SourceRange, False)
    For i = LBound(values) To UBound(values)
        avg = avg + values(i)
    Next i
    avg = avg / (UBound(values) - LBound(values) + 1)
    For i = LBound(values) To UBound(values)
        stdev = stdev + ((values(i) - avg) ^ 2)
    Next i
    If Population Then
        stdev = Sqr(stdev / (UBound(values) - LBound(values) + 1))
    Else
        stdev = Sqr(stdev / Application.Max(1, UBound(values) - LBound(values)))
    End If

    ReDim result(1 To UBound(values) - LBound(values) + 1, 1 To 1)
    For i = LBound(values) To UBound(values)
        If stdev = 0 Then
            result(i - LBound(values) + 1, 1) = 0
        Else
            result(i - LBound(values) + 1, 1) = (values(i) - avg) / stdev
        End If
    Next i
    BI_ZScore = result
    Exit Function
ErrHandler:
    BI_ZScore = CVErr(xlErrValue)
End Function

'/**
' * Description: Returns Relative Size Factor by category.
' * Parameters: Categories - Grouping values; RowValue - Numeric values; First - Largest rank; Second - Second rank.
' * Returns: Four-column table sorted by relative size factor descending.
' */
Public Function BI_RelativeSizeFactor(ByVal Categories As Variant, ByVal RowValue As Variant, Optional ByVal First As Long = 1, Optional ByVal Second As Long = 2) As Variant
    On Error GoTo ErrHandler
    Dim catData As Variant, valueData As Variant
    Dim buckets As Object, orderList As Collection
    Dim labels As Object
    Dim key As String
    Dim rowCount As Long, i As Long, j As Long, k As Long
    Dim values() As Double
    Dim result() As Variant
    Dim firstValue As Double, secondValue As Double
    Dim temp As Double
    Dim body() As Variant

    catData = BI_ToColumnVector(Categories)
    valueData = BI_ToColumnVector(RowValue)
    If UBound(catData, 1) <> UBound(valueData, 1) Then
        BI_RelativeSizeFactor = CVErr(xlErrRef)
        Exit Function
    End If

    Set buckets = CreateObject("Scripting.Dictionary")
    Set labels = CreateObject("Scripting.Dictionary")
    Set orderList = New Collection
    For i = 1 To UBound(catData, 1)
        key = BI_ScalarKey(catData(i, 1))
        If Not buckets.Exists(key) Then
            buckets.Add key, New Collection
            labels.Add key, catData(i, 1)
            orderList.Add key
        End If
        buckets(key).Add CDbl(valueData(i, 1))
    Next i

    ReDim result(1 To orderList.Count + 1, 1 To 4)
    result(1, 1) = "Category"
    result(1, 2) = "First Value"
    result(1, 3) = "Second Value"
    result(1, 4) = "Relative Size Factor"

    For i = 1 To orderList.Count
        key = CStr(orderList(i))
        rowCount = buckets(key).Count
        ReDim values(1 To rowCount)
        For j = 1 To rowCount
            values(j) = CDbl(buckets(key)(j))
        Next j
        For j = 1 To rowCount - 1
            For k = j + 1 To rowCount
                If values(k) > values(j) Then
                    temp = values(j)
                    values(j) = values(k)
                    values(k) = temp
                End If
            Next k
        Next j

        firstValue = IIf(rowCount >= First, values(First), 0)
        secondValue = IIf(rowCount >= Second, values(Second), 0)
        result(i + 1, 1) = labels(key)
        result(i + 1, 2) = firstValue
        result(i + 1, 3) = secondValue
        If secondValue = 0 Then
            result(i + 1, 4) = 0
        Else
            result(i + 1, 4) = firstValue / secondValue
        End If
    Next i

    ReDim body(1 To UBound(result, 1) - 1, 1 To 4)
    For i = 2 To UBound(result, 1)
        For j = 1 To 4
            body(i - 1, j) = result(i, j)
        Next j
    Next i
    BI_Sort2DByColumn body, 4, True
    For i = 1 To UBound(body, 1)
        For j = 1 To 4
            result(i + 1, j) = body(i, j)
        Next j
    Next i

    BI_RelativeSizeFactor = result
    Exit Function
ErrHandler:
    BI_RelativeSizeFactor = CVErr(xlErrValue)
End Function

'/**
' * Description: Sorts a range by one or more columns.
' * Parameters: SourceRange - Source table/range; Order - 1 for ascending or -1 for descending; ColumnIndex - Optional last sort column.
' * Returns: Sorted range.
' */
Public Function BI_ASort(ByVal SourceRange As Variant, Optional ByVal Order As Long = 1, Optional ByVal ColumnIndex As Long = 0) As Variant
    On Error GoTo ErrHandler
    Dim data As Variant
    Dim sortCol As Long, currentCol As Long
    Dim descending As Boolean

    data = BI_SourceTo2D(SourceRange)
    If ColumnIndex <= 0 Or ColumnIndex > UBound(data, 2) Then sortCol = UBound(data, 2) Else sortCol = ColumnIndex
    descending = (Order < 0)
    For currentCol = sortCol To 1 Step -1
        BI_Sort2DByColumn data, currentCol, descending
    Next currentCol
    BI_ASort = data
    Exit Function
ErrHandler:
    BI_ASort = CVErr(xlErrValue)
End Function

'/**
' * Description: Returns cumulative attribute counts for values in a range.
' * Parameters: SourceRange - Source list/range.
' * Returns: Two-column table of source values and cumulative count.
' */
Public Function BI_AttributeUnitCount(ByVal SourceRange As Variant) As Variant
    On Error GoTo ErrHandler
    Dim data As Variant
    Dim counts As Object
    Dim result() As Variant
    Dim key As String
    Dim r As Long

    data = BI_ToColumnVector(SourceRange)
    Set counts = CreateObject("Scripting.Dictionary")
    ReDim result(1 To UBound(data, 1), 1 To 2)
    For r = 1 To UBound(data, 1)
        key = BI_ScalarKey(data(r, 1))
        If counts.Exists(key) Then
            counts(key) = counts(key) + 1
        Else
            counts.Add key, 1
        End If
        result(r, 1) = data(r, 1)
        result(r, 2) = counts(key)
    Next r
    BI_AttributeUnitCount = result
    Exit Function
ErrHandler:
    BI_AttributeUnitCount = CVErr(xlErrValue)
End Function

'/**
' * Description: Returns unpivoted data from a crosstab range.
' * Parameters: UnpivotData - Data to unpivot including headers; RemainingData - Optional identifier columns.
' * Returns: Unpivoted range with Attributes and Values columns.
' */
Public Function BI_Unpivot(ByVal UnpivotData As Variant, Optional ByVal RemainingData As Variant) As Variant
    On Error GoTo ErrHandler
    Dim data As Variant, remain As Variant
    Dim hasRemaining As Boolean
    Dim result() As Variant
    Dim outCols As Long, outRows As Long
    Dim r As Long, c As Long, outRow As Long, remainCols As Long, rc As Long

    data = BI_SourceTo2D(UnpivotData)
    hasRemaining = Not (IsMissing(RemainingData) Or IsEmpty(RemainingData))
    If hasRemaining Then
        remain = BI_SourceTo2D(RemainingData)
        remainCols = UBound(remain, 2)
    End If

    outRows = (UBound(data, 1) - 1) * UBound(data, 2) + 1
    outCols = 2 + remainCols
    ReDim result(1 To outRows, 1 To outCols)

    outRow = 1
    If hasRemaining Then
        For rc = 1 To remainCols
            result(1, rc) = remain(1, rc)
        Next rc
        result(1, remainCols + 1) = "Attributes"
        result(1, remainCols + 2) = "Values"
    Else
        result(1, 1) = "Attributes"
        result(1, 2) = "Values"
    End If

    For r = 2 To UBound(data, 1)
        For c = 1 To UBound(data, 2)
            If BI_HasValue(data(r, c)) And CDbl(Val(CStr(data(r, c)))) <> 0 Then
                outRow = outRow + 1
                If hasRemaining Then
                    For rc = 1 To remainCols
                        result(outRow, rc) = remain(r, rc)
                    Next rc
                    result(outRow, remainCols + 1) = data(1, c)
                    result(outRow, remainCols + 2) = data(r, c)
                Else
                    result(outRow, 1) = data(1, c)
                    result(outRow, 2) = data(r, c)
                End If
            End If
        Next c
    Next r

    BI_Unpivot = BI_CopyFirstRows(result, outRow)
    Exit Function
ErrHandler:
    BI_Unpivot = CVErr(xlErrValue)
End Function

'/**
' * Description: Returns unpivoted data leaving the first column as the identifier column.
' * Parameters: SourceRange - Crosstab range with first column preserved.
' * Returns: Unpivoted range.
' */
Public Function BI_UnpivotExceptFirst(ByVal SourceRange As Variant) As Variant
    On Error GoTo ErrHandler
    Dim data As Variant
    Dim unpivotPart() As Variant, remainPart() As Variant
    Dim r As Long, c As Long

    data = BI_SourceTo2D(SourceRange)
    ReDim unpivotPart(1 To UBound(data, 1), 1 To UBound(data, 2) - 1)
    ReDim remainPart(1 To UBound(data, 1), 1 To 1)
    For r = 1 To UBound(data, 1)
        remainPart(r, 1) = data(r, 1)
        For c = 2 To UBound(data, 2)
            unpivotPart(r, c - 1) = data(r, c)
        Next c
    Next r
    BI_UnpivotExceptFirst = BI_Unpivot(unpivotPart, remainPart)
    Exit Function
ErrHandler:
    BI_UnpivotExceptFirst = CVErr(xlErrValue)
End Function

'/**
' * Description: Creates a pivot-style summary from normalized data.
' * Parameters: SourceRange - Source normalized data; ValueColumn - Value column; RowColumn - Row grouping column; ColColumn - Column grouping column; FunctionNum - Optional summary function.
' * Returns: Pivoted range.
' */
Public Function BI_Pivot(ByVal SourceRange As Variant, ByVal ValueColumn As Long, ByVal RowColumn As Long, ByVal ColColumn As Long, Optional ByVal FunctionNum As Long = 9) As Variant
    On Error GoTo ErrHandler
    Dim data As Variant
    Dim rowKeys As Object, colKeys As Object, cellMap As Object
    Dim rowOrder As Collection, colOrder As Collection
    Dim rowKey As String, colKey As String, bucketKey As String
    Dim result() As Variant
    Dim r As Long, i As Long, j As Long
    Dim values As Collection

    data = BI_SourceTo2D(SourceRange)
    Set rowKeys = CreateObject("Scripting.Dictionary")
    Set colKeys = CreateObject("Scripting.Dictionary")
    Set cellMap = CreateObject("Scripting.Dictionary")
    Set rowOrder = New Collection
    Set colOrder = New Collection

    For r = 1 To UBound(data, 1)
        rowKey = CStr(data(r, RowColumn))
        colKey = CStr(data(r, ColColumn))
        If Not rowKeys.Exists(rowKey) Then
            rowKeys.Add rowKey, rowKey
            rowOrder.Add rowKey
        End If
        If Not colKeys.Exists(colKey) Then
            colKeys.Add colKey, colKey
            colOrder.Add colKey
        End If
        bucketKey = rowKey & "<|>" & colKey
        If Not cellMap.Exists(bucketKey) Then cellMap.Add bucketKey, New Collection
        cellMap(bucketKey).Add data(r, ValueColumn)
    Next r

    ReDim result(1 To rowOrder.Count + 1, 1 To colOrder.Count + 1)
    result(1, 1) = vbNullString
    For j = 1 To colOrder.Count
        result(1, j + 1) = colOrder(j)
    Next j
    For i = 1 To rowOrder.Count
        result(i + 1, 1) = rowOrder(i)
        For j = 1 To colOrder.Count
            bucketKey = CStr(rowOrder(i)) & "<|>" & CStr(colOrder(j))
            If cellMap.Exists(bucketKey) Then
                Set values = cellMap(bucketKey)
                result(i + 1, j + 1) = BI_AggregateCollection(values, FunctionNum)
            Else
                result(i + 1, j + 1) = 0
            End If
        Next j
    Next i

    BI_Pivot = result
    Exit Function
ErrHandler:
    BI_Pivot = CVErr(xlErrValue)
End Function

'/**
' * Description: Consolidates multiple crosstab ranges into a single summary table.
' * Parameters: FunctionNum - Summary function; DataRange1..DataRange10 - Crosstab ranges.
' * Returns: Consolidated summary.
' */
Public Function BI_Consolidate(ByVal FunctionNum As Long, ByVal DataRange1 As Variant, Optional ByVal DataRange2 As Variant, Optional ByVal DataRange3 As Variant, Optional ByVal DataRange4 As Variant, Optional ByVal DataRange5 As Variant, Optional ByVal DataRange6 As Variant, Optional ByVal DataRange7 As Variant, Optional ByVal DataRange8 As Variant, Optional ByVal DataRange9 As Variant, Optional ByVal DataRange10 As Variant) As Variant
    On Error GoTo ErrHandler
    Dim totalRows As Long, outCols As Long, startRow As Long
    Dim part As Variant
    Dim combined() As Variant
    Dim r As Long, c As Long

    part = BI_UnpivotExceptFirst(DataRange1)
    If IsArray(part) Then
        outCols = UBound(part, 2)
        totalRows = totalRows + UBound(part, 1) - 1
    End If
    If Not IsMissing(DataRange2) Then
        part = BI_UnpivotExceptFirst(DataRange2)
        If IsArray(part) Then totalRows = totalRows + UBound(part, 1) - 1
    End If
    If Not IsMissing(DataRange3) Then
        part = BI_UnpivotExceptFirst(DataRange3)
        If IsArray(part) Then totalRows = totalRows + UBound(part, 1) - 1
    End If
    If Not IsMissing(DataRange4) Then
        part = BI_UnpivotExceptFirst(DataRange4)
        If IsArray(part) Then totalRows = totalRows + UBound(part, 1) - 1
    End If
    If Not IsMissing(DataRange5) Then
        part = BI_UnpivotExceptFirst(DataRange5)
        If IsArray(part) Then totalRows = totalRows + UBound(part, 1) - 1
    End If
    If Not IsMissing(DataRange6) Then
        part = BI_UnpivotExceptFirst(DataRange6)
        If IsArray(part) Then totalRows = totalRows + UBound(part, 1) - 1
    End If
    If Not IsMissing(DataRange7) Then
        part = BI_UnpivotExceptFirst(DataRange7)
        If IsArray(part) Then totalRows = totalRows + UBound(part, 1) - 1
    End If
    If Not IsMissing(DataRange8) Then
        part = BI_UnpivotExceptFirst(DataRange8)
        If IsArray(part) Then totalRows = totalRows + UBound(part, 1) - 1
    End If
    If Not IsMissing(DataRange9) Then
        part = BI_UnpivotExceptFirst(DataRange9)
        If IsArray(part) Then totalRows = totalRows + UBound(part, 1) - 1
    End If
    If Not IsMissing(DataRange10) Then
        part = BI_UnpivotExceptFirst(DataRange10)
        If IsArray(part) Then totalRows = totalRows + UBound(part, 1) - 1
    End If

    If totalRows = 0 Then
        BI_Consolidate = CVErr(xlErrNA)
        Exit Function
    End If

    ReDim combined(1 To totalRows, 1 To outCols)
    startRow = 1
    part = BI_UnpivotExceptFirst(DataRange1)
    If IsArray(part) Then
        For r = 2 To UBound(part, 1)
            For c = 1 To UBound(part, 2)
                combined(startRow, c) = part(r, c)
            Next c
            startRow = startRow + 1
        Next r
    End If
    If Not IsMissing(DataRange2) Then BI_CopyConsolidateRows combined, startRow, BI_UnpivotExceptFirst(DataRange2)
    If Not IsMissing(DataRange3) Then BI_CopyConsolidateRows combined, startRow, BI_UnpivotExceptFirst(DataRange3)
    If Not IsMissing(DataRange4) Then BI_CopyConsolidateRows combined, startRow, BI_UnpivotExceptFirst(DataRange4)
    If Not IsMissing(DataRange5) Then BI_CopyConsolidateRows combined, startRow, BI_UnpivotExceptFirst(DataRange5)
    If Not IsMissing(DataRange6) Then BI_CopyConsolidateRows combined, startRow, BI_UnpivotExceptFirst(DataRange6)
    If Not IsMissing(DataRange7) Then BI_CopyConsolidateRows combined, startRow, BI_UnpivotExceptFirst(DataRange7)
    If Not IsMissing(DataRange8) Then BI_CopyConsolidateRows combined, startRow, BI_UnpivotExceptFirst(DataRange8)
    If Not IsMissing(DataRange9) Then BI_CopyConsolidateRows combined, startRow, BI_UnpivotExceptFirst(DataRange9)
    If Not IsMissing(DataRange10) Then BI_CopyConsolidateRows combined, startRow, BI_UnpivotExceptFirst(DataRange10)

    BI_Consolidate = BI_Pivot(combined, 3, 1, 2, FunctionNum)
    Exit Function
ErrHandler:
    BI_Consolidate = CVErr(xlErrValue)
End Function

Private Sub BI_CopyConsolidateRows(ByRef combined As Variant, ByRef startRow As Long, ByVal part As Variant)
    Dim r As Long, c As Long
    If Not IsArray(part) Then Exit Sub
    For r = 2 To UBound(part, 1)
        For c = 1 To UBound(part, 2)
            combined(startRow, c) = part(r, c)
        Next c
        startRow = startRow + 1
    Next r
End Sub

Private Function BI_PickRange(ByVal prompt As String, Optional ByVal defaultRange As Variant) As Range
    On Error GoTo ErrHandler
    If IsMissing(defaultRange) Or IsEmpty(defaultRange) Then
        Set BI_PickRange = Application.InputBox(prompt, "BeIndian", Type:=8)
    Else
        Set BI_PickRange = Application.InputBox(prompt, "BeIndian", defaultRange, Type:=8)
    End If
    Exit Function
ErrHandler:
    Set BI_PickRange = Nothing
End Function

Private Function BI_PickOutputCell(ByVal prompt As String) As Range
    On Error GoTo ErrHandler
    Set BI_PickOutputCell = Application.InputBox(prompt, "BeIndian", ActiveCell.Address(External:=True), Type:=8)
    Exit Function
ErrHandler:
    Set BI_PickOutputCell = Nothing
End Function

Private Sub BI_WriteResult(ByVal targetCell As Range, ByVal result As Variant)
    If targetCell Is Nothing Then Exit Sub
    If IsArray(result) Then
        targetCell.Resize(UBound(result, 1), UBound(result, 2)).Value = result
        targetCell.Resize(UBound(result, 1), UBound(result, 2)).Columns.AutoFit
    Else
        targetCell.Value = result
    End If
End Sub

Private Sub BI_RunSingleRangeTool(ByVal sourcePrompt As String, ByVal outputPrompt As String, ByVal toolName As String)
    On Error GoTo ErrHandler
    Dim sourceRange As Range
    Dim outputCell As Range
    Dim result As Variant

    Set sourceRange = BI_PickRange(sourcePrompt, Selection.Address(External:=True))
    If sourceRange Is Nothing Then Exit Sub
    Set outputCell = BI_PickOutputCell(outputPrompt)
    If outputCell Is Nothing Then Exit Sub

    Select Case toolName
        Case "FindDuplicates": result = BI_FindDuplicates(sourceRange)
        Case "FindUnique": result = BI_FindUnique(sourceRange)
        Case "FindMissingNumbers": result = BI_FindMissingNumbers(sourceRange)
        Case "RunningTotal": result = BI_RunningTotal(sourceRange)
        Case Else: Err.Raise vbObjectError + 640, "BI_RunSingleRangeTool", "Unknown tool name."
    End Select

    BI_WriteResult outputCell.Cells(1, 1), result
    Exit Sub
ErrHandler:
    MsgBox toolName & " failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_FindDuplicates(control As IRibbonControl)
    BI_RunSingleRangeTool "Select the source list for duplicate detection.", "Select output top-left cell.", "FindDuplicates"
End Sub

Public Sub BI_Tool_FindUnique(control As IRibbonControl)
    BI_RunSingleRangeTool "Select the source list for unique-only detection.", "Select output top-left cell.", "FindUnique"
End Sub

Public Sub BI_Tool_FindMissingNumbers(control As IRibbonControl)
    BI_RunSingleRangeTool "Select the numeric list for missing number detection.", "Select output top-left cell.", "FindMissingNumbers"
End Sub

Public Sub BI_Tool_RunningTotal(control As IRibbonControl)
    BI_RunSingleRangeTool "Select the numeric range for running total.", "Select output top-left cell.", "RunningTotal"
End Sub

Public Sub BI_Tool_TopN(control As IRibbonControl)
    On Error GoTo ErrHandler
    Dim dataRange As Range
    Dim outputCell As Range
    Dim indexColumn As Variant
    Dim topN As Variant

    Set dataRange = BI_PickRange("Select the source table for Top N.", Selection.Address(External:=True))
    If dataRange Is Nothing Then Exit Sub
    indexColumn = Application.InputBox("Enter the column index to rank by.", "BeIndian Top N", 1, Type:=1)
    If VarType(indexColumn) = vbBoolean Then Exit Sub
    topN = Application.InputBox("Enter the number of rows to return.", "BeIndian Top N", 10, Type:=1)
    If VarType(topN) = vbBoolean Then Exit Sub
    Set outputCell = BI_PickOutputCell("Select output top-left cell.")
    If outputCell Is Nothing Then Exit Sub

    BI_WriteResult outputCell.Cells(1, 1), BI_TopN(dataRange, CLng(indexColumn), CLng(topN))
    Exit Sub
ErrHandler:
    MsgBox "Top N failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_BottomN(control As IRibbonControl)
    On Error GoTo ErrHandler
    Dim dataRange As Range
    Dim outputCell As Range
    Dim indexColumn As Variant
    Dim bottomN As Variant

    Set dataRange = BI_PickRange("Select the source table for Bottom N.", Selection.Address(External:=True))
    If dataRange Is Nothing Then Exit Sub
    indexColumn = Application.InputBox("Enter the column index to rank by.", "BeIndian Bottom N", 1, Type:=1)
    If VarType(indexColumn) = vbBoolean Then Exit Sub
    bottomN = Application.InputBox("Enter the number of rows to return.", "BeIndian Bottom N", 10, Type:=1)
    If VarType(bottomN) = vbBoolean Then Exit Sub
    Set outputCell = BI_PickOutputCell("Select output top-left cell.")
    If outputCell Is Nothing Then Exit Sub

    BI_WriteResult outputCell.Cells(1, 1), BI_BottomN(dataRange, CLng(indexColumn), CLng(bottomN))
    Exit Sub
ErrHandler:
    MsgBox "Bottom N failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_TopPercent(control As IRibbonControl)
    On Error GoTo ErrHandler
    Dim dataRange As Range
    Dim outputCell As Range
    Dim indexColumn As Variant
    Dim percentValue As Variant

    Set dataRange = BI_PickRange("Select the source table for Top Percent.", Selection.Address(External:=True))
    If dataRange Is Nothing Then Exit Sub
    indexColumn = Application.InputBox("Enter the column index to rank by.", "BeIndian Top Percent", 1, Type:=1)
    If VarType(indexColumn) = vbBoolean Then Exit Sub
    percentValue = Application.InputBox("Enter the cumulative percentage as decimal (example 0.2 for 20%).", "BeIndian Top Percent", 0.2, Type:=1)
    If VarType(percentValue) = vbBoolean Then Exit Sub
    Set outputCell = BI_PickOutputCell("Select output top-left cell.")
    If outputCell Is Nothing Then Exit Sub

    BI_WriteResult outputCell.Cells(1, 1), BI_TopPercent(dataRange, CLng(indexColumn), CDbl(percentValue))
    Exit Sub
ErrHandler:
    MsgBox "Top Percent failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_BottomPercent(control As IRibbonControl)
    On Error GoTo ErrHandler
    Dim dataRange As Range
    Dim outputCell As Range
    Dim indexColumn As Variant
    Dim percentValue As Variant

    Set dataRange = BI_PickRange("Select the source table for Bottom Percent.", Selection.Address(External:=True))
    If dataRange Is Nothing Then Exit Sub
    indexColumn = Application.InputBox("Enter the column index to rank by.", "BeIndian Bottom Percent", 1, Type:=1)
    If VarType(indexColumn) = vbBoolean Then Exit Sub
    percentValue = Application.InputBox("Enter the cumulative percentage as decimal (example 0.2 for 20%).", "BeIndian Bottom Percent", 0.2, Type:=1)
    If VarType(percentValue) = vbBoolean Then Exit Sub
    Set outputCell = BI_PickOutputCell("Select output top-left cell.")
    If outputCell Is Nothing Then Exit Sub

    BI_WriteResult outputCell.Cells(1, 1), BI_BottomPercent(dataRange, CLng(indexColumn), CDbl(percentValue))
    Exit Sub
ErrHandler:
    MsgBox "Bottom Percent failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_ZScore(control As IRibbonControl)
    On Error GoTo ErrHandler
    Dim dataRange As Range
    Dim outputCell As Range
    Dim populationMode As Variant

    Set dataRange = BI_PickRange("Select the numeric range for Z-Score.", Selection.Address(External:=True))
    If dataRange Is Nothing Then Exit Sub
    populationMode = Application.InputBox("Use population standard deviation? Enter TRUE or FALSE.", "BeIndian Z-Score", "FALSE", Type:=2)
    If VarType(populationMode) = vbBoolean Then Exit Sub
    Set outputCell = BI_PickOutputCell("Select output top-left cell.")
    If outputCell Is Nothing Then Exit Sub

    BI_WriteResult outputCell.Cells(1, 1), BI_ZScore(dataRange, CBool(populationMode))
    Exit Sub
ErrHandler:
    MsgBox "Z-Score failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_ASort(control As IRibbonControl)
    On Error GoTo ErrHandler
    Dim sourceRange As Range
    Dim outputCell As Range
    Dim sortOrder As Variant
    Dim lastSortColumn As Variant
    Dim result As Variant

    Set sourceRange = BI_PickRange("Select the range to sort.", Selection.Address(External:=True))
    If sourceRange Is Nothing Then Exit Sub
    sortOrder = Application.InputBox("Enter 1 for ascending or -1 for descending sort.", "BeIndian Sort", 1, Type:=1)
    If VarType(sortOrder) = vbBoolean Then Exit Sub
    lastSortColumn = Application.InputBox("Enter the last column index to sort by. Enter 0 for all columns.", "BeIndian Sort", sourceRange.Columns.Count, Type:=1)
    If VarType(lastSortColumn) = vbBoolean Then Exit Sub
    Set outputCell = BI_PickOutputCell("Select output top-left cell.")
    If outputCell Is Nothing Then Exit Sub
    result = BI_ASort(sourceRange, CLng(sortOrder), CLng(lastSortColumn))
    BI_WriteResult outputCell.Cells(1, 1), result
    Exit Sub
ErrHandler:
    MsgBox "Sort failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_AttributeUnitCount(control As IRibbonControl)
    On Error GoTo ErrHandler
    Dim sourceRange As Range
    Dim outputCell As Range
    Dim result As Variant

    Set sourceRange = BI_PickRange("Select the range for cumulative attribute count.", Selection.Address(External:=True))
    If sourceRange Is Nothing Then Exit Sub
    Set outputCell = BI_PickOutputCell("Select output top-left cell.")
    If outputCell Is Nothing Then Exit Sub
    result = BI_AttributeUnitCount(sourceRange)
    BI_WriteResult outputCell.Cells(1, 1), result
    Exit Sub
ErrHandler:
    MsgBox "Attribute Unit Count failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_Unpivot(control As IRibbonControl)
    On Error GoTo ErrHandler
    Dim sourceRange As Range
    Dim outputCell As Range
    Dim result As Variant

    Set sourceRange = BI_PickRange("Select the crosstab range to unpivot.", Selection.Address(External:=True))
    If sourceRange Is Nothing Then Exit Sub
    Set outputCell = BI_PickOutputCell("Select output top-left cell.")
    If outputCell Is Nothing Then Exit Sub
    result = BI_UnpivotExceptFirst(sourceRange)
    BI_WriteResult outputCell.Cells(1, 1), result
    Exit Sub
ErrHandler:
    MsgBox "Unpivot failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_Pivot(control As IRibbonControl)
    On Error GoTo ErrHandler
    Dim sourceRange As Range
    Dim outputCell As Range
    Dim valueColumn As Variant, rowColumn As Variant, colColumn As Variant, functionNum As Variant
    Dim result As Variant

    Set sourceRange = BI_PickRange("Select the normalized source range for pivot.", Selection.Address(External:=True))
    If sourceRange Is Nothing Then Exit Sub
    valueColumn = Application.InputBox("Enter the value column index.", "BeIndian Pivot", 3, Type:=1)
    If VarType(valueColumn) = vbBoolean Then Exit Sub
    rowColumn = Application.InputBox("Enter the row grouping column index.", "BeIndian Pivot", 1, Type:=1)
    If VarType(rowColumn) = vbBoolean Then Exit Sub
    colColumn = Application.InputBox("Enter the column grouping column index.", "BeIndian Pivot", 2, Type:=1)
    If VarType(colColumn) = vbBoolean Then Exit Sub
    functionNum = Application.InputBox("Enter subtotal-style function number. Default is 9 for Sum.", "BeIndian Pivot", 9, Type:=1)
    If VarType(functionNum) = vbBoolean Then Exit Sub
    Set outputCell = BI_PickOutputCell("Select output top-left cell.")
    If outputCell Is Nothing Then Exit Sub
    result = BI_Pivot(sourceRange, CLng(valueColumn), CLng(rowColumn), CLng(colColumn), CLng(functionNum))
    BI_WriteResult outputCell.Cells(1, 1), result
    Exit Sub
ErrHandler:
    MsgBox "Pivot failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub
