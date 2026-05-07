Attribute VB_Name = "modBI_Core"
Option Explicit

Public Const BI_VERSION As String = "2.0-vba-preview-2026.05.07.2"

'/**
' * Description: Returns add-in metadata.
' * Parameters: None.
' * Returns: Two-column Variant array with property names and values.
' */
Public Function BI_About() As Variant
    Dim result(1 To 6, 1 To 2) As Variant
    result(1, 1) = "Template Name": result(1, 2) = "Be Indian Excel Audit Tool"
    result(2, 1) = "Version": result(2, 2) = BI_VERSION
    result(3, 1) = "Author": result(3, 2) = "CA S. Rathinagiri"
    result(4, 1) = "Copyright": result(4, 2) = "None. 100% Free without warranties"
    result(5, 1) = "Port": result(5, 2) = "Pure VBA Excel Add-in"
    result(6, 1) = "Build Date": result(6, 2) = "2026-05-07"
    BI_About = result
End Function

Public Function BI_CleanText(ByVal value As Variant) As String
    BI_CleanText = Trim$(CStr(value))
End Function

Public Function BI_NormalizeText(ByVal value As Variant, Optional ByVal caseSensitive As Boolean = False) As String
    If caseSensitive Then
        BI_NormalizeText = CStr(value)
    Else
        BI_NormalizeText = UCase$(CStr(value))
    End If
End Function

Public Function BI_IsMissingOptional(Optional ByVal value As Variant) As Boolean
    BI_IsMissingOptional = IsMissing(value) Or IsEmpty(value)
End Function

Public Function BI_ToColumnVector(ByVal source As Variant) As Variant
    Dim data As Variant
    Dim r As Long, c As Long, n As Long
    Dim rowsCount As Long, colsCount As Long
    Dim result() As Variant
    Dim lower2 As Long, upper2 As Long

    If TypeName(source) = "Range" Then
        data = source.Value2
    Else
        data = source
    End If

    If IsArray(data) Then
        On Error Resume Next
        lower2 = LBound(data, 2)
        upper2 = UBound(data, 2)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            rowsCount = UBound(data) - LBound(data) + 1
            ReDim result(1 To rowsCount, 1 To 1)
            For r = LBound(data) To UBound(data)
                n = n + 1
                result(n, 1) = data(r)
            Next r
            BI_ToColumnVector = result
            Exit Function
        End If
        On Error GoTo 0

        rowsCount = UBound(data, 1) - LBound(data, 1) + 1
        colsCount = upper2 - lower2 + 1
        ReDim result(1 To rowsCount * colsCount, 1 To 1)
        For r = LBound(data, 1) To UBound(data, 1)
            For c = lower2 To upper2
                n = n + 1
                result(n, 1) = data(r, c)
            Next c
        Next r
    Else
        ReDim result(1 To 1, 1 To 1)
        result(1, 1) = data
    End If

    BI_ToColumnVector = result
End Function

Public Function BI_ToDoubleVector(ByVal source As Variant, Optional ByVal ignoreNonNumeric As Boolean = True) As Double()
    Dim col As Variant
    Dim values() As Double
    Dim r As Long, n As Long

    col = BI_ToColumnVector(source)
    ReDim values(1 To UBound(col, 1))

    For r = 1 To UBound(col, 1)
        If IsNumeric(col(r, 1)) Then
            n = n + 1
            values(n) = CDbl(col(r, 1))
        ElseIf Not ignoreNonNumeric Then
            Err.Raise vbObjectError + 510, "BI_ToDoubleVector", "Non-numeric value found."
        End If
    Next r

    If n = 0 Then
        ReDim values(1 To 1)
        values(1) = 0
    Else
        ReDim Preserve values(1 To n)
    End If

    BI_ToDoubleVector = values
End Function

Public Function BI_OutputToActiveCell(ByVal data As Variant) As Range
    Dim rowsCount As Long, colsCount As Long
    If Not IsArray(data) Then
        ActiveCell.Value = data
        Set BI_OutputToActiveCell = ActiveCell
        Exit Function
    End If

    rowsCount = UBound(data, 1) - LBound(data, 1) + 1
    colsCount = UBound(data, 2) - LBound(data, 2) + 1
    Set BI_OutputToActiveCell = ActiveCell.Resize(rowsCount, colsCount)
    BI_OutputToActiveCell.Value = data
End Function

Public Function BI_IsSecondOrFourthSaturday(ByVal value As Date) As Boolean
    BI_IsSecondOrFourthSaturday = (Weekday(value, vbSunday) = vbSaturday) _
        And ((Day(value) > 7 And Day(value) <= 14) Or (Day(value) > 21 And Day(value) <= 28))
End Function

Public Function BI_IsIndianWorkday(ByVal value As Date, Optional ByVal holidays As Variant) As Boolean
    BI_IsIndianWorkday = True
    If Weekday(value, vbSunday) = vbSunday Then BI_IsIndianWorkday = False
    If BI_IsSecondOrFourthSaturday(value) Then BI_IsIndianWorkday = False
    If Not IsMissing(holidays) Then
        If TypeName(holidays) = "Range" Then
            If Application.WorksheetFunction.CountIf(holidays, CLng(value)) > 0 Then BI_IsIndianWorkday = False
        End If
    End If
End Function
