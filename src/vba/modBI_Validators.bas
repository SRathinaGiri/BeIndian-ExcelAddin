Attribute VB_Name = "modBI_Validators"
Option Explicit

'/**
' * Description: Verifies whether the character is an alphabet from A to Z.
' * Parameters: Character - Single character to test.
' * Returns: TRUE if the character is A-Z, otherwise FALSE.
' */
Public Function BI_IsCharAtoZ(ByVal Character As Variant) As Boolean
    Dim s As String
    s = UCase$(Left$(CStr(Character), 1))
    BI_IsCharAtoZ = (Len(s) = 1 And s >= "A" And s <= "Z")
End Function

'/**
' * Description: Validates an Indian Income Tax Permanent Account Number.
' * Parameters: PAN - Indian Income Tax PAN.
' * Returns: TRUE when PAN follows AAAAA9999A pattern and valid fourth-character status codes.
' */
Public Function BI_IsValidPAN(ByVal PAN As Variant) As Boolean
    On Error GoTo SafeExit
    Dim s As String
    Dim i As Long
    s = UCase$(Trim$(CStr(PAN)))
    If Len(s) <> 10 Then GoTo SafeExit
    For i = 1 To 5
        If Not BI_IsCharAtoZ(Mid$(s, i, 1)) Then GoTo SafeExit
    Next i
    If InStr(1, "PFCATH", Mid$(s, 4, 1), vbBinaryCompare) = 0 Then GoTo SafeExit
    For i = 6 To 9
        If Not IsNumeric(Mid$(s, i, 1)) Then GoTo SafeExit
    Next i
    If Not BI_IsCharAtoZ(Right$(s, 1)) Then GoTo SafeExit
    BI_IsValidPAN = True
SafeExit:
End Function

'/**
' * Description: Uses the Luhn algorithm to test credit/debit card number validity.
' * Parameters: CCNo - Number to test.
' * Returns: TRUE if the number passes the Luhn checksum.
' */
Public Function BI_LuhnAlgorithm(ByVal CCNo As Variant) As Boolean
    On Error GoTo SafeExit
    Dim s As String
    Dim i As Long, digit As Long, sumValue As Long
    Dim shouldDouble As Boolean

    s = Trim$(CStr(CCNo))
    If Len(s) = 0 Then GoTo SafeExit
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then GoTo SafeExit
    Next i

    For i = Len(s) To 1 Step -1
        digit = CLng(Mid$(s, i, 1))
        If shouldDouble Then
            digit = digit * 2
            If digit > 9 Then digit = digit - 9
        End If
        sumValue = sumValue + digit
        shouldDouble = Not shouldDouble
    Next i

    BI_LuhnAlgorithm = (sumValue Mod 10 = 0)
SafeExit:
End Function

Public Sub BI_Tool_ValidatePAN(control As IRibbonControl)
    On Error GoTo ErrHandler
    Dim target As Range
    Set target = Application.InputBox("Select PAN cells to validate.", "BeIndian PAN Validator", Type:=8)
    If target Is Nothing Then Exit Sub
    BI_ValidatePANRange target
    Exit Sub
ErrHandler:
    MsgBox "PAN validation failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_ValidatePANRange(ByVal target As Range)
    On Error GoTo ErrHandler
    Dim cell As Range
    Dim output As Range
    Set output = target.Offset(0, target.Columns.Count).Resize(target.Rows.Count, 1)
    output.Cells(1, 1).Value = "PAN Valid?"
    For Each cell In target.Cells
        cell.Offset(0, target.Columns.Count).Value = BI_IsValidPAN(cell.Value)
    Next cell
    Exit Sub
ErrHandler:
    MsgBox "Could not validate PAN range: " & Err.Description, vbExclamation, "BeIndian"
End Sub
