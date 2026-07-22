VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBI_F091 
   Caption         =   "BI_Summary"
   ClientHeight    =   2840
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7100
   OleObjectBlob   =   "frmBI_F091.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBI_F091"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const FUNCTION_NAME As String = "BI_Summary"
Private Const PARAM_COUNT As Long = 4
Private Const PARAM_LABELS As String = "Data|ValueColumn|SummaryColumn|Function"
Private Const DESCRIPTION_TEXT As String = "Returns Summary of data according to any attribute."

Private Function BI_ParamLabel(ByVal index As Long) As String
    Dim labels() As String
    labels = Split(PARAM_LABELS, "|")
    If index <= UBound(labels) + 1 Then
        BI_ParamLabel = labels(index - 1)
    Else
        BI_ParamLabel = "Argument " & CStr(index)
    End If
End Function

Private Function BI_ControlExists(ByVal controlName As String) As Boolean
    On Error GoTo MissingControl
    Dim ctl As Object
    Set ctl = Me.Controls(controlName)
    BI_ControlExists = True
    Exit Function
MissingControl:
    BI_ControlExists = False
End Function

Private Sub BI_SetLabel(ByVal controlName As String, ByVal captionText As String)
    On Error Resume Next
    Me.Controls(controlName).Caption = captionText
    Me.Controls(controlName).Visible = True
    On Error GoTo 0
End Sub

Private Sub BI_SetVisible(ByVal controlName As String, ByVal isVisible As Boolean)
    On Error Resume Next
    Me.Controls(controlName).Visible = isVisible
    On Error GoTo 0
End Sub

Private Sub BI_MoveControl(ByVal controlName As String, ByVal leftValue As Single, ByVal topValue As Single, Optional ByVal widthValue As Single = -1)
    On Error Resume Next
    With Me.Controls(controlName)
        .Left = leftValue
        .Top = topValue
        If widthValue >= 0 Then .Width = widthValue
    End With
    On Error GoTo 0
End Sub

Private Function BI_DisplayName() As String
    If Left$(FUNCTION_NAME, 3) = "BI_" Then
        BI_DisplayName = Mid$(FUNCTION_NAME, 4)
    Else
        BI_DisplayName = FUNCTION_NAME
    End If
End Function

Private Function BI_IsManagedControl(ByVal controlName As String) As Boolean
    BI_IsManagedControl = (controlName = "lblBIHeader" Or controlName = "lblBIDescription" Or _
        controlName = "lblOutput" Or controlName = "txtOutput" Or controlName = "cmdOutput" Or _
        controlName = "cmdRun" Or controlName = "cmdCancel" Or controlName = "chkNewSheet" Or _
        Left$(controlName, 6) = "cboArg" Or _
        Left$(controlName, 6) = "lblArg" Or Left$(controlName, 6) = "txtArg" Or Left$(controlName, 6) = "cmdArg")
End Function

Private Sub BI_HideLegacyUsageLabels()
    On Error Resume Next
    Dim ctl As Object
    Dim nm As String
    For Each ctl In Me.Controls
        nm = ctl.Name
        If Not BI_IsManagedControl(nm) Then
            Select Case TypeName(ctl)
                Case "Label", "TextBox"
                    ctl.Visible = False
            End Select
        End If
    Next ctl
    On Error GoTo 0
End Sub

Private Function BI_DescriptionHeight() As Single
    Dim h As Single
    h = 30 + (Len(DESCRIPTION_TEXT) \ 95) * 14
    If h < 34 Then h = 34
    If h > 76 Then h = 76
    BI_DescriptionHeight = h
End Function

Private Sub BI_EnsureHeader()
    On Error Resume Next
    Dim lbl As Object
    Set lbl = Me.Controls("lblBIHeader")
    If lbl Is Nothing Then Set lbl = Me.Controls.Add("Forms.Label.1", "lblBIHeader", True)
    With lbl
        .Caption = BI_DisplayName()
        .Left = 12
        .Top = 12
        .Width = 560
        .Height = 22
        .WordWrap = False
        .Font.Bold = True
        .Font.Size = 11
        .ForeColor = RGB(31, 78, 121)
    End With
    On Error GoTo 0
End Sub

Private Sub BI_EnsureDescription(ByVal topValue As Single)
    On Error Resume Next
    Dim lbl As Object
    Set lbl = Me.Controls("lblBIDescription")
    If lbl Is Nothing Then Set lbl = Me.Controls.Add("Forms.Label.1", "lblBIDescription", True)
    With lbl
        .Caption = DESCRIPTION_TEXT
        .Left = 12
        .Top = topValue
        .Width = 548
        .Height = BI_DescriptionHeight()
        .WordWrap = True
        .Font.Bold = False
        .Font.Size = 9
        .ForeColor = RGB(80, 80, 80)
        .Visible = (Len(DESCRIPTION_TEXT) > 0)
    End With
    On Error GoTo 0
End Sub

Private Function BI_LongestLabelLength() As Long
    Dim i As Long
    Dim n As Long
    For i = 1 To PARAM_COUNT
        If Len(BI_ParamLabel(i)) > n Then n = Len(BI_ParamLabel(i))
    Next i
    If n < Len("Output") Then n = Len("Output")
    BI_LongestLabelLength = n
End Function

Private Function BI_CompactFormWidth() As Single
    Dim w As Single
    w = 330 + (BI_LongestLabelLength() * 4.2)
    If w < 460 Then w = 460
    If w > 620 Then w = 620
    BI_CompactFormWidth = w
End Function


Private Function BI_CurrentSelectionRange() As Range
    On Error Resume Next
    If TypeName(Selection) = "Range" Then
        Set BI_CurrentSelectionRange = Selection
    Else
        Set BI_CurrentSelectionRange = ActiveCell
    End If
    On Error GoTo 0
End Function

Private Function BI_DefaultInputAddress() As String
    Dim selectedRange As Range
    Set selectedRange = BI_CurrentSelectionRange()
    If Not selectedRange Is Nothing Then BI_DefaultInputAddress = selectedRange.Address(External:=True)
End Function

Private Function BI_DefaultOutputAddress() As String
    Dim selectedRange As Range
    Set selectedRange = BI_CurrentSelectionRange()
    If Not selectedRange Is Nothing Then
        BI_DefaultOutputAddress = selectedRange.Cells(1, 1).Offset(0, selectedRange.Columns.Count).Address(External:=True)
    End If
End Function

Private Sub BI_ApplySelectionDefaults()
    On Error Resume Next
    If PARAM_COUNT > 0 Then Me.Controls("txtArg1").Text = BI_DefaultInputAddress()
    Me.txtOutput.Text = BI_DefaultOutputAddress()
    On Error GoTo 0
End Sub

Private Function BI_ArgumentUsesSelector(ByVal argIndex As Long) As Boolean
    Select Case FUNCTION_NAME & "|" & CStr(argIndex)
        Case "BI_AnovaSingleFactor|3", "BI_AnovaTwoFactorsWithoutReplication|3", _
             "BI_BeginsWith|3", "BI_Contains|3", "BI_Correlation|2", "BI_Covariance|2", _
             "BI_EndsWith|3", "BI_ExactArray|3", "BI_ExponentialSmoothing|3", _
             "BI_GenerateRandomBetween|5", "BI_MonthlyCalendar|4", "BI_MovingAverage|3", _
             "BI_MovingAverage|4", "BI_MUS_ExtractSample|5", "BI_Regression|3", _
             "BI_YearlyCalendar|3", "BI_ZScore|2", _
             "BI_ASort|2", "BI_Rank_Dense|3", _
             "BI_FinancialYear|2", "BI_FinancialYear|3", "BI_FinancialYearEnd|2", _
             "BI_FinancialYearStart|2", "BI_MonthlyCalendar|1", "BI_MonthlyCalendar|3", _
             "BI_Pivot|5", "BI_Quarter|2", "BI_QuarterEnd|2", "BI_QuarterStart|2", _
             "BI_Summary|4", "BI_YearlyCalendar|2"
            BI_ArgumentUsesSelector = True
    End Select
End Function

Private Sub BI_AddSelectorItem(ByVal cbo As Object, ByVal captionText As String)
    cbo.AddItem captionText
End Sub

Private Function BI_IsBooleanSelector(ByVal argIndex As Long) As Boolean
    Select Case FUNCTION_NAME & "|" & CStr(argIndex)
        Case "BI_AnovaSingleFactor|3", "BI_AnovaTwoFactorsWithoutReplication|3", _
             "BI_BeginsWith|3", "BI_Contains|3", "BI_Correlation|2", "BI_Covariance|2", _
             "BI_EndsWith|3", "BI_ExactArray|3", "BI_ExponentialSmoothing|3", _
             "BI_GenerateRandomBetween|5", "BI_MonthlyCalendar|4", "BI_MovingAverage|3", _
             "BI_MovingAverage|4", "BI_MUS_ExtractSample|5", "BI_Regression|3", _
             "BI_YearlyCalendar|3", "BI_ZScore|2"
            BI_IsBooleanSelector = True
    End Select
End Function

Private Function BI_IsSummarySelector(ByVal argIndex As Long) As Boolean
    Select Case FUNCTION_NAME & "|" & CStr(argIndex)
        Case "BI_Pivot|5", "BI_Summary|4"
            BI_IsSummarySelector = True
    End Select
End Function

Private Function BI_IsYearEndMonthSelector(ByVal argIndex As Long) As Boolean
    Select Case FUNCTION_NAME & "|" & CStr(argIndex)
        Case "BI_FinancialYear|2", "BI_FinancialYearEnd|2", "BI_FinancialYearStart|2", _
             "BI_Quarter|2", "BI_QuarterEnd|2", "BI_QuarterStart|2"
            BI_IsYearEndMonthSelector = True
    End Select
End Function

Private Sub BI_EnsureArgumentSelector(ByVal argIndex As Long, ByVal leftValue As Single, ByVal topValue As Single, ByVal widthValue As Single)
    On Error Resume Next
    Dim cbo As Object
    Dim controlName As String
    controlName = "cboArg" & CStr(argIndex)
    Set cbo = Me.Controls(controlName)
    If cbo Is Nothing Then Set cbo = Me.Controls.Add("Forms.ComboBox.1", controlName, True)
    With cbo
        .Left = leftValue
        .Top = topValue
        .Width = widthValue
        .Height = 18
        .Style = 2
        .Clear
        If BI_IsBooleanSelector(argIndex) Then
            BI_AddSelectorItem cbo, "False"
            BI_AddSelectorItem cbo, "True"
            .Value = "False"
        ElseIf BI_IsSummarySelector(argIndex) Then
            BI_AddSelectorItem cbo, "Sum"
            BI_AddSelectorItem cbo, "Average"
            BI_AddSelectorItem cbo, "Count"
            BI_AddSelectorItem cbo, "Count Numbers"
            BI_AddSelectorItem cbo, "Max"
            BI_AddSelectorItem cbo, "Min"
            BI_AddSelectorItem cbo, "Product"
            BI_AddSelectorItem cbo, "StdDev Sample"
            BI_AddSelectorItem cbo, "StdDev Population"
            BI_AddSelectorItem cbo, "Variance Sample"
            BI_AddSelectorItem cbo, "Variance Population"
            .Value = "Sum"
        ElseIf BI_IsYearEndMonthSelector(argIndex) Then
            BI_AddSelectorItem cbo, "January"
            BI_AddSelectorItem cbo, "February"
            BI_AddSelectorItem cbo, "March"
            BI_AddSelectorItem cbo, "April"
            BI_AddSelectorItem cbo, "May"
            BI_AddSelectorItem cbo, "June"
            BI_AddSelectorItem cbo, "July"
            BI_AddSelectorItem cbo, "August"
            BI_AddSelectorItem cbo, "September"
            BI_AddSelectorItem cbo, "October"
            BI_AddSelectorItem cbo, "November"
            BI_AddSelectorItem cbo, "December"
            .Value = "March"
        ElseIf FUNCTION_NAME = "BI_ASort" And argIndex = 2 Then
            BI_AddSelectorItem cbo, "Ascending"
            BI_AddSelectorItem cbo, "Descending"
            .Value = "Ascending"
        ElseIf FUNCTION_NAME = "BI_Rank_Dense" And argIndex = 3 Then
            BI_AddSelectorItem cbo, "Descending"
            BI_AddSelectorItem cbo, "Ascending"
            .Value = "Descending"
        ElseIf FUNCTION_NAME = "BI_FinancialYear" And argIndex = 3 Then
            BI_AddSelectorItem cbo, "FY 2023-24"
            BI_AddSelectorItem cbo, "2023-24"
            BI_AddSelectorItem cbo, "2024"
            .Value = "FY 2023-24"
        ElseIf FUNCTION_NAME = "BI_MonthlyCalendar" And argIndex = 1 Then
            BI_AddSelectorItem cbo, "Current month"
            BI_AddSelectorItem cbo, "January"
            BI_AddSelectorItem cbo, "February"
            BI_AddSelectorItem cbo, "March"
            BI_AddSelectorItem cbo, "April"
            BI_AddSelectorItem cbo, "May"
            BI_AddSelectorItem cbo, "June"
            BI_AddSelectorItem cbo, "July"
            BI_AddSelectorItem cbo, "August"
            BI_AddSelectorItem cbo, "September"
            BI_AddSelectorItem cbo, "October"
            BI_AddSelectorItem cbo, "November"
            BI_AddSelectorItem cbo, "December"
            .Value = "Current month"
        ElseIf (FUNCTION_NAME = "BI_MonthlyCalendar" And argIndex = 3) Or (FUNCTION_NAME = "BI_YearlyCalendar" And argIndex = 2) Then
            BI_AddSelectorItem cbo, "DDD"
            BI_AddSelectorItem cbo, "DDDD"
            .Value = "DDD"
        End If
        .Visible = True
    End With
    On Error GoTo 0
End Sub

Private Function BI_MonthNumberFromName(ByVal monthName As String) As Variant
    Select Case monthName
        Case "January": BI_MonthNumberFromName = 1
        Case "February": BI_MonthNumberFromName = 2
        Case "March": BI_MonthNumberFromName = 3
        Case "April": BI_MonthNumberFromName = 4
        Case "May": BI_MonthNumberFromName = 5
        Case "June": BI_MonthNumberFromName = 6
        Case "July": BI_MonthNumberFromName = 7
        Case "August": BI_MonthNumberFromName = 8
        Case "September": BI_MonthNumberFromName = 9
        Case "October": BI_MonthNumberFromName = 10
        Case "November": BI_MonthNumberFromName = 11
        Case "December": BI_MonthNumberFromName = 12
        Case Else: BI_MonthNumberFromName = Empty
    End Select
End Function

Private Function BI_SummaryFunctionNumFromText(ByVal captionText As String) As Long
    Select Case captionText
        Case "Average": BI_SummaryFunctionNumFromText = 1
        Case "Count": BI_SummaryFunctionNumFromText = 2
        Case "Count Numbers": BI_SummaryFunctionNumFromText = 3
        Case "Max": BI_SummaryFunctionNumFromText = 4
        Case "Min": BI_SummaryFunctionNumFromText = 5
        Case "Product": BI_SummaryFunctionNumFromText = 6
        Case "StdDev Sample": BI_SummaryFunctionNumFromText = 7
        Case "StdDev Population": BI_SummaryFunctionNumFromText = 8
        Case "Variance Sample": BI_SummaryFunctionNumFromText = 10
        Case "Variance Population": BI_SummaryFunctionNumFromText = 11
        Case Else: BI_SummaryFunctionNumFromText = 9
    End Select
End Function

Private Function BI_SelectorValue(ByVal argIndex As Long) As Variant
    Dim selectedText As String
    selectedText = CStr(Me.Controls("cboArg" & CStr(argIndex)).Value)
    If BI_IsBooleanSelector(argIndex) Then
        BI_SelectorValue = (selectedText = "True")
    ElseIf BI_IsSummarySelector(argIndex) Then
        BI_SelectorValue = BI_SummaryFunctionNumFromText(selectedText)
    ElseIf BI_IsYearEndMonthSelector(argIndex) Then
        BI_SelectorValue = BI_MonthNumberFromName(selectedText)
    ElseIf FUNCTION_NAME = "BI_ASort" And argIndex = 2 Then
        If selectedText = "Descending" Then BI_SelectorValue = -1 Else BI_SelectorValue = 1
    ElseIf FUNCTION_NAME = "BI_Rank_Dense" And argIndex = 3 Then
        If selectedText = "Ascending" Then BI_SelectorValue = 1 Else BI_SelectorValue = 0
    ElseIf FUNCTION_NAME = "BI_FinancialYear" And argIndex = 3 Then
        Select Case selectedText
            Case "2023-24": BI_SelectorValue = 2
            Case "2024": BI_SelectorValue = 3
            Case Else: BI_SelectorValue = 1
        End Select
    ElseIf FUNCTION_NAME = "BI_MonthlyCalendar" And argIndex = 1 Then
        BI_SelectorValue = BI_MonthNumberFromName(selectedText)
    ElseIf (FUNCTION_NAME = "BI_MonthlyCalendar" And argIndex = 3) Or (FUNCTION_NAME = "BI_YearlyCalendar" And argIndex = 2) Then
        BI_SelectorValue = selectedText
    Else
        BI_SelectorValue = selectedText
    End If
End Function
Private Function BI_NewSheetName() As String
    Dim baseName As String
    baseName = Replace(BI_DisplayName(), ".", "_")
    If Len(baseName) > 24 Then baseName = Left$(baseName, 24)
    BI_NewSheetName = baseName
End Function

Private Function BI_UniqueSheetName(ByVal baseName As String) As String
    Dim candidate As String
    Dim suffix As Long
    candidate = baseName
    Do While BI_WorksheetExists(candidate)
        suffix = suffix + 1
        candidate = Left$(baseName, 27 - Len(CStr(suffix))) & "_" & CStr(suffix)
    Loop
    BI_UniqueSheetName = candidate
End Function

Private Function BI_WorksheetExists(ByVal sheetName As String) As Boolean
    On Error GoTo MissingSheet
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets(sheetName)
    BI_WorksheetExists = True
    Exit Function
MissingSheet:
    BI_WorksheetExists = False
End Function

Private Function BI_UseNewSheet() As Boolean
    On Error Resume Next
    BI_UseNewSheet = CBool(Me.Controls("chkNewSheet").Value)
    On Error GoTo 0
End Function

Private Function BI_ResolveOutputCell() As Range
    Dim ws As Worksheet
    If BI_UseNewSheet() Then
        Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveSheet)
        ws.Name = BI_UniqueSheetName(BI_NewSheetName())
        Set BI_ResolveOutputCell = ws.Range("A1")
    Else
        Set BI_ResolveOutputCell = Application.Range(Me.txtOutput.Text).Cells(1, 1)
    End If
End Function

Private Sub BI_EnsureNewSheetOption(ByVal leftValue As Single, ByVal topValue As Single, ByVal widthValue As Single)
    On Error Resume Next
    Dim chk As Object
    Set chk = Me.Controls("chkNewSheet")
    If chk Is Nothing Then Set chk = Me.Controls.Add("Forms.CheckBox.1", "chkNewSheet", True)
    With chk
        .Caption = "New sheet"
        .Left = leftValue
        .Top = topValue
        .Width = widthValue
        .Height = 18
        .Value = False
        .ControlTipText = "Write the result to a new worksheet at A1 instead of the selected output cell."
        .Visible = True
    End With
    On Error GoTo 0
End Sub

Private Sub ConfigureGeneratedForm()
    Dim i As Long
    Dim rowTop As Single
    Dim rowHeight As Single
    Dim visibleParams As Long
    Dim bottomTop As Single
    Dim buttonTop As Single
    Dim descTop As Single
    Dim margin As Single
    Dim labelWidth As Single
    Dim inputLeft As Single
    Dim inputWidth As Single
    Dim selectLeft As Single
    Dim formWidth As Single

    Me.Caption = BI_DisplayName()
    BI_ApplySelectionDefaults
    BI_HideLegacyUsageLabels

    margin = 10
    rowTop = 38
    rowHeight = 24
    visibleParams = PARAM_COUNT
    formWidth = BI_CompactFormWidth()
    labelWidth = 92 + (BI_LongestLabelLength() * 3.2)
    If labelWidth < 110 Then labelWidth = 110
    If labelWidth > 220 Then labelWidth = 220
    inputLeft = margin + labelWidth + 8
    selectLeft = formWidth - margin - 54
    inputWidth = selectLeft - inputLeft - 6
    If inputWidth < 190 Then inputWidth = 190

    BI_EnsureHeader
    On Error Resume Next
    Me.Controls("lblBIHeader").Width = formWidth - (margin * 2)
    On Error GoTo 0

    For i = 1 To 12
        BI_SetVisible "lblArg" & CStr(i), (i <= visibleParams)
        BI_SetVisible "txtArg" & CStr(i), (i <= visibleParams)
        BI_SetVisible "cmdArg" & CStr(i), (i <= visibleParams)
        If i <= visibleParams Then
            BI_SetLabel "lblArg" & CStr(i), BI_ParamLabel(i)
            BI_MoveControl "lblArg" & CStr(i), margin, rowTop + (i - 1) * rowHeight + 2, labelWidth
            If BI_ArgumentUsesSelector(i) Then
                BI_SetVisible "txtArg" & CStr(i), False
                BI_SetVisible "cmdArg" & CStr(i), False
                BI_EnsureArgumentSelector i, inputLeft, rowTop + (i - 1) * rowHeight, inputWidth
            Else
                BI_SetVisible "cboArg" & CStr(i), False
                BI_MoveControl "txtArg" & CStr(i), inputLeft, rowTop + (i - 1) * rowHeight, inputWidth
                BI_MoveControl "cmdArg" & CStr(i), selectLeft, rowTop + (i - 1) * rowHeight - 1, 54
                On Error Resume Next
                Me.Controls("cmdArg" & CStr(i)).Caption = "..."
                On Error GoTo 0
            End If
        End If
    Next i

    bottomTop = rowTop + visibleParams * rowHeight + 8
    BI_SetLabel "lblOutput", "Output"
    BI_MoveControl "lblOutput", margin, bottomTop + 2, labelWidth
    BI_MoveControl "txtOutput", inputLeft, bottomTop, inputWidth
    BI_MoveControl "cmdOutput", selectLeft, bottomTop - 1, 54
    On Error Resume Next
    Me.Controls("cmdOutput").Caption = "..."
    On Error GoTo 0
    BI_EnsureNewSheetOption inputLeft, bottomTop + 25, inputWidth

    buttonTop = bottomTop + 49
    BI_MoveControl "cmdRun", formWidth - margin - 174, buttonTop, 78
    BI_MoveControl "cmdCancel", formWidth - margin - 86, buttonTop, 76
    Me.Controls("cmdCancel").Caption = "Close"
    descTop = buttonTop + 34
    BI_EnsureDescription descTop
    On Error Resume Next
    Me.Controls("lblBIDescription").Left = margin
    Me.Controls("lblBIDescription").Width = formWidth - (margin * 2)
    On Error GoTo 0

    Me.Width = formWidth
    Me.Height = descTop + BI_DescriptionHeight() + 28
End Sub

Private Sub UserForm_Initialize()
    ConfigureGeneratedForm
End Sub
Private Sub PickRangeInto(ByVal targetTextBox As Object, ByVal prompt As String)
    On Error GoTo ErrHandler
    Dim picked As Range
    Me.Hide
    Set picked = Application.InputBox(prompt, FUNCTION_NAME, Type:=8)
    If Not picked Is Nothing Then targetTextBox.Text = picked.Address(External:=True)
CleanExit:
    Me.Show
    Exit Sub
ErrHandler:
    Resume CleanExit
End Sub
Private Sub cmdOutput_Click()
    PickRangeInto Me.txtOutput, "Select output top-left cell"
End Sub
Private Function TryParseRange(ByVal rawText As String, ByRef resultRange As Range) As Boolean
    On Error GoTo NotRange
    Set resultRange = Application.Range(Trim$(rawText))
    TryParseRange = Not resultRange Is Nothing
    Exit Function
NotRange:
    Set resultRange = Nothing
    TryParseRange = False
End Function
Private Function ParseScalarArgument(ByVal rawText As String) As Variant
    Dim s As String
    s = Trim$(rawText)
    If Len(s) = 0 Then
        ParseScalarArgument = Empty
    ElseIf UCase$(s) = "TRUE" Then
        ParseScalarArgument = True
    ElseIf UCase$(s) = "FALSE" Then
        ParseScalarArgument = False
    ElseIf IsNumeric(s) Then
        ParseScalarArgument = CDbl(s)
    ElseIf IsDate(s) Then
        ParseScalarArgument = CDate(s)
    Else
        ParseScalarArgument = s
    End If
End Function
Private Sub WriteResult(ByVal outputCell As Range, ByVal result As Variant)
    If IsArray(result) Then
        outputCell.Resize(UBound(result, 1), UBound(result, 2)).Value = result
        outputCell.CurrentRegion.Columns.AutoFit
    Else
        outputCell.Value = result
    End If
End Sub

Private Sub WriteCellWiseNewSheetReport(ByVal outputCell As Range, ByVal inputRange As Range, ByVal result As Variant)
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim outRow As Long
    With outputCell.Worksheet
        outputCell.Value = BI_DisplayName()
        outputCell.Offset(1, 0).Value = "Input"
        outputCell.Offset(1, 1).Value = "Output"
        outputCell.Resize(1, 2).Font.Bold = True
        outputCell.Offset(1, 0).Resize(1, 2).Font.Bold = True
        outRow = 2
        For rowIndex = 1 To inputRange.Rows.Count
            For colIndex = 1 To inputRange.Columns.Count
                outputCell.Offset(outRow, 0).Value = inputRange.Cells(rowIndex, colIndex).Value
                outputCell.Offset(outRow, 1).Value = result(rowIndex, colIndex)
                outRow = outRow + 1
            Next colIndex
        Next rowIndex
        outputCell.CurrentRegion.Columns.AutoFit
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function BI_IsCellWiseFunction() As Boolean
    Select Case FUNCTION_NAME
        Case "BI_BeginsWith", "BI_Contains", "BI_EndsWith", "BI_ExtractNumbers", "BI_ExtractText", _
             "BI_FinancialYear", "BI_FinancialYearEnd", "BI_FinancialYearStart", "BI_FindOneOf", _
             "BI_Fuzzy", "BI_IndianNum2Word", "BI_Insert", "BI_IsCharAtoZ", "BI_IsLeap", _
             "BI_IsValidPAN", "BI_LuhnAlgorithm", "BI_NetworkDays_Indian", "BI_Num2Word", _
             "BI_Quarter", "BI_QuarterEnd", "BI_QuarterStart", "BI_ReverseText", "BI_WordCount", _
             "BI_WorkDay_Indian"
            BI_IsCellWiseFunction = True
    End Select
End Function

Private Function BI_IsVectorArgument(ByVal argIndex As Long) As Boolean
    Select Case FUNCTION_NAME
        Case "BI_NetworkDays_Indian"
            BI_IsVectorArgument = (argIndex = 1 Or argIndex = 2)
        Case "BI_WorkDay_Indian"
            BI_IsVectorArgument = (argIndex = 1 Or argIndex = 2)
        Case "BI_BeginsWith", "BI_Contains", "BI_EndsWith"
            BI_IsVectorArgument = (argIndex = 1 Or argIndex = 2)
        Case "BI_FindOneOf", "BI_Fuzzy"
            BI_IsVectorArgument = (argIndex = 1 Or argIndex = 2)
        Case "BI_Insert"
            BI_IsVectorArgument = (argIndex = 1)
        Case Else
            BI_IsVectorArgument = (argIndex = 1)
    End Select
End Function

Private Function BI_FirstMappableRange(ByRef argRanges() As Range) As Range
    Dim i As Long
    If Not BI_IsCellWiseFunction() Then Exit Function
    For i = 1 To PARAM_COUNT
        If BI_IsVectorArgument(i) Then
            If Not argRanges(i) Is Nothing Then
                If argRanges(i).Cells.CountLarge > 1 Then
                    Set BI_FirstMappableRange = argRanges(i)
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Private Function BI_RangeCellValue(ByVal sourceRange As Range, ByVal templateRange As Range, ByVal rowIndex As Long, ByVal colIndex As Long) As Variant
    If sourceRange.Cells.CountLarge = 1 Then
        BI_RangeCellValue = sourceRange.Cells(1, 1).Value
    ElseIf sourceRange.Rows.Count = templateRange.Rows.Count And sourceRange.Columns.Count = templateRange.Columns.Count Then
        BI_RangeCellValue = sourceRange.Cells(rowIndex, colIndex).Value
    ElseIf sourceRange.Rows.Count = templateRange.Rows.Count And sourceRange.Columns.Count = 1 Then
        BI_RangeCellValue = sourceRange.Cells(rowIndex, 1).Value
    ElseIf sourceRange.Rows.Count = 1 And sourceRange.Columns.Count = templateRange.Columns.Count Then
        BI_RangeCellValue = sourceRange.Cells(1, colIndex).Value
    Else
        Err.Raise vbObjectError + 513, FUNCTION_NAME, "Vector input ranges must match the selected range shape, or be a single row/column aligned with it."
    End If
End Function



Private Function BI_RunWithArgs(ByVal argc As Long, ByRef args() As Variant) As Variant
    Select Case argc
        Case 0: BI_RunWithArgs = Application.Run(FUNCTION_NAME)
        Case 1: BI_RunWithArgs = Application.Run(FUNCTION_NAME, args(1))
        Case 2: BI_RunWithArgs = Application.Run(FUNCTION_NAME, args(1), args(2))
        Case 3: BI_RunWithArgs = Application.Run(FUNCTION_NAME, args(1), args(2), args(3))
        Case 4: BI_RunWithArgs = Application.Run(FUNCTION_NAME, args(1), args(2), args(3), args(4))
        Case 5: BI_RunWithArgs = Application.Run(FUNCTION_NAME, args(1), args(2), args(3), args(4), args(5))
        Case 6: BI_RunWithArgs = Application.Run(FUNCTION_NAME, args(1), args(2), args(3), args(4), args(5), args(6))
        Case 7: BI_RunWithArgs = Application.Run(FUNCTION_NAME, args(1), args(2), args(3), args(4), args(5), args(6), args(7))
        Case 8: BI_RunWithArgs = Application.Run(FUNCTION_NAME, args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8))
        Case 9: BI_RunWithArgs = Application.Run(FUNCTION_NAME, args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9))
        Case 10: BI_RunWithArgs = Application.Run(FUNCTION_NAME, args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10))
        Case 11: BI_RunWithArgs = Application.Run(FUNCTION_NAME, args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11))
        Case 12: BI_RunWithArgs = Application.Run(FUNCTION_NAME, args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12))
    End Select
End Function

Private Function BI_RunCellWise(ByVal argc As Long, ByRef args() As Variant, ByRef argRanges() As Range, ByVal templateRange As Range) As Variant
    Dim outputValues() As Variant
    Dim callArgs(1 To 12) As Variant
    Dim r As Long
    Dim c As Long
    Dim i As Long
    ReDim outputValues(1 To templateRange.Rows.Count, 1 To templateRange.Columns.Count)
    For r = 1 To templateRange.Rows.Count
        For c = 1 To templateRange.Columns.Count
            For i = 1 To PARAM_COUNT
                If BI_IsVectorArgument(i) And Not argRanges(i) Is Nothing Then
                    callArgs(i) = BI_RangeCellValue(argRanges(i), templateRange, r, c)
                ElseIf IsObject(args(i)) Then
                    Set callArgs(i) = args(i)
                Else
                    callArgs(i) = args(i)
                End If
            Next i
            outputValues(r, c) = BI_RunWithArgs(argc, callArgs)
        Next c
    Next r
    BI_RunCellWise = outputValues
End Function

Private Sub RunFunction()
    On Error GoTo ErrHandler
    Dim args(1 To 12) As Variant
    Dim argRanges(1 To 12) As Range
    Dim rng As Range
    Dim mapRange As Range
    Dim result As Variant
    Dim outputCell As Range
    Dim i As Long
    Dim argc As Long
    Dim rawText As String

    For i = 1 To PARAM_COUNT
        If BI_ArgumentUsesSelector(i) Then
            args(i) = BI_SelectorValue(i)
            argc = i
        Else
            rawText = Trim$(Me.Controls("txtArg" & i).Text)
            If Len(rawText) > 0 Then argc = i
            If TryParseRange(rawText, rng) Then
                Set argRanges(i) = rng
                Set args(i) = rng
            Else
                args(i) = ParseScalarArgument(rawText)
            End If
        End If
    Next i

    Set mapRange = BI_FirstMappableRange(argRanges)
    If Not mapRange Is Nothing Then
        result = BI_RunCellWise(argc, args, argRanges, mapRange)
    Else
        result = BI_RunWithArgs(argc, args)
    End If

    Set outputCell = BI_ResolveOutputCell()
    If BI_UseNewSheet() And Not mapRange Is Nothing Then
        WriteCellWiseNewSheetReport outputCell, mapRange, result
    Else
        WriteResult outputCell, result
    End If
    If FUNCTION_NAME = "BI_AnsCombeQuartet" Then BI_CreateAnscombeQuartetChart outputCell
    Unload Me
    Exit Sub
ErrHandler:
    MsgBox FUNCTION_NAME & " failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub
Private Sub cmdRun_Click()
    RunFunction
End Sub
Private Sub cmdArg1_Click()
    PickRangeInto Me.txtArg1, "Select argument 1 range"
End Sub

Private Sub cmdArg2_Click()
    PickRangeInto Me.txtArg2, "Select argument 2 range"
End Sub

Private Sub cmdArg3_Click()
    PickRangeInto Me.txtArg3, "Select argument 3 range"
End Sub

Private Sub cmdArg4_Click()
    PickRangeInto Me.txtArg4, "Select argument 4 range"
End Sub

