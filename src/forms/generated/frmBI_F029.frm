VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBI_F029 
   Caption         =   "BI_Correlation"
   ClientHeight    =   4640
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   10180
   OleObjectBlob   =   "frmBI_F029.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBI_F029"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const FUNCTION_NAME As String = "BI_Correlation"
Private Const PARAM_COUNT As Long = 2
Private Const PARAM_LABELS As String = "Data|Labels"
Private Const DESCRIPTION_TEXT As String = "Returns a Correlation Table of many attributes. This is one of the Data Analysis Toolpak features."

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
        controlName = "cmdRun" Or controlName = "cmdCancel" Or _
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

Private Sub ConfigureGeneratedForm()
    Dim i As Long
    Dim rowTop As Single
    Dim rowHeight As Single
    Dim visibleParams As Long
    Dim bottomTop As Single
    Dim buttonTop As Single
    Dim descTop As Single

    Me.Caption = BI_DisplayName()
    Me.txtOutput.Text = ActiveCell.Address(External:=True)
    BI_HideLegacyUsageLabels
    BI_EnsureHeader

    rowTop = 48
    rowHeight = 30
    visibleParams = PARAM_COUNT
    For i = 1 To 12
        BI_SetVisible "lblArg" & CStr(i), (i <= visibleParams)
        BI_SetVisible "txtArg" & CStr(i), (i <= visibleParams)
        BI_SetVisible "cmdArg" & CStr(i), (i <= visibleParams)
        If i <= visibleParams Then
            BI_SetLabel "lblArg" & CStr(i), BI_ParamLabel(i)
            BI_MoveControl "lblArg" & CStr(i), 12, rowTop + (i - 1) * rowHeight, 130
            BI_MoveControl "txtArg" & CStr(i), 150, rowTop + (i - 1) * rowHeight - 2, 330
            BI_MoveControl "cmdArg" & CStr(i), 490, rowTop + (i - 1) * rowHeight - 3, 70
            On Error Resume Next
            Me.Controls("cmdArg" & CStr(i)).Caption = "Select"
            If i = 1 And TypeName(Selection) = "Range" Then Me.Controls("txtArg1").Text = Selection.Address(External:=True)
            On Error GoTo 0
        End If
    Next i

    bottomTop = rowTop + visibleParams * rowHeight + 14
    BI_SetLabel "lblOutput", "Output"
    BI_MoveControl "lblOutput", 12, bottomTop, 130
    BI_MoveControl "txtOutput", 150, bottomTop - 2, 330
    BI_MoveControl "cmdOutput", 490, bottomTop - 3, 70
    On Error Resume Next
    Me.Controls("cmdOutput").Caption = "Select"
    buttonTop = bottomTop + 42
    BI_MoveControl "cmdRun", 350, buttonTop, 90
    BI_MoveControl "cmdCancel", 450, buttonTop, 90
    descTop = buttonTop + 46
    BI_EnsureDescription descTop
    Me.Height = descTop + BI_DescriptionHeight() + 42
    Me.Width = 600
    On Error GoTo 0
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
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub RunFunction()
    On Error GoTo ErrHandler
    Dim args(1 To 12) As Variant
    Dim rng As Range
    Dim result As Variant
    Dim outputCell As Range
    Dim i As Long
    Dim argc As Long
    Dim rawText As String

    For i = 1 To PARAM_COUNT
        rawText = Trim$(Me.Controls("txtArg" & i).Text)
        If Len(rawText) > 0 Then argc = i
        If TryParseRange(rawText, rng) Then
            Set args(i) = rng
        Else
            args(i) = ParseScalarArgument(rawText)
        End If
    Next i

    Select Case argc
        Case 0: result = Application.Run(FUNCTION_NAME)
        Case 1: result = Application.Run(FUNCTION_NAME, args(1))
        Case 2: result = Application.Run(FUNCTION_NAME, args(1), args(2))
        Case 3: result = Application.Run(FUNCTION_NAME, args(1), args(2), args(3))
        Case 4: result = Application.Run(FUNCTION_NAME, args(1), args(2), args(3), args(4))
        Case 5: result = Application.Run(FUNCTION_NAME, args(1), args(2), args(3), args(4), args(5))
        Case 6: result = Application.Run(FUNCTION_NAME, args(1), args(2), args(3), args(4), args(5), args(6))
        Case 7: result = Application.Run(FUNCTION_NAME, args(1), args(2), args(3), args(4), args(5), args(6), args(7))
        Case 8: result = Application.Run(FUNCTION_NAME, args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8))
        Case 9: result = Application.Run(FUNCTION_NAME, args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9))
        Case 10: result = Application.Run(FUNCTION_NAME, args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10))
        Case 11: result = Application.Run(FUNCTION_NAME, args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11))
        Case 12: result = Application.Run(FUNCTION_NAME, args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12))
    End Select
    Set outputCell = Application.Range(Me.txtOutput.Text).Cells(1, 1)
    WriteResult outputCell, result
    If FUNCTION_NAME = "BI_AnsCombeQuartet" Then BI_CreateAnscombeQuartetChart outputCell
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

