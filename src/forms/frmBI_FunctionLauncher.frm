VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBI_FunctionLauncher 
   Caption         =   "BeIndian Function Launcher"
   ClientHeight    =   8040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   10980
   OleObjectBlob   =   "frmBI_FunctionLauncher.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBI_FunctionLauncher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mFunctions As Variant

Private Function IsLauncherHidden(ByVal functionName As String) As Boolean
    IsLauncherHidden = (UCase$(functionName) = "BI_ABOUT")
End Function

Private Sub UserForm_Initialize()
    Dim i As Long
    mFunctions = BI_LauncherFunctionList()
    Me.cmbFunction.Clear
    Me.cmbFunction.ColumnCount = 2
    Me.cmbFunction.ColumnWidths = "180 pt;0 pt"
    For i = LBound(mFunctions, 1) To UBound(mFunctions, 1)
        If Not IsLauncherHidden(CStr(mFunctions(i, 1))) Then
            Me.cmbFunction.AddItem CStr(mFunctions(i, 1))
            Me.cmbFunction.List(Me.cmbFunction.ListCount - 1, 1) = CStr(i)
        End If
    Next i
    Me.txtOutput.Text = ActiveCell.Address(External:=True)
    If Len(Me.Tag) > 0 Then
        For i = 0 To Me.cmbFunction.ListCount - 1
            If Me.cmbFunction.List(i) = Me.Tag Then
                Me.cmbFunction.ListIndex = i
                Exit For
            End If
        Next i
    End If
    If Me.cmbFunction.ListIndex < 0 And Me.cmbFunction.ListCount > 0 Then Me.cmbFunction.ListIndex = 0
End Sub

Private Sub cmbFunction_Change()
    Dim i As Long
    If Me.cmbFunction.ListIndex < 0 Then Exit Sub
    i = CLng(Me.cmbFunction.List(Me.cmbFunction.ListIndex, 1))
    Me.txtSyntax.Text = CStr(mFunctions(i, 2))
    Me.txtDescription.Text = CStr(mFunctions(i, 3))
End Sub

Private Sub cmdPickRange_Click()
    On Error GoTo ErrHandler
    Dim picked As Range
    Me.Hide
    Set picked = Application.InputBox("Select range argument", "BeIndian Function Launcher", Type:=8)
    If Not picked Is Nothing Then
        If Len(Me.txtArgs.Text) > 0 Then Me.txtArgs.Text = Me.txtArgs.Text & vbCrLf
        Me.txtArgs.Text = Me.txtArgs.Text & picked.Address(External:=True)
    End If
CleanExit:
    Me.Show
    Exit Sub
ErrHandler:
    Resume CleanExit
End Sub

Private Sub cmdPickOutput_Click()
    On Error GoTo ErrHandler
    Dim picked As Range
    Me.Hide
    Set picked = Application.InputBox("Select output top-left cell", "BeIndian Function Launcher", Type:=8)
    If Not picked Is Nothing Then Me.txtOutput.Text = picked.Cells(1, 1).Address(External:=True)
CleanExit:
    Me.Show
    Exit Sub
ErrHandler:
    Resume CleanExit
End Sub

Private Sub cmdClearArgs_Click()
    Me.txtArgs.Text = vbNullString
End Sub

Private Sub cmdCancel_Click()
    Unload Me
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
        outputCell.Resize(UBound(result, 1), UBound(result, 2)).value = result
        outputCell.CurrentRegion.Columns.AutoFit
    Else
        outputCell.value = result
    End If
End Sub

Private Sub cmdRun_Click()
    On Error GoTo ErrHandler
    Dim fn As String
    Dim lines() As String
    Dim args(1 To 12) As Variant
    Dim argc As Long, i As Long
    Dim result As Variant
    Dim outputCell As Range
    Dim parsedRange As Range

    If Me.cmbFunction.ListIndex < 0 Then Exit Sub
    fn = Me.cmbFunction.value
    If Len(Trim$(Me.txtArgs.Text)) > 0 Then
        lines = Split(Replace(Me.txtArgs.Text, vbCrLf, vbLf), vbLf)
        For i = LBound(lines) To UBound(lines)
            If Len(Trim$(lines(i))) > 0 Then
                argc = argc + 1
                If argc > 12 Then Err.Raise vbObjectError + 810, "frmBI_FunctionLauncher", "Maximum 12 arguments supported."
                If TryParseRange(lines(i), parsedRange) Then
                    Set args(argc) = parsedRange
                Else
                    args(argc) = ParseScalarArgument(lines(i))
                End If
            End If
        Next i
    End If

    Select Case argc
        Case 0: result = Application.Run(fn)
        Case 1: result = Application.Run(fn, args(1))
        Case 2: result = Application.Run(fn, args(1), args(2))
        Case 3: result = Application.Run(fn, args(1), args(2), args(3))
        Case 4: result = Application.Run(fn, args(1), args(2), args(3), args(4))
        Case 5: result = Application.Run(fn, args(1), args(2), args(3), args(4), args(5))
        Case 6: result = Application.Run(fn, args(1), args(2), args(3), args(4), args(5), args(6))
        Case 7: result = Application.Run(fn, args(1), args(2), args(3), args(4), args(5), args(6), args(7))
        Case 8: result = Application.Run(fn, args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8))
        Case 9: result = Application.Run(fn, args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9))
        Case 10: result = Application.Run(fn, args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10))
        Case 11: result = Application.Run(fn, args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11))
        Case 12: result = Application.Run(fn, args(1), args(2), args(3), args(4), args(5), args(6), args(7), args(8), args(9), args(10), args(11), args(12))
    End Select

    Set outputCell = Application.Range(Me.txtOutput.Text).Cells(1, 1)
    WriteResult outputCell, result
    Exit Sub
ErrHandler:
    MsgBox "Function execution failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub
