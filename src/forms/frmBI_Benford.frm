VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBI_Benford 
   Caption         =   "BeIndian Benford Analysis"
   ClientHeight    =   4140
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7580
   OleObjectBlob   =   "frmBI_Benford.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBI_Benford"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.cmdCancel.Caption = "Close"
    Me.cmbTest.Clear
    Me.cmbTest.AddItem "First Digit"
    Me.cmbTest.AddItem "First Two Digits"
    Me.cmbTest.AddItem "First Three Digits"
    Me.cmbTest.AddItem "Second Digit"
    Me.cmbTest.AddItem "Third Digit"
    Me.cmbTest.AddItem "Last Two Digits"
    Me.cmbTest.AddItem "Second Order"
    Me.cmbTest.AddItem "Summary/Summation Test"
    Me.cmbTest.ListIndex = 0
    Me.chkChart.value = True
    If TypeName(Selection) = "Range" Then Me.txtRange.Text = Selection.Address(External:=True)
    Me.txtOutput.Text = ActiveCell.Address(External:=True)
End Sub

Private Sub cmdPickRange_Click()
    PickRangeInto Me.txtRange, "Select the Benford input range"
End Sub

Private Sub cmdPickOutput_Click()
    PickRangeInto Me.txtOutput, "Select the output top-left cell"
End Sub

Private Sub PickRangeInto(ByVal targetTextBox As Object, ByVal prompt As String)
    On Error GoTo ErrHandler
    Dim picked As Range
    Me.Hide
    Set picked = Application.InputBox(prompt, "BeIndian Benford Analysis", Type:=8)
    If Not picked Is Nothing Then targetTextBox.Text = picked.Address(External:=True)
CleanExit:
    Me.Show
    Exit Sub
ErrHandler:
    Resume CleanExit
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRun_Click()
    On Error GoTo ErrHandler
    Dim inputRange As Range
    Dim outputCell As Range
    Dim outputRange As Range
    Dim result As Variant
    Dim chartTitle As String

    Set inputRange = Application.Range(Me.txtRange.Text)
    Set outputCell = Application.Range(Me.txtOutput.Text).Cells(1, 1)
    Select Case Me.cmbTest.value
        Case "First Digit"
            result = BI_BenfordLaw_FirstDigit(inputRange)
            chartTitle = "Benford Law First Digit - Actual vs Expected"
        Case "First Two Digits"
            result = BI_BenfordLaw_FirstTwoDigits(inputRange)
            chartTitle = "Benford Law First Two Digits - Actual vs Expected"
        Case "First Three Digits"
            result = BI_BenfordLaw_FirstThreeDigits(inputRange)
            chartTitle = "Benford Law First Three Digits - Actual vs Expected"
        Case "Second Digit"
            result = BI_BenfordLaw_SecondDigit(inputRange)
            chartTitle = "Benford Law Second Digit - Actual vs Expected"
        Case "Third Digit"
            result = BI_BenfordLaw_ThirdDigit(inputRange)
            chartTitle = "Benford Law Third Digit - Actual vs Expected"
        Case "Last Two Digits"
            result = BI_BenfordLaw_LastTwoDigits(inputRange)
            chartTitle = "Benford Law Last Two Digits - Actual vs Expected"
        Case "Second Order"
            result = BI_BenfordLaw_SecondOrder(inputRange)
            chartTitle = "Benford Law Second Order - Actual vs Expected"
        Case "Summary/Summation Test"
            result = BI_BenfordLaw_SummaryTest(inputRange)
            chartTitle = "Benford Law Summary/Summation Test - Actual vs Expected"
    End Select

    Set outputRange = outputCell.Resize(UBound(result, 1), UBound(result, 2))
    outputRange.value = result
    outputRange.Columns(3).NumberFormat = "0.00%"
    outputRange.Columns(4).NumberFormat = "0.00%"
    outputRange.Columns.AutoFit
    If Me.chkChart.value Then BI_CreateBenfordActualExpectedChart outputCell, chartTitle
    Unload Me
    Exit Sub
ErrHandler:
    MsgBox "Benford form failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub
