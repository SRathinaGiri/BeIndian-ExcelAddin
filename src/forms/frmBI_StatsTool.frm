VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBI_StatsTool 
   Caption         =   "BeIndian Statistics Tool"
   ClientHeight    =   4440
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8380.001
   OleObjectBlob   =   "frmBI_StatsTool.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBI_StatsTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    If Len(Me.Tag) = 0 Then Me.Tag = "Describe"
    Me.txtTool.Text = Me.Tag
    Me.txtTool.Locked = True
    Me.txtOutput.Text = ActiveCell.Address(External:=True)
    If TypeName(Selection) = "Range" Then Me.txtCashFlows.Text = Selection.Address(External:=True)
    ConfigureTool
End Sub

Private Sub ConfigureTool()
    Dim toolName As String
    toolName = UCase$(Me.txtTool.Text)

    Me.lblCashFlows.Caption = "Source range"
    Me.lblRate.Caption = "Value 1"
    Me.lblCost.Caption = "Range 2"
    Me.lblLife.Caption = "Option 1"
    Me.lblScrap.Caption = "Option 2"

    Me.lblCashFlows.Visible = True: Me.txtCashFlows.Visible = True: Me.cmdPickCashFlows.Visible = True
    Me.lblRate.Visible = False: Me.txtRate.Visible = False
    Me.lblCost.Visible = False: Me.txtCost.Visible = False
    Me.lblLife.Visible = False: Me.txtLife.Visible = False
    Me.lblScrap.Visible = False: Me.txtScrap.Visible = False

    Select Case toolName
        Case "DESCRIBE"
            Me.lblCashFlows.Caption = "Numeric range"
        Case "CORRELATION", "COVARIANCE"
            Me.lblCashFlows.Caption = "Numeric table"
            Me.lblLife.Visible = True: Me.txtLife.Visible = True
            Me.lblLife.Caption = "Labels (TRUE/FALSE)"
            If Len(Me.txtLife.Text) = 0 Then Me.txtLife.Text = "TRUE"
        Case "REGRESSION"
            Me.lblCashFlows.Caption = "Dependent range (Y)"
            Me.lblCost.Visible = True: Me.txtCost.Visible = True
            Me.lblCost.Caption = "Independent range (X)"
            Me.lblRate.Visible = True: Me.txtRate.Visible = True
            Me.lblRate.Caption = "Confidence (0-1)"
            Me.lblLife.Visible = True: Me.txtLife.Visible = True
            Me.lblLife.Caption = "Labels (TRUE/FALSE)"
            If Len(Me.txtRate.Text) = 0 Then Me.txtRate.Text = "0.95"
            If Len(Me.txtLife.Text) = 0 Then Me.txtLife.Text = "TRUE"
        Case "HISTOGRAM"
            Me.lblCashFlows.Caption = "Numeric range"
            Me.lblRate.Visible = True: Me.txtRate.Visible = True
            Me.lblRate.Caption = "Bin start"
            Me.lblLife.Visible = True: Me.txtLife.Visible = True
            Me.lblLife.Caption = "Bin increment"
        Case "MOVINGAVERAGE"
            Me.lblCashFlows.Caption = "Numeric range"
            Me.lblRate.Visible = True: Me.txtRate.Visible = True
            Me.lblRate.Caption = "Interval"
            Me.lblLife.Visible = True: Me.txtLife.Visible = True
            Me.lblLife.Caption = "Labels (TRUE/FALSE)"
            Me.lblScrap.Visible = True: Me.txtScrap.Visible = True
            Me.lblScrap.Caption = "Std Error (TRUE/FALSE)"
            If Len(Me.txtRate.Text) = 0 Then Me.txtRate.Text = "3"
            If Len(Me.txtLife.Text) = 0 Then Me.txtLife.Text = "FALSE"
            If Len(Me.txtScrap.Text) = 0 Then Me.txtScrap.Text = "FALSE"
        Case "EXPONENTIALSMOOTHING"
            Me.lblCashFlows.Caption = "Numeric range"
            Me.lblRate.Visible = True: Me.txtRate.Visible = True
            Me.lblRate.Caption = "Damping factor"
            Me.lblLife.Visible = True: Me.txtLife.Visible = True
            Me.lblLife.Caption = "Labels (TRUE/FALSE)"
            If Len(Me.txtRate.Text) = 0 Then Me.txtRate.Text = "0.3"
            If Len(Me.txtLife.Text) = 0 Then Me.txtLife.Text = "FALSE"
        Case "ANOVASINGLE"
            Me.lblCashFlows.Caption = "Grouped table"
            Me.lblRate.Visible = True: Me.txtRate.Visible = True
            Me.lblRate.Caption = "Alpha"
            Me.lblLife.Visible = True: Me.txtLife.Visible = True
            Me.lblLife.Caption = "Labels (TRUE/FALSE)"
            If Len(Me.txtRate.Text) = 0 Then Me.txtRate.Text = "0.05"
            If Len(Me.txtLife.Text) = 0 Then Me.txtLife.Text = "TRUE"
        Case "ANOVATWONOREPLICATION"
            Me.lblCashFlows.Caption = "Matrix / table"
            Me.lblRate.Visible = True: Me.txtRate.Visible = True
            Me.lblRate.Caption = "Alpha"
            Me.lblLife.Visible = True: Me.txtLife.Visible = True
            Me.lblLife.Caption = "Labels (TRUE/FALSE)"
            If Len(Me.txtRate.Text) = 0 Then Me.txtRate.Text = "0.05"
            If Len(Me.txtLife.Text) = 0 Then Me.txtLife.Text = "TRUE"
        Case "ANSCOMBEQUARTET"
            Me.lblCashFlows.Visible = False: Me.txtCashFlows.Visible = False: Me.cmdPickCashFlows.Visible = False
            Me.lblRate.Visible = False: Me.txtRate.Visible = False
            Me.lblCost.Visible = False: Me.txtCost.Visible = False
            Me.lblLife.Visible = False: Me.txtLife.Visible = False
            Me.lblScrap.Visible = False: Me.txtScrap.Visible = False
    End Select
End Sub

Private Sub cmdPickCashFlows_Click()
    PickRangeInto Me.txtCashFlows, "Select source range"
End Sub

Private Sub cmdPickOutput_Click()
    PickRangeInto Me.txtOutput, "Select output top-left cell"
End Sub

Private Sub txtTool_Change()
    ConfigureTool
End Sub

Private Function ParseBooleanText(ByVal textValue As String, Optional ByVal defaultValue As Boolean = False) As Boolean
    Dim normalized As String
    normalized = UCase$(Trim$(textValue))
    If Len(normalized) = 0 Then
        ParseBooleanText = defaultValue
    Else
        ParseBooleanText = (normalized = "TRUE" Or normalized = "YES" Or normalized = "Y" Or normalized = "1")
    End If
End Function

Private Sub PickRangeInto(ByVal targetTextBox As Object, ByVal prompt As String)
    On Error GoTo ErrHandler
    Dim picked As Range
    Me.Hide
    Set picked = Application.InputBox(prompt, "BeIndian Statistics Tool", Type:=8)
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
    Dim result As Variant
    Dim outputCell As Range
    Dim toolName As String

    toolName = UCase$(Me.txtTool.Text)
    Set outputCell = Application.Range(Me.txtOutput.Text).Cells(1, 1)

    Select Case toolName
        Case "DESCRIBE"
            result = BI_Describe(Application.Range(Me.txtCashFlows.Text))
        Case "CORRELATION"
            result = BI_Correlation(Application.Range(Me.txtCashFlows.Text), ParseBooleanText(Me.txtLife.Text, True))
        Case "COVARIANCE"
            result = BI_Covariance(Application.Range(Me.txtCashFlows.Text), ParseBooleanText(Me.txtLife.Text, True))
        Case "REGRESSION"
            result = BI_Regression(Application.Range(Me.txtCashFlows.Text), Application.Range(Me.txtCost.Text), ParseBooleanText(Me.txtLife.Text, True), CDbl(Me.txtRate.Text))
        Case "HISTOGRAM"
            If Len(Trim$(Me.txtRate.Text)) = 0 And Len(Trim$(Me.txtLife.Text)) = 0 Then
                result = BI_Histogram(Application.Range(Me.txtCashFlows.Text))
            ElseIf Len(Trim$(Me.txtLife.Text)) = 0 Then
                result = BI_Histogram(Application.Range(Me.txtCashFlows.Text), CDbl(Me.txtRate.Text))
            Else
                result = BI_Histogram(Application.Range(Me.txtCashFlows.Text), CDbl(Me.txtRate.Text), CDbl(Me.txtLife.Text))
            End If
        Case "MOVINGAVERAGE"
            result = BI_MovingAverage(Application.Range(Me.txtCashFlows.Text), CLng(Me.txtRate.Text), ParseBooleanText(Me.txtLife.Text, False), ParseBooleanText(Me.txtScrap.Text, False))
        Case "EXPONENTIALSMOOTHING"
            result = BI_ExponentialSmoothing(Application.Range(Me.txtCashFlows.Text), CDbl(Me.txtRate.Text), ParseBooleanText(Me.txtLife.Text, False))
        Case "ANOVASINGLE"
            result = BI_AnovaSingleFactor(Application.Range(Me.txtCashFlows.Text), CDbl(Me.txtRate.Text), ParseBooleanText(Me.txtLife.Text, True))
        Case "ANOVATWONOREPLICATION"
            result = BI_AnovaTwoFactorsWithoutReplication(Application.Range(Me.txtCashFlows.Text), CDbl(Me.txtRate.Text), ParseBooleanText(Me.txtLife.Text, True))
        Case "ANSCOMBEQUARTET"
            result = BI_AnsCombeQuartet()
        Case Else
            Err.Raise vbObjectError + 610, "frmBI_StatsTool", "Unknown statistics tool."
    End Select

    If IsArray(result) Then
        outputCell.Resize(UBound(result, 1), UBound(result, 2)).Value = result
    Else
        outputCell.Value = result
    End If
    outputCell.CurrentRegion.Columns.AutoFit
    If toolName = "ANSCOMBEQUARTET" Then BI_CreateAnscombeQuartetChart outputCell
    Unload Me
    Exit Sub
ErrHandler:
    MsgBox "Statistics tool failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub
