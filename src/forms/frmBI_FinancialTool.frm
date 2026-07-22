VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBI_FinancialTool 
   Caption         =   "BeIndian Financial Tool"
   ClientHeight    =   4440
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8380.001
   OleObjectBlob   =   "frmBI_FinancialTool.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBI_FinancialTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.cmdCancel.Caption = "Close"
    If Len(Me.Tag) = 0 Then Me.Tag = "Payback"
    Me.txtTool.Text = Me.Tag
    Me.txtTool.Locked = True
    Me.txtOutput.Text = ActiveCell.Address(External:=True)
    If TypeName(Selection) = "Range" Then Me.txtCashFlows.Text = Selection.Address(External:=True)
    ConfigureTool
End Sub

Private Sub ConfigureTool()
    Dim isDep As Boolean, isDiscounted As Boolean
    isDep = (UCase$(Me.txtTool.Text) = "SLN" Or UCase$(Me.txtTool.Text) = "WDV")
    isDiscounted = (UCase$(Me.txtTool.Text) = "DISCOUNTEDPAYBACK")
    Me.lblCashFlows.Visible = Not isDep: Me.txtCashFlows.Visible = Not isDep: Me.cmdPickCashFlows.Visible = Not isDep
    Me.lblRate.Visible = isDiscounted: Me.txtRate.Visible = isDiscounted
    Me.lblCost.Visible = isDep: Me.txtCost.Visible = isDep
    Me.lblLife.Visible = isDep: Me.txtLife.Visible = isDep
    Me.lblScrap.Visible = isDep: Me.txtScrap.Visible = isDep
End Sub

Private Sub cmdPickCashFlows_Click()
    PickRangeInto Me.txtCashFlows, "Select cash flow range"
End Sub

Private Sub cmdPickOutput_Click()
    PickRangeInto Me.txtOutput, "Select output top-left cell"
End Sub

Private Sub PickRangeInto(ByVal targetTextBox As Object, ByVal prompt As String)
    On Error GoTo ErrHandler
    Dim picked As Range
    Me.Hide
    Set picked = Application.InputBox(prompt, "BeIndian Financial Tool", Type:=8)
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
    Dim result As Variant, outputCell As Range
    Set outputCell = Application.Range(Me.txtOutput.Text).Cells(1, 1)
    Select Case UCase$(Me.txtTool.Text)
        Case "PAYBACK"
            result = BI_PaybackPeriod(Application.Range(Me.txtCashFlows.Text))
        Case "DISCOUNTEDPAYBACK"
            result = BI_DiscountedPayback(Application.Range(Me.txtCashFlows.Text), CDbl(Me.txtRate.Text))
        Case "SLN"
            result = BI_SLNSchedule(CDbl(Me.txtCost.Text), CLng(Me.txtLife.Text), CDbl(Me.txtScrap.Text))
        Case "WDV"
            result = BI_WDVSchedule(CDbl(Me.txtCost.Text), CLng(Me.txtLife.Text), CDbl(Me.txtScrap.Text))
    End Select
    If IsArray(result) Then
        outputCell.Resize(UBound(result, 1), UBound(result, 2)).value = result
    Else
        outputCell.value = result
    End If
    outputCell.CurrentRegion.Columns.AutoFit
    Unload Me
    Exit Sub
ErrHandler:
    MsgBox "Financial tool failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub
