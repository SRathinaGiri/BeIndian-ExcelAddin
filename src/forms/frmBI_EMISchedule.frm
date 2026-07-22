VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBI_EMISchedule 
   Caption         =   "BeIndian EMI Schedule"
   ClientHeight    =   3940
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7380
   OleObjectBlob   =   "frmBI_EMISchedule.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBI_EMISchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.cmdCancel.Caption = "Close"
    Me.txtOutput.Text = ActiveCell.Address(External:=True)
End Sub

Private Sub cmdPickOutput_Click()
    On Error GoTo ErrHandler
    Dim picked As Range
    Me.Hide
    Set picked = Application.InputBox("Select the output top-left cell", "BeIndian EMI Schedule", Type:=8)
    If Not picked Is Nothing Then Me.txtOutput.Text = picked.Cells(1, 1).Address(External:=True)
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
    result = BI_EMISchedule(CDbl(Me.txtPrincipal.Text), CLng(Me.txtPeriods.Text), CDbl(Me.txtRate.Text))
    Set outputCell = Application.Range(Me.txtOutput.Text).Cells(1, 1)
    outputCell.Resize(UBound(result, 1), UBound(result, 2)).value = result
    outputCell.CurrentRegion.Columns.AutoFit
    Unload Me
    Exit Sub
ErrHandler:
    MsgBox "EMI schedule failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub
