VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBI_Score 
   Caption         =   "BeIndian Score Calculator"
   ClientHeight    =   4340
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8380.001
   OleObjectBlob   =   "frmBI_Score.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBI_Score"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.cmdCancel.Caption = "Close"
    If Len(Me.Tag) = 0 Then Me.Tag = "Altman"
    Me.txtModel.Text = Me.Tag
    Me.txtModel.Locked = True
    Me.txtOutput.Text = ActiveCell.Address(External:=True)
    If TypeName(Selection) = "Range" Then Me.txtData.Text = Selection.Address(External:=True)
    ConfigureModel
End Sub

Private Sub ConfigureModel()
    Dim twoRange As Boolean
    twoRange = (UCase$(Me.txtModel.Text) = "BENEISH" Or UCase$(Me.txtModel.Text) = "PIOTROSKI")
    Me.lblData.Caption = IIf(twoRange, "Last year", "Data range")
    Me.lblCurrent.Visible = twoRange
    Me.txtCurrent.Visible = twoRange
    Me.cmdPickCurrent.Visible = twoRange
End Sub

Private Sub cmdPickData_Click()
    PickRangeInto Me.txtData, "Select the input range"
End Sub

Private Sub cmdPickCurrent_Click()
    PickRangeInto Me.txtCurrent, "Select the current year range"
End Sub

Private Sub cmdPickOutput_Click()
    PickRangeInto Me.txtOutput, "Select the output top-left cell"
End Sub

Private Sub PickRangeInto(ByVal targetTextBox As Object, ByVal prompt As String)
    On Error GoTo ErrHandler
    Dim picked As Range
    Me.Hide
    Set picked = Application.InputBox(prompt, "BeIndian Score Calculator", Type:=8)
    If Not picked Is Nothing Then targetTextBox.Text = picked.Address(External:=True)
CleanExit:
    Me.Show
    Exit Sub
ErrHandler:
    Resume CleanExit
End Sub

Private Sub cmdTemplate_Click()
    On Error GoTo ErrHandler
    BI_OutputScoreInputTemplate Me.txtModel.Text, Application.Range(Me.txtOutput.Text).Cells(1, 1)
    Exit Sub
ErrHandler:
    MsgBox "Template creation failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRun_Click()
    On Error GoTo ErrHandler
    Dim result As Variant
    Dim outputCell As Range
    Set outputCell = Application.Range(Me.txtOutput.Text).Cells(1, 1)
    Select Case UCase$(Me.txtModel.Text)
        Case "ALTMAN"
            result = BI_AltmanZScore(Application.Range(Me.txtData.Text))
        Case "BENEISH"
            result = BI_BeneishMScore(Application.Range(Me.txtData.Text), Application.Range(Me.txtCurrent.Text))
        Case "OHLSON"
            result = BI_OhlsonsOScore(Application.Range(Me.txtData.Text))
        Case "PIOTROSKI"
            result = BI_PiotroskiFScore(Application.Range(Me.txtData.Text), Application.Range(Me.txtCurrent.Text))
        Case Else
            Err.Raise vbObjectError + 530, "frmBI_Score", "Unknown model."
    End Select
    outputCell.Resize(UBound(result, 1), UBound(result, 2)).value = result
    outputCell.CurrentRegion.Columns.AutoFit
    Unload Me
    Exit Sub
ErrHandler:
    MsgBox "Score calculation failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub
