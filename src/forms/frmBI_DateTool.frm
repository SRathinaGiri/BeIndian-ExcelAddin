VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBI_DateTool 
   Caption         =   "BeIndian Indian Workday Tool"
   ClientHeight    =   4240
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8380.001
   OleObjectBlob   =   "frmBI_DateTool.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBI_DateTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    If Len(Me.Tag) = 0 Then Me.Tag = "WorkDay"
    Me.txtTool.Text = Me.Tag
    Me.txtTool.Locked = True
    Me.txtOutput.Text = ActiveCell.Address(External:=True)
    ConfigureTool
End Sub

Private Sub ConfigureTool()
    Dim isNetwork As Boolean
    isNetwork = (UCase$(Me.txtTool.Text) = "NETWORKDAYS")
    Me.lblEnd.Visible = isNetwork: Me.txtEnd.Visible = isNetwork
    Me.lblDays.Visible = Not isNetwork: Me.txtDays.Visible = Not isNetwork
End Sub

Private Sub cmdPickHolidays_Click()
    PickRangeInto Me.txtHolidays, "Select optional holiday range"
End Sub

Private Sub cmdPickOutput_Click()
    PickRangeInto Me.txtOutput, "Select output cell"
End Sub

Private Sub PickRangeInto(ByVal targetTextBox As Object, ByVal prompt As String)
    On Error GoTo ErrHandler
    Dim picked As Range
    Me.Hide
    Set picked = Application.InputBox(prompt, "BeIndian Date Tool", Type:=8)
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
    Dim outputCell As Range, holidays As Variant, result As Variant
    Set outputCell = Application.Range(Me.txtOutput.Text).Cells(1, 1)
    If Len(Me.txtHolidays.Text) > 0 Then Set holidays = Application.Range(Me.txtHolidays.Text)
    If UCase$(Me.txtTool.Text) = "NETWORKDAYS" Then
        If IsObject(holidays) Then result = BI_NetworkDays_Indian(CDate(Me.txtStart.Text), CDate(Me.txtEnd.Text), holidays) Else result = BI_NetworkDays_Indian(CDate(Me.txtStart.Text), CDate(Me.txtEnd.Text))
    Else
        If IsObject(holidays) Then result = BI_WorkDay_Indian(CDate(Me.txtStart.Text), CLng(Me.txtDays.Text), holidays) Else result = BI_WorkDay_Indian(CDate(Me.txtStart.Text), CLng(Me.txtDays.Text))
    End If
    outputCell.value = result
    If IsDate(result) Then outputCell.NumberFormat = "dd-mmm-yyyy"
    Unload Me
    Exit Sub
ErrHandler:
    MsgBox "Date tool failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub
