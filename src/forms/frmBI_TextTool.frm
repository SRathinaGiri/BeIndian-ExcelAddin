VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBI_TextTool 
   Caption         =   "BeIndian Text Tool"
   ClientHeight    =   4140
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8380.001
   OleObjectBlob   =   "frmBI_TextTool.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBI_TextTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.cmdCancel.Caption = "Close"
    With Me.cmbAction
        .AddItem "BeginsWith"
        .AddItem "Contains"
        .AddItem "EndsWith"
        .AddItem "ExtractNumbers"
        .AddItem "ExtractText"
        .AddItem "FindOneOf"
        .AddItem "Fuzzy"
        .AddItem "ReverseText"
        .AddItem "WordCount"
        .AddItem "Insert"
        .AddItem "Num2Word"
        .AddItem "IndianNum2Word"
        .ListIndex = 0
    End With
    If TypeName(Selection) = "Range" Then Me.txtText.Text = Selection.Cells(1, 1).Text
    Me.txtOutput.Text = ActiveCell.Address(External:=True)
End Sub

Private Sub cmdPickOutput_Click()
    On Error GoTo ErrHandler
    Dim picked As Range
    Me.Hide
    Set picked = Application.InputBox("Select output cell", "BeIndian Text Tool", Type:=8)
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
    Select Case Me.cmbAction.value
        Case "BeginsWith": result = BI_BeginsWith(Me.txtFind.Text, Me.txtText.Text, Me.chkCase.value)
        Case "Contains": result = BI_Contains(Me.txtFind.Text, Me.txtText.Text, Me.chkCase.value)
        Case "EndsWith": result = BI_EndsWith(Me.txtFind.Text, Me.txtText.Text, Me.chkCase.value)
        Case "ExtractNumbers": result = BI_ExtractNumbers(Me.txtText.Text)
        Case "ExtractText": result = BI_ExtractText(Me.txtText.Text)
        Case "FindOneOf": result = BI_FindOneOf(Me.txtText.Text, Me.txtFind.Text)
        Case "Fuzzy": result = BI_Fuzzy(Me.txtText.Text, Me.txtFind.Text)
        Case "ReverseText": result = BI_ReverseText(Me.txtText.Text)
        Case "WordCount": result = BI_WordCount(Me.txtText.Text)
        Case "Insert": result = BI_Insert(Me.txtText.Text, CLng(Me.txtPosition.Text), Me.txtFind.Text)
        Case "Num2Word": result = BI_Num2Word(Me.txtText.Text)
        Case "IndianNum2Word": result = BI_IndianNum2Word(Me.txtText.Text)
    End Select
    Application.Range(Me.txtOutput.Text).Cells(1, 1).value = result
    Unload Me
    Exit Sub
ErrHandler:
    MsgBox "Text tool failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub
