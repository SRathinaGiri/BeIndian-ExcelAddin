VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBI_About 
   Caption         =   "About BeIndian"
   ClientHeight    =   4920
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10200
   OleObjectBlob   =   "frmBI_About.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBI_About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Option Explicit

Private Sub UserForm_Initialize()
    Me.lblVersion.Caption = "Version: " & BI_VERSION
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
