Attribute VB_Name = "modBI_Ribbon"
Option Explicit

Private gRibbon As IRibbonUI

Public Sub BI_RibbonOnLoad(ribbon As IRibbonUI)
    On Error GoTo ErrHandler
    Set gRibbon = ribbon
    Exit Sub
ErrHandler:
    MsgBox "BeIndian Ribbon failed to load: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_ShowAbout(control As IRibbonControl)
    On Error GoTo ErrHandler
    frmBI_About.Show
    Exit Sub
ErrHandler:
    MsgBox "Could not show About: " & Err.Description, vbExclamation, "BeIndian"
End Sub

Public Sub BI_Tool_NotYetPorted(control As IRibbonControl)
    On Error GoTo ErrHandler
    MsgBox "This tool is in the porting queue. Formula/UDF conversion is being done in source batches.", vbInformation, "BeIndian"
    Exit Sub
ErrHandler:
    MsgBox "Tool failed: " & Err.Description, vbExclamation, "BeIndian"
End Sub
