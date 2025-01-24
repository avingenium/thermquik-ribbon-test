Attribute VB_Name = "Module1"

Public etaPathFile As String

Public Sub defineEtaPath()
    etaPathFile = Application.StartupPath & "\20250102_ThermQuik_V1.xlam'!"
End Sub

'Callback for Grp1_Btn1 onAction
Sub TQ_Run(control As IRibbonControl)
    If etaPathFile = "" Then Call defineEtaPath
    Application.Run ("'" + etaPathFile & "eta.eta")
End Sub

'Callback for Grp2_Btn1 onAction
Sub conBoldSub(control As IRibbonControl)
    If etaPathFile = "" Then Call defineEtaPath
    Application.Run ("'" + etaPathFile & "eta_import.eta_import")
End Sub

'Callback for Grp3_Btn1 onAction
Sub TQ_Plot(control As IRibbonControl)
    If etaPathFile = "" Then Call defineEtaPath
    Application.Run ("'" + etaPathFile & "tq_plot.tq_plot")
End Sub

'Callback for Grp3_Btn2 onAction
Sub TQ_Export(control As IRibbonControl)
    If etaPathFile = "" Then Call defineEtaPath
    Application.Run ("'" + etaPathFile & "tq_export.tq_export")
End Sub

'Callback for Grp4_Btn1 onAction
Sub TQ_Help(control As IRibbonControl)
    If etaPathFile = "" Then Call defineEtaPath
    Application.Run ("'" + etaPathFile & "tq_help.tq_help")
End Sub



