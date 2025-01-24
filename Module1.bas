Attribute VB_Name = "Module1"
Public thermQuikPathFile As String

Sub defineThermQuikPath()
    thermQuikPathFile = Application.StartupPath & "\20250102_ThermQuik_V1.xlam'!"
End Sub
'Callback for Grp1_Btn1 onAction
Sub TQ_Run(control As IRibbonControl)
    Application.Run ("'" + thermQuikPathFile & "eta.eta")
End Sub

'Callback for Grp2_Btn1 onAction
Sub conBoldSub(control As IRibbonControl)
    Application.Run ("'" + thermQuikPathFile & "eta_import.eta_import")
End Sub

'Callback for Grp3_Btn1 onAction
Sub TQ_Plot(control As IRibbonControl)
    Application.Run ("'" + thermQuikPathFile & "tq_plot.tq_plot")
End Sub

'Callback for Grp3_Btn2 onAction
Sub TQ_Export(control As IRibbonControl)
    Application.Run ("'" + thermQuikPathFile & "tq_export.tq_export")
End Sub

'Callback for Grp4_Btn1 onAction
Sub TQ_Help(control As IRibbonControl)
    Application.Run ("'" + thermQuikPathFile & "tq_help.tq_help")
End Sub


