Attribute VB_Name = "GitHubTool_Export"

Sub ExportAllToFolderForGitHub()

    Dim vbComp As Object
    Dim folderPath As String
    Dim file As Object
    Dim folderDialog As FileDialog
    Dim fso As Object

    'folder selection
    Set folderDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With folderDialog
        .Title = "Select Folder to Export Files to"
        If .Show = -1 Then
            folderPath = .SelectedItems(1) + "\" '
        Else
            MsgBox "No folder selected.", vbExclamation
            Exit Sub
        End If
    End With
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        MsgBox "The selected folder does not exist.", vbExclamation
        Exit Sub
    End If
    
    'export files
    For Each vbComp In ThisWorkbook.vbProject.VBComponents
    
        If Not vbComp.Name = "Sheet1" Then 'write to skip ThisWorkbook

            'likely only exports for .bas, .frm, and .cls modules
            Select Case vbComp.Type
                Case 1 Or vbext_ct_StdModule
                    fileExtension = ".bas"
                Case vbext_ct_ClassModule
                    fileExtension = ".cls"
                Case 3 Or vbext_ct_MSForm
                    fileExtension = ".frm"
                Case 100 Or vbext_ct_Document
                    fileExtension = ".cls"
                Case Else
                    fileExtension = ""
            End Select
            
            ' export if valid extension
            If fileExtension <> "" Then
                vbComp.Export folderPath & vbComp.Name & fileExtension
            End If
            
        End If
    Next vbComp
    
    MsgBox "VBA files exported!"
    
End Sub


