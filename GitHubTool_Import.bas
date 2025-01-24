Attribute VB_Name = "GitHubTool_Import"

Sub ImportVBAFilesFromGitHub()
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim fileName As String
    Dim moduleName As String
    Dim vbProject As Object
    Dim vbComp As Object
    Dim filePath As String
    Dim extName As String
    
    'folder selection
    Set folderDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With folderDialog
        .Title = "Select Folder to Export Files to"
        If .Show = -1 Then
            filePath = .SelectedItems(1) ' Get selected folder
        Else
            MsgBox "No folder selected.", vbExclamation
            Exit Sub
        End If
    End With
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(filePath) Then
        MsgBox "The selected folder does not exist.", vbExclamation
        Exit Sub
    End If
    
    Set folder = fso.GetFolder(filePath)
    Set vbProject = ThisWorkbook.vbProject

    For Each file In folder.Files
        
        If Not file.Name = "GitHubTool_Import.bas" Then
        
            extName = Split(file.Name, ".")(1)
            moduleName = Split(file.Name, ".")(0)
            
            'import valid file
            If extName = "bas" Or extName = "txt" Or extName = "frm" Then
                
                On Error Resume Next
                Set vbComp = vbProject.VBComponents(moduleName)
                On Error GoTo 0
                
                'if the module exists, remove it
                If Not vbComp Is Nothing Then
                    vbProject.VBComponents.Remove vbComp
                    Set vbComp = Nothing
                End If
                
                vbProject.VBComponents.Import file.path
            End If
            
            'hardcode to import ThisWorkbook
            If extName = "cls" And moduleName = "ThisWorkbook" Then
                
                vbProject.VBComponents("ThisWorkbook").CodeModule.DeleteLines 1, vbProject.VBComponents("ThisWorkbook").CodeModule.CountOfLines
                vbProject.VBComponents("ThisWorkbook").CodeModule.AddFromFile file
                vbProject.VBComponents("ThisWorkbook").CodeModule.DeleteLines 1, 4
                
            End If
        
        End If
        
    Next file
    
    
    MsgBox "VBA files imported!"
    
End Sub



