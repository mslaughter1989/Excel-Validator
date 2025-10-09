Attribute VB_Name = "Module1"
    Sub ExportAllModules()
        Dim strPath As String
        Dim vbc As Object

        ' Prompt user to select a folder for saving the modules
        With Application.fileDialog(msoFileDialogFolderPicker)
            .Title = "Select a folder to save VBA modules"
            If .Show Then
                strPath = .SelectedItems(1)
                ' Ensure the path ends with a backslash
                If Right(strPath, 1) <> "\" Then
                    strPath = strPath & "\"
                End If
            Else
                MsgBox "No folder selected. Operation cancelled.", vbExclamation
                Exit Sub
            End If
        End With

        ' Loop through each component (module, class, userform) in the active project
        For Each vbc In ActiveWorkbook.VBProject.VBComponents ' Or ActiveDocument.VBProject for Word
            ' Export the component
            vbc.Export strPath & vbc.Name & ".bas" ' Exports as .bas for standard modules
        Next vbc

        MsgBox "All modules exported successfully to: " & strPath, vbInformation
    End Sub
