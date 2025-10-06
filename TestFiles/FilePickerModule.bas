Attribute VB_Name = "FilePickerModule"

Sub RunValidationOnSelectedCSVs()
    Dim fDialog As FileDialog
    Dim selectedFiles As Variant
    Dim i As Long
    Dim wb As Workbook
    Dim baseName As String
    Dim filePath As String

    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .Title = "Select one or more CSV files"
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        .AllowMultiSelect = True
        If .Show <> -1 Then
            MsgBox "No files selected. Macro canceled.", vbExclamation
            Exit Sub
        End If
        Set selectedFiles = .SelectedItems
    End With

    For i = 1 To selectedFiles.Count
        filePath = selectedFiles(i)
        On Error Resume Next
        Set wb = Workbooks.Open(Filename:=filePath)
        If wb Is Nothing Then
            MsgBox "Could not open file: " & filePath, vbExclamation
            On Error GoTo 0
            GoTo NextFile
        End If
        On Error GoTo 0

        baseName = Mid(filePath, InStrRev(filePath, "\") + 1)
        baseName = Left(baseName, InStrRev(baseName, ".") - 1)

        Call Validate_WithRequiredColumnsSplitLog(baseName)

        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0

NextFile:
    Next i

    MsgBox "Validation complete for selected CSV files.", vbInformation
End Sub
