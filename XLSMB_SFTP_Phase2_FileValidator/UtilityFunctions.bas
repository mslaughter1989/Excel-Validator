Attribute VB_Name = "UtilityFunctions"
Option Explicit

Public Function SelectValidationFile() As String
    Dim sFilePath As String
    
    With Application.fileDialog(msoFileDialogFilePicker)
        .Title = "Select File for Validation"
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        .Filters.Add "Text Files", "*.txt"
        .Filters.Add "All Files", "*.*"
        .FilterIndex = 1
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            sFilePath = .SelectedItems(1)
        End If
    End With
    
    SelectValidationFile = sFilePath
End Function


Public Sub ErrorHandler_Central(sProcedure As String, lErrorNum As Long, sErrorDesc As String, Optional sAdditionalInfo As String = "")
    Dim sLogEntry As String
    Dim sLogPath As String
    
    ' Create detailed log entry
    sLogEntry = Format(Now, "yyyy-mm-dd hh:mm:ss") & vbTab & _
                Environ("USERNAME") & vbTab & _
                sProcedure & vbTab & _
                lErrorNum & vbTab & _
                sErrorDesc
    
    If sAdditionalInfo <> "" Then
        sLogEntry = sLogEntry & vbTab & sAdditionalInfo
    End If
    
    ' Log to file
    sLogPath = ThisWorkbook.Path & "\ValidationErrors.log"
    WriteToLogFile sLogPath, sLogEntry
    
    ' Show user message
    MsgBox "An error occurred in " & sProcedure & ":" & vbCrLf & vbCrLf & _
           sErrorDesc & vbCrLf & vbCrLf & _
           "Error details have been logged.", vbCritical, "Validation System Error"
End Sub

Private Sub WriteToLogFile(sLogPath As String, sEntry As String)
    Dim lFileNum As Long
    
    On Error Resume Next
    lFileNum = FreeFile
    Open sLogPath For Append As #lFileNum
    Print #lFileNum, sEntry
    Close #lFileNum
    On Error GoTo 0
End Sub

Public Sub DisplayValidationResults(oResult As ValidationResult)
    Dim sMessage As String
    
    If oResult.IsValid Then
        sMessage = "Validation PASSED!" & vbCrLf & vbCrLf & _
                   "Total Records: " & oResult.TotalRecords & vbCrLf & _
                   "No errors found."
        MsgBox sMessage, vbInformation, "Validation Complete"
    Else
        sMessage = "Validation FAILED!" & vbCrLf & vbCrLf & _
                   "Total Records: " & oResult.TotalRecords & vbCrLf & _
                   "Errors Found: " & oResult.ErrorCount & vbCrLf & _
                   "Warnings: " & oResult.WarningCount & vbCrLf & vbCrLf & _
                   "Detailed report saved to: " & vbCrLf & oResult.ReportPath
        MsgBox sMessage, vbCritical, "Validation Complete"
    End If
    
    ' Option to open report
    If oResult.ReportPath <> "" Then
        If MsgBox("Would you like to open the validation report?", vbYesNo + vbQuestion, "Open Report") = vbYes Then
            shell "notepad.exe """ & oResult.ReportPath & """", vbNormalFocus
        End If
    End If
End Sub

Public Function GetFileNameFromPath(sFullPath As String) As String
    Dim vParts As Variant
    vParts = Split(sFullPath, "\")
    GetFileNameFromPath = CStr(vParts(UBound(vParts)))
End Function
