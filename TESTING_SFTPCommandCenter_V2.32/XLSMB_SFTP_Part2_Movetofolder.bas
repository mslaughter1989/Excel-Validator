Attribute VB_Name = "XLSMB_SFTP_Part2_Movetofolder"
Sub MoveCSVFilesToFolders()
    Dim ws As Worksheet
    Dim csvPath As Variant
    Dim refData As Variant
    Dim fileDialog As fileDialog
    Dim selectedFiles As FileDialogSelectedItems
    Dim originalFileName As String
    Dim originalDateFormat As String
    Dim updatedDateFormat As String
    Dim sftpName As String
    Dim GroupID As String
    Dim finalSaveFormat As String
    Dim saveFolder As String
    Dim i As Long
    Dim fileDate As String
    Dim fileDateFormatted As String
    Dim fso As Object
    Dim matchFound As Boolean
    Dim processedCount As Long
    Dim ErrorCount As Long
    
    ' Summary tracking variables
    Dim movedList As String
    Dim createdFoldersList As String
    
    ' UPDATED: Read from internal worksheet instead of external CSV file
    refData = GetWorksheetData("Parsed_SFTPFiles")
    
    ' Check if worksheet data was loaded properly
    If IsEmpty(refData) Then
        MsgBox "Failed to load data from 'Parsed_SFTPFiles' worksheet. Please ensure the worksheet exists and contains data.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Open file dialog to select multiple CSV files
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    With fileDialog
        .Title = "Select CSV Files to Move to Folders"
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        .Filters.Add "Excel Files", "*.xlsx"
        .AllowMultiSelect = True
        If .Show = -1 Then
            Set selectedFiles = .SelectedItems
        Else
            MsgBox "No files selected.", vbInformation, "Operation Cancelled"
            Exit Sub
        End If
    End With

    ' File System Object for moving files
    Set fso = CreateObject("Scripting.FileSystemObject")
    processedCount = 0
    ErrorCount = 0
    
    ' Initialize summary strings
    movedList = "Files Successfully Moved:" & vbCrLf
    createdFoldersList = "New Folders Created:" & vbCrLf

    ' Loop through selected files
    For Each csvPath In selectedFiles
        originalFileName = fso.GetFileName(csvPath)
        matchFound = False
        
        ' Find matching entry in worksheet data (skip header row)
        For i = 2 To UBound(refData, 1)
            ' Check if we have enough columns (14 COLUMNS)
            If UBound(refData, 2) >= 14 Then
                ' Use Column 13 (M) for Final Save Format
                Dim formatToCheck As String
                formatToCheck = CStr(refData(i, 13)) ' Column M (13th column)
                
                If IsFileMatchFinalFormat(originalFileName, formatToCheck, CStr(refData(i, 1))) Then
                    ' Extract necessary details from matching row
                    sftpName = CStr(refData(i, 1))
                    originalDateFormat = CStr(refData(i, 6)) ' File Date Format from filename (Column F)
                    updatedDateFormat = CStr(refData(i, 12)) ' Updated file date format (Column L)
                    finalSaveFormat = formatToCheck
                    saveFolder = CStr(refData(i, 14)) ' Save Folder (Column N)
                    
                    ' UPDATED: Replace path placeholders for cross-user compatibility
                    saveFolder = ResolveFolderPath(saveFolder)
                    
                    ' UPDATED: Replace placeholders in saveFolder path
                    saveFolder = Replace(saveFolder, "[Adjusted GroupName]", CStr(refData(i, 10))) ' Column J
                    saveFolder = Replace(saveFolder, "[Adjusted groupID]", CStr(refData(i, 11)))   ' Column K
                    saveFolder = Replace(saveFolder, "[GroupName]", CStr(refData(i, 10)))          ' Alternative
                    saveFolder = Replace(saveFolder, "[groupID]", CStr(refData(i, 11)))            ' Alternative
                    
                    ' Skip if no save folder specified or placeholders not resolved
                    If saveFolder = "" Or saveFolder = "**NOT INCLUDED**" Or InStr(saveFolder, "[") > 0 Then
                        MsgBox "No save folder specified or placeholders not resolved for: " & originalFileName & vbCrLf & _
                               "Save folder: " & saveFolder, vbExclamation, "Warning"
                        ErrorCount = ErrorCount + 1
                        GoTo NextFile
                    End If
                    
                    ' Extract date from filename
                    fileDate = ExtractDateFromFinalFormat(originalFileName, finalSaveFormat, originalDateFormat, sftpName)
                    If fileDate = "" Then
                        MsgBox "Date not found or invalid in file: " & originalFileName, vbExclamation, "Error"
                        ErrorCount = ErrorCount + 1
                        GoTo NextFile
                    End If
                    
                    ' Format date for folder name
                    Dim folderDateFormat As String
                    folderDateFormat = FormatDateForFolder(fileDate, updatedDateFormat)
                    If folderDateFormat = "" Then
                        MsgBox "Unable to format date for folder: " & originalFileName, vbExclamation, "Error"
                        ErrorCount = ErrorCount + 1
                        GoTo NextFile
                    End If
                    
                    ' Build target folder path
                    Dim targetFolder As String
                    Dim targetPath As String
                    targetFolder = fso.BuildPath(saveFolder, folderDateFormat)
                    
                    ' Create folder if it doesn't exist
                    On Error GoTo FolderError
                    If Not fso.FolderExists(targetFolder) Then
                        fso.CreateFolder targetFolder
                        createdFoldersList = createdFoldersList & targetFolder & vbCrLf
                    End If
                    On Error GoTo 0
                    
                    ' Build final target path for the file
                    targetPath = fso.BuildPath(targetFolder, originalFileName)
                    
                    ' Check if target file exists
                    If fso.FileExists(targetPath) Then
                        If MsgBox("File " & originalFileName & " already exists in target folder. Overwrite?", vbYesNo + vbQuestion, "File Exists") = vbNo Then
                            GoTo NextFile
                        Else
                            fso.DeleteFile targetPath
                        End If
                    End If
                    
                    ' Move file to target location
                    On Error GoTo FileError
                    fso.MoveFile csvPath, targetPath
                    movedList = movedList & originalFileName & " ? " & targetFolder & vbCrLf
                    processedCount = processedCount + 1
                    matchFound = True
                    On Error GoTo 0
                    Exit For
                    
FolderError:
                    MsgBox "Error creating folder: " & targetFolder & vbCrLf & "Error: " & Err.Description, vbExclamation, "Folder Error"
                    ErrorCount = ErrorCount + 1
                    On Error GoTo 0
                    GoTo NextFile
                End If
            Else
                MsgBox "Worksheet doesn't have enough columns. Expected at least 14 columns.", vbCritical
                Exit Sub
            End If
        Next i
        
        If Not matchFound Then
            MsgBox "No matching final format found for file: " & originalFileName & vbCrLf & _
                   "Please check the Parsed_SFTPFiles worksheet for the correct final save format.", vbExclamation, "Warning"
            ErrorCount = ErrorCount + 1
        End If
        
NextFile:
    Next csvPath

    ' Show comprehensive summary
    Dim summaryMsg As String
    summaryMsg = "FILE MOVING COMPLETE!" & vbCrLf & vbCrLf & _
                 createdFoldersList & vbCrLf & _
                 movedList & vbCrLf & _
                 "Files processed: " & processedCount & vbCrLf & _
                 "Errors: " & ErrorCount
    
    MsgBox summaryMsg, vbInformation, "Moving Summary"
    Exit Sub

FileError:
    MsgBox "Error moving file: " & originalFileName & vbCrLf & _
           "Target: " & targetPath & vbCrLf & _
           "Error: " & Err.Description, vbCritical, "File Error"
    ErrorCount = ErrorCount + 1
    On Error GoTo 0
    Resume NextFile
End Sub

' UPDATED: Function to resolve folder path placeholders
Function ResolveFolderPath(folderPath As String) As String
    Dim resolvedPath As String
    resolvedPath = folderPath
    
    ' Replace OneDrive placeholders with actual paths
    resolvedPath = Replace(resolvedPath, "{OneDriveCommercial}", GetOneDriveCommercialPath())
    resolvedPath = Replace(resolvedPath, "{OneDrive}", GetOneDrivePath())
    resolvedPath = Replace(resolvedPath, "{UserProfile}", Environ("USERPROFILE"))
    
    ResolveFolderPath = resolvedPath
End Function

' Function to get OneDrive for Business path
Function GetOneDriveCommercialPath() As String
    Dim oneDrivePath As String
    
    ' Try OneDrive for Business environment variable first
    oneDrivePath = Environ("OneDriveCommercial")
    
    ' If not found, try alternative methods
    If oneDrivePath = "" Then
        oneDrivePath = Environ("OneDrive")
    End If
    
    ' If still not found, try default path
    If oneDrivePath = "" Then
        oneDrivePath = Environ("USERPROFILE") & "\OneDrive - Recuro Health"
    End If
    
    GetOneDriveCommercialPath = oneDrivePath
End Function

' Function to get personal OneDrive path
Function GetOneDrivePath() As String
    Dim oneDrivePath As String
    oneDrivePath = Environ("OneDrive")
    
    If oneDrivePath = "" Then
        oneDrivePath = Environ("USERPROFILE") & "\OneDrive"
    End If
    
    GetOneDrivePath = oneDrivePath
End Function

' UPDATED: Shared function for reading worksheet data
Function GetWorksheetData(sheetName As String) As Variant
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets(sheetName)
    
    With ws
        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row
        lastCol = .Cells(1, .Columns.count).End(xlToLeft).Column
        
        ' Ensure we have data (expecting 14 columns)
        If lastRow > 1 And lastCol >= 14 Then
            GetWorksheetData = .Range(.Cells(1, 1), .Cells(lastRow, lastCol)).Value
        Else
            MsgBox "Worksheet '" & sheetName & "' must have at least 2 rows and 14 columns. Found " & lastRow & " rows and " & lastCol & " columns.", vbCritical
            GetWorksheetData = Empty
        End If
    End With
    Exit Function
    
ErrorHandler:
    MsgBox "Error reading worksheet '" & sheetName & "': " & Err.Description & vbCrLf & _
           "Please ensure the worksheet exists and contains the required data.", vbCritical
    GetWorksheetData = Empty
End Function

Function IsFileMatchFinalFormat(FileName As String, finalSaveFormat As String, sftpName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim regex As Object
    Dim regexPattern As String
    
    regexPattern = finalSaveFormat
    
    ' Replace ALL possible placeholders with regex patterns
    regexPattern = Replace(regexPattern, "[Adjusted GroupName]", ".*")
    regexPattern = Replace(regexPattern, "[Adjusted groupID]", "\d+")
    regexPattern = Replace(regexPattern, "Adjusted GroupName", sftpName)
    regexPattern = Replace(regexPattern, "groupid", "\d+")
    regexPattern = Replace(regexPattern, "mmddyyyy", "\d{8}")
    regexPattern = Replace(regexPattern, "mmddyy", "\d{6}")
    regexPattern = Replace(regexPattern, "yyyymmdd", "\d{8}")
    
    ' Handle file extensions flexibly
    If InStr(regexPattern, ".csv") > 0 Then
        regexPattern = Replace(regexPattern, ".csv", "\.(csv|xlsx)")
    ElseIf InStr(regexPattern, ".xlsx") > 0 Then
        regexPattern = Replace(regexPattern, ".xlsx", "\.(csv|xlsx)")
    Else
        regexPattern = regexPattern & "\.(csv|xlsx)"
    End If
    
    ' Make it an exact match pattern
    regexPattern = "^" & regexPattern & "$"
    
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Pattern = regexPattern
        .IgnoreCase = True
        .Global = False
    End With
    
    IsFileMatchFinalFormat = regex.Test(FileName)
    Exit Function
    
ErrorHandler:
    IsFileMatchFinalFormat = False
End Function

Function ExtractDateFromFinalFormat(FileName As String, finalSaveFormat As String, originalDateFormat As String, sftpName As String) As String
    On Error GoTo ErrorHandler
    
    Dim regex As Object
    Dim matches As Object
    
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Pattern = "(\d{8})"
        .IgnoreCase = True
        .Global = False
    End With
    
    Set matches = regex.Execute(FileName)
    If matches.count > 0 Then
        ExtractDateFromFinalFormat = matches(0).SubMatches(0)
    Else
        ExtractDateFromFinalFormat = ""
    End If
    Exit Function
    
ErrorHandler:
    ExtractDateFromFinalFormat = ""
End Function

Function FormatDateForFolder(dateString As String, updatedDateFormat As String) As String
    On Error GoTo ErrorHandler
    
    Dim yearPart As String, monthPart As String, dayPart As String
    Dim monthNames As Variant
    Dim monthIndex As Integer
    
    monthNames = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", _
                      "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    
    If Len(dateString) <> 8 Then
        FormatDateForFolder = ""
        Exit Function
    End If
    
    Dim actualFormat As String
    actualFormat = DetermineDateFormat(dateString)
    
    Select Case actualFormat
        Case "mmddyy"
            monthPart = Left(dateString, 2)
            dayPart = Mid(dateString, 3, 2)
            yearPart = Right(dateString, 2)
        Case "mmddyyyy"
            monthPart = Left(dateString, 2)
            dayPart = Mid(dateString, 3, 2)
            yearPart = Right(dateString, 2)
        Case "yyyymmdd"
            yearPart = Right(Left(dateString, 4), 2)
            monthPart = Mid(dateString, 5, 2)
            dayPart = Right(dateString, 2)
        Case Else
            FormatDateForFolder = ""
            Exit Function
    End Select
    
    monthIndex = CInt(monthPart) - 1
    If monthIndex < 0 Or monthIndex > 11 Then
        FormatDateForFolder = ""
        Exit Function
    End If
    
    FormatDateForFolder = Format(monthPart, "00") & monthNames(monthIndex) & yearPart
    Exit Function
    
ErrorHandler:
    FormatDateForFolder = ""
End Function

Function DetermineDateFormat(dateString As String) As String
    If Len(dateString) <> 8 Then
        DetermineDateFormat = "mmddyyyy"
        Exit Function
    End If
    
    Dim firstTwo As Integer
    firstTwo = CInt(Left(dateString, 2))
    
    If firstTwo > 12 Then
        DetermineDateFormat = "yyyymmdd"
    Else
        DetermineDateFormat = "mmddyyyy"
    End If
End Function

