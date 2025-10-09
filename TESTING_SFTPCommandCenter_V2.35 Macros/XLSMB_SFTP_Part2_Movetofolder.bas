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
    Dim groupID As String
    Dim finalSaveFormat As String
    Dim saveFolder As String
    Dim i As Long
    Dim fileDate As String
    Dim fileDateFormatted As String
    Dim fso As Object
    Dim matchFound As Boolean
    Dim processedCount As Long
    Dim ErrorCount As Long
    
    Dim movedList As String
    Dim createdFoldersList As String
    
    ' *** NEW: Get base folder path from config ***
    Dim baseFolderPath As String
    baseFolderPath = GetBaseFolderPath()
    
    If baseFolderPath = "" Then
        MsgBox "No base folder path configured. Please run Part1 first to set up the base folder.", vbExclamation, "Configuration Missing"
        Exit Sub
    End If
    
    refData = GetWorksheetData("Parsed_SFTPFiles")
    
    If IsEmpty(refData) Then
        MsgBox "Failed to load data from 'Parsed_SFTPFiles' worksheet. Please ensure the worksheet exists and contains data.", vbCritical, "Error"
        Exit Sub
    End If
    
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

    Set fso = CreateObject("Scripting.FileSystemObject")
    processedCount = 0
    ErrorCount = 0
    
    movedList = "Files Successfully Moved:" & vbCrLf
    createdFoldersList = "New Folders Created:" & vbCrLf

    For Each csvPath In selectedFiles
        originalFileName = fso.GetFileName(csvPath)
        matchFound = False
        
        For i = 2 To UBound(refData, 1)
            If UBound(refData, 2) >= 14 Then
                Dim formatToCheck As String
                formatToCheck = CStr(refData(i, 13))
                
                If IsFileMatchFinalFormat(originalFileName, formatToCheck, CStr(refData(i, 1))) Then
                    sftpName = CStr(refData(i, 1))
                    originalDateFormat = CStr(refData(i, 6))
                    updatedDateFormat = CStr(refData(i, 12))
                    finalSaveFormat = formatToCheck
                    saveFolder = CStr(refData(i, 14))
                    
                    ' *** MODIFIED: Use base folder path instead of OneDrive resolution ***
                    saveFolder = ResolveFolderPath(saveFolder, baseFolderPath)
                    
                    saveFolder = Replace(saveFolder, "[Adjusted GroupName]", CStr(refData(i, 10)))
                    saveFolder = Replace(saveFolder, "[Adjusted groupID]", CStr(refData(i, 11)))
                    saveFolder = Replace(saveFolder, "[GroupName]", CStr(refData(i, 10)))
                    saveFolder = Replace(saveFolder, "[groupID]", CStr(refData(i, 11)))
                    
                    If saveFolder = "" Or saveFolder = "**NOT INCLUDED**" Or InStr(saveFolder, "[") > 0 Then
                        MsgBox "No save folder specified or placeholders not resolved for: " & originalFileName & vbCrLf & _
                               "Save folder: " & saveFolder, vbExclamation, "Warning"
                        ErrorCount = ErrorCount + 1
                        GoTo NextFile
                    End If
                    
                    fileDate = ExtractDateFromFinalFormat(originalFileName, finalSaveFormat, originalDateFormat, sftpName)
                    If fileDate = "" Then
                        MsgBox "Date not found or invalid in file: " & originalFileName, vbExclamation, "Error"
                        ErrorCount = ErrorCount + 1
                        GoTo NextFile
                    End If
                    
                    Dim folderDateFormat As String
                    folderDateFormat = FormatDateForFolder(fileDate, updatedDateFormat)
                    If folderDateFormat = "" Then
                        MsgBox "Unable to format date for folder: " & originalFileName, vbExclamation, "Error"
                        ErrorCount = ErrorCount + 1
                        GoTo NextFile
                    End If
                    
                    Dim targetFolder As String
                    Dim targetPath As String
                    targetFolder = fso.BuildPath(saveFolder, folderDateFormat)
                    
                    On Error GoTo FolderError
                    If Not fso.FolderExists(targetFolder) Then
                        fso.CreateFolder targetFolder
                        createdFoldersList = createdFoldersList & targetFolder & vbCrLf
                    End If
                    On Error GoTo 0
                    
                    targetPath = fso.BuildPath(targetFolder, originalFileName)
                    
                    If fso.FileExists(targetPath) Then
                        If MsgBox("File " & originalFileName & " already exists in target folder. Overwrite?", vbYesNo + vbQuestion, "File Exists") = vbNo Then
                            GoTo NextFile
                        Else
                            fso.DeleteFile targetPath
                        End If
                    End If
                    
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

Function GetBaseFolderPath() As String
    Dim basePath As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Try OneDrive environment variable
    basePath = Environ("OneDrive")
    If basePath <> "" Then
        basePath = basePath & "\Documents - Ops"
        If fso.FolderExists(basePath) Then
            GetBaseFolderPath = basePath
            Exit Function
        End If
    End If
    
    ' Try OneDrive - Recuro Health
    basePath = Environ("USERPROFILE") & "\OneDrive - Recuro Health\Documents - Ops"
    If fso.FolderExists(basePath) Then
        GetBaseFolderPath = basePath
        Exit Function
    End If
    
    ' Try OneDriveCommercial
    basePath = Environ("OneDriveCommercial")
    If basePath <> "" Then
        basePath = basePath & "\Documents - Ops"
        If fso.FolderExists(basePath) Then
            GetBaseFolderPath = basePath
            Exit Function
        End If
    End If
    
    MsgBox "Could not find OneDrive folder at:" & vbCrLf & _
           Environ("USERPROFILE") & "\OneDrive - Recuro Health\Documents - Ops", _
           vbCritical, "OneDrive Path Not Found"
    GetBaseFolderPath = ""
End Function

' ===== MODIFIED: Resolve folder path using base folder =====
Function ResolveFolderPath(folderPath As String, baseFolderPath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim cleanPath As String
    cleanPath = folderPath
    
    cleanPath = Replace(cleanPath, "{OneDriveCommercial}\", "")
    cleanPath = Replace(cleanPath, "{OneDrive}\", "")
    cleanPath = Replace(cleanPath, "{UserProfile}\", "")
    
    If Left(cleanPath, Len("Documents - Ops")) = "Documents - Ops" Then
        If Right(baseFolderPath, Len("Documents - Ops")) = "Documents - Ops" Then
            cleanPath = Mid(cleanPath, Len("Documents - Ops") + 2)
            ResolveFolderPath = fso.BuildPath(baseFolderPath, cleanPath)
        Else
            ResolveFolderPath = fso.BuildPath(baseFolderPath, cleanPath)
        End If
    Else
        ResolveFolderPath = fso.BuildPath(baseFolderPath, cleanPath)
    End If
End Function

' ===== ALL EXISTING FUNCTIONS BELOW (UNCHANGED) =====

Function GetWorksheetData(sheetName As String) As Variant
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets(sheetName)
    
    With ws
        lastRow = .Cells(.Rows.count, 1).End(xlUp).row
        lastCol = .Cells(1, .Columns.count).End(xlToLeft).Column
        
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

Function IsFileMatchFinalFormat(fileName As String, finalSaveFormat As String, sftpName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim regex As Object
    Dim regexPattern As String
    
    regexPattern = finalSaveFormat
    
    regexPattern = Replace(regexPattern, "[Adjusted GroupName]", ".*")
    regexPattern = Replace(regexPattern, "[Adjusted groupID]", "\d+")
    regexPattern = Replace(regexPattern, "Adjusted GroupName", sftpName)
    regexPattern = Replace(regexPattern, "groupid", "\d+")
    regexPattern = Replace(regexPattern, "mmddyyyy", "\d{8}")
    regexPattern = Replace(regexPattern, "mmddyy", "\d{6}")
    regexPattern = Replace(regexPattern, "yyyymmdd", "\d{8}")
    
    If InStr(regexPattern, ".csv") > 0 Then
        regexPattern = Replace(regexPattern, ".csv", "\.(csv|xlsx)")
    ElseIf InStr(regexPattern, ".xlsx") > 0 Then
        regexPattern = Replace(regexPattern, ".xlsx", "\.(csv|xlsx)")
    Else
        regexPattern = regexPattern & "\.(csv|xlsx)"
    End If
    
    regexPattern = "^" & regexPattern & "$"
    
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .pattern = regexPattern
        .IgnoreCase = True
        .Global = False
    End With
    
    IsFileMatchFinalFormat = regex.Test(fileName)
    Exit Function
    
ErrorHandler:
    IsFileMatchFinalFormat = False
End Function

Function ExtractDateFromFinalFormat(fileName As String, finalSaveFormat As String, originalDateFormat As String, sftpName As String) As String
    On Error GoTo ErrorHandler
    
    Dim regex As Object
    Dim matches As Object
    
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .pattern = "(\d{8})"
        .IgnoreCase = True
        .Global = False
    End With
    
    Set matches = regex.Execute(fileName)
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

