Attribute VB_Name = "XLSMB_SFTP_Part1_Process"


Sub RenameCSVFiles_ParsedVersion()
    Dim ws As Worksheet, wb As Workbook
    Dim csvPath As Variant
    Dim refData As Variant
    Dim fileDialog As fileDialog
    Dim selectedFiles As FileDialogSelectedItems
    Dim originalFileName As String
    Dim sftpName As String
    Dim groupID As String
    Dim newFileName As String
    Dim i As Long
    Dim fso As Object
    Dim matchFound As Boolean
    Dim processedCount As Long
    Dim ErrorCount As Long
    Dim currentFolder As String
    Dim targetPath As String
    Dim sourceFile As String
    Dim extractedDate As String
    Dim originalDateFormat As String
    Dim zipFormatted As Boolean
    Dim apexApplied As Boolean
    Dim specificMacroExecuted As Boolean
    Dim genderFixed As Boolean
    
    Dim zipFormattedList As String, apexAppliedList As String
    Dim macroExecutedList As String, renamedList As String
    Dim originalsBackedUpList As String
    Dim genderFixedList As String
    
    ' *** NEW: Get and validate base folder path ***
    Dim baseFolderPath As String
    baseFolderPath = GetOrSetBaseFolderPath()
    
    If baseFolderPath = "" Then
        MsgBox "Operation cancelled. No base folder selected.", vbInformation, "Cancelled"
        Exit Sub
    End If
    
    refData = GetWorksheetData("Parsed_SFTPFiles")
    
    If IsEmpty(refData) Then
        MsgBox "Failed to load data from 'Parsed_SFTPFiles' worksheet. Please ensure the worksheet exists and contains data.", vbCritical, "Error"
        Exit Sub
    End If
    
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    With fileDialog
        .Title = "Select CSV Files to Format and Rename"
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
    
    zipFormattedList = "ZIP Formatting Applied To:" & vbCrLf
    apexAppliedList = "APEX Logic Applied To:" & vbCrLf
    macroExecutedList = "Specific Macros Executed For:" & vbCrLf
    renamedList = "Successfully Renamed:" & vbCrLf
    originalsBackedUpList = "Original Files Backed Up:" & vbCrLf
    genderFixedList = "Gender Field Fixed In:" & vbCrLf

    For Each csvPath In selectedFiles
        originalFileName = fso.GetFileName(csvPath)
        matchFound = False
        
        For i = 2 To UBound(refData, 1)
            If IsFileMatchPattern_ProcessRename_Fixed(originalFileName, CStr(refData(i, 1)), CStr(refData(i, 6))) Then
                sftpName = CStr(refData(i, 10))
                groupID = CStr(refData(i, 11))
                newFileName = CStr(refData(i, 13))
                
                originalDateFormat = CStr(refData(i, 6))
                extractedDate = ExtractAndConvertDate_FIXED_FINAL(originalFileName, originalDateFormat)
                
                If extractedDate <> "" Then
                    newFileName = Replace(newFileName, "mmddyyyy", extractedDate)
                End If
                
                ' === BACKUP ORIGINAL FILE TO DATE-BASED FOLDER ===
                Dim saveFolder As String
                Dim originalsFolder As String
                Dim originalBackupPath As String
                Dim updatedDateFormat As String
                Dim folderDateFormat As String
                Dim dateBasedFolder As String
                
                saveFolder = CStr(refData(i, 14))
                updatedDateFormat = CStr(refData(i, 12))
                
                ' *** MODIFIED: Use base folder path instead of OneDrive resolution ***
                saveFolder = ResolveFolderPath(saveFolder, baseFolderPath)
                saveFolder = Replace(saveFolder, "[Adjusted GroupName]", sftpName)
                saveFolder = Replace(saveFolder, "[Adjusted groupID]", groupID)
                saveFolder = Replace(saveFolder, "[GroupName]", sftpName)
                saveFolder = Replace(saveFolder, "[groupID]", groupID)
                
                If saveFolder <> "" And saveFolder <> "**NOT INCLUDED**" And InStr(saveFolder, "[") = 0 And extractedDate <> "" Then
                    folderDateFormat = FormatDateForFolder_Part1(extractedDate, updatedDateFormat)
                    
                    If folderDateFormat <> "" Then
                        dateBasedFolder = fso.BuildPath(saveFolder, folderDateFormat)
                        originalsFolder = fso.BuildPath(dateBasedFolder, "ORIGINALS")
                        
                        If Not fso.FolderExists(saveFolder) Then
                            CreateFolderPath fso, saveFolder
                        End If
                        If Not fso.FolderExists(dateBasedFolder) Then
                            fso.CreateFolder dateBasedFolder
                        End If
                        If Not fso.FolderExists(originalsFolder) Then
                            fso.CreateFolder originalsFolder
                        End If
                        
                        originalBackupPath = fso.BuildPath(originalsFolder, originalFileName)
                        On Error Resume Next
                        fso.CopyFile csvPath, originalBackupPath, True
                        If Err.Number = 0 Then
                            originalsBackedUpList = originalsBackedUpList & originalFileName & " -> " & folderDateFormat & "\ORIGINALS\" & vbCrLf
                        End If
                        On Error GoTo 0
                    End If
                End If
                
                On Error GoTo FileOpenError
                
                If LCase(fso.GetExtensionName(csvPath)) = "csv" Then
                    Application.DisplayAlerts = False
                    Workbooks.OpenText fileName:=csvPath, _
                        Origin:=xlMSDOS, _
                        StartRow:=1, _
                        DataType:=xlDelimited, _
                        TextQualifier:=xlDoubleQuote, _
                        ConsecutiveDelimiter:=False, _
                        Tab:=False, _
                        Semicolon:=False, _
                        Comma:=True, _
                        Space:=False, _
                        Other:=False
                    Application.DisplayAlerts = True
                    Set wb = ActiveWorkbook
                Else
                    Set wb = Workbooks.Open(csvPath)
                End If
                
                Set ws = wb.Sheets(1)
                On Error GoTo 0
                
                zipFormatted = ApplyZipFormatting_ProcessRename_Fixed(ws)
                If zipFormatted Then zipFormattedList = zipFormattedList & originalFileName & vbCrLf
                
                genderFixed = ApplyUniversalGenderFix(ws)
                If genderFixed Then genderFixedList = genderFixedList & originalFileName & vbCrLf
                
                ' Extract FileType from refData Column O (index 15)
                Dim sFileType As String
                sFileType = Trim(CStr(refData(i, 15)))

                ' Apply duplicate removal logic to ALL files (not just Apex)
               Dim duplicatesRemoved As Boolean
                duplicatesRemoved = ApplyDuplicateRemovalLogic(ws, sFileType)
                If duplicatesRemoved Then apexAppliedList = apexAppliedList & originalFileName & vbCrLf
                
                specificMacroExecuted = ExecuteSpecificMacro_ProcessRename_Fixed(ws, sftpName, CStr(refData(i, 8)))
                If specificMacroExecuted Then macroExecutedList = macroExecutedList & originalFileName & " (" & sftpName & " specific macro)" & vbCrLf
                
                On Error Resume Next
                Application.DisplayAlerts = False
                
                currentFolder = fso.GetParentFolderName(csvPath)
                
                Dim todayDate As String
                todayDate = Format(Date, "mmddyyyy")
                
                Dim subfolderName As String
                subfolderName = todayDate & " SFTP files"
                
                Dim subfolderPath As String
                subfolderPath = fso.BuildPath(currentFolder, subfolderName)
                
                If Not fso.FolderExists(subfolderPath) Then
                    fso.CreateFolder (subfolderPath)
                End If
                
                targetPath = fso.BuildPath(subfolderPath, newFileName)
                
                wb.SaveAs fileName:=targetPath, FileFormat:=xlCSV
                
                Application.DisplayAlerts = True
                On Error GoTo 0
                wb.Close
                
                processedCount = processedCount + 1
                renamedList = renamedList & originalFileName & " ? " & newFileName & vbCrLf
                matchFound = True
                Exit For
                
FileOpenError:
                Application.DisplayAlerts = True
                MsgBox "Error opening file: " & originalFileName & vbCrLf & "Error: " & Err.Description, vbExclamation, "File Open Error"
                ErrorCount = ErrorCount + 1
                On Error GoTo 0
                GoTo NextFile
            End If
        Next i
        
        If Not matchFound Then
            MsgBox "No matching pattern found for file: " & originalFileName & vbCrLf & _
                   "Please check the Parsed_SFTPFiles worksheet for a matching pattern.", vbExclamation, "Warning"
            ErrorCount = ErrorCount + 1
        End If
        
NextFile:
    Next csvPath

    Dim summaryMsg As String
    summaryMsg = "FORMATTING & RENAMING COMPLETE!" & vbCrLf & vbCrLf & _
                 originalsBackedUpList & vbCrLf & _
                 genderFixedList & vbCrLf & _
                 zipFormattedList & vbCrLf & _
                 apexAppliedList & vbCrLf & _
                 macroExecutedList & vbCrLf & _
                 renamedList & vbCrLf & _
                 "Files processed: " & processedCount & vbCrLf & _
                 "Errors: " & ErrorCount
    
    MsgBox summaryMsg, vbInformation, "Processing Summary"
    Exit Sub

FileError:
    MsgBox "Error renaming file: " & originalFileName & vbCrLf & _
           "Target: " & newFileName & vbCrLf & _
           "Error: " & Err.Description, vbCritical, "File Error"
    ErrorCount = ErrorCount + 1
    On Error GoTo 0
    Resume NextFile
End Sub


' ===== MODIFIED: Resolve folder path using base folder =====
Function ResolveFolderPath(folderPath As String, baseFolderPath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim cleanPath As String
    cleanPath = folderPath
    
    ' Remove placeholders
    cleanPath = Replace(cleanPath, "{OneDriveCommercial}\", "")
    cleanPath = Replace(cleanPath, "{OneDrive}\", "")
    cleanPath = Replace(cleanPath, "{UserProfile}\", "")
    
    ' Build path from base
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

Function ApplyZipFormatting_ProcessRename_Fixed(ws As Worksheet) As Boolean
    Dim keywordList As Variant, keyword As Variant
    Dim headerCell As Range
    Dim cleanHeader As String
    Dim j As Long, lastCol As Long
    Dim zipFormatted As Boolean
    
    keywordList = Array("zip", "zipcode", "zip code", "postalcode", "postal code")
    zipFormatted = False
    
    On Error Resume Next
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For j = 1 To lastCol
        Set headerCell = ws.Cells(1, j)
        cleanHeader = LCase(Replace(Replace(Replace(Trim(headerCell.Value), "_", ""), "-", ""), " ", ""))
        For Each keyword In keywordList
            If InStr(cleanHeader, Replace(LCase(keyword), " ", "")) > 0 Then
                ws.Columns(j).NumberFormat = "@"
                
                Dim lastRow As Long
                lastRow = ws.Cells(ws.Rows.count, j).End(xlUp).row
                Dim i As Long
                For i = 2 To lastRow
                    If ws.Cells(i, j).Value <> "" And IsNumeric(ws.Cells(i, j).Value) Then
                        ws.Cells(i, j).Value = Format(ws.Cells(i, j).Value, "00000")
                    End If
                Next i
                
                zipFormatted = True
                Exit For
            End If
        Next keyword
    Next j
    On Error GoTo 0
    
    ApplyZipFormatting_ProcessRename_Fixed = zipFormatted
End Function

Function ApplyUniversalGenderFix(ws As Worksheet) As Boolean
    Dim genderCol As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim j As Long
    Dim fixedCount As Long
    
    On Error Resume Next
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    
    genderCol = 0
    For j = 1 To lastCol
        Dim headerText As String
        headerText = UCase(Trim(ws.Cells(1, j).Value))
        
        If headerText = "GENDER" Or _
           headerText = "SEX" Or _
           headerText = "MEMBER GENDER" Or _
           headerText = "MEMBERGENDER" Or _
           InStr(headerText, "GENDER") > 0 Then
            genderCol = j
            Exit For
        End If
    Next j
    
    If genderCol > 0 Then
        fixedCount = 0
        For i = 2 To lastRow
            If Trim(ws.Cells(i, genderCol).Value) = "" Then
                ws.Cells(i, genderCol).Value = "M"
                fixedCount = fixedCount + 1
            End If
        Next i
        
        ApplyUniversalGenderFix = (fixedCount > 0)
    Else
        ApplyUniversalGenderFix = False
    End If
    
    On Error GoTo 0
End Function

Function ExtractAndConvertDate_FIXED_FINAL(fileName As String, dateFormat As String) As String
    On Error GoTo ErrorHandler
    
    Dim regex As Object
    Dim matches As Object
    Dim extractedDate As String
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True
    regex.Global = True
    
    Select Case LCase(Trim(dateFormat))
        Case "mmddyy"
            regex.pattern = "\d{6}"
        Case "mmddyyyy"
            regex.pattern = "\d{8}"
        Case "yyyymmdd"
            regex.pattern = "\d{8}"
        Case Else
            ExtractAndConvertDate_FIXED_FINAL = ""
            Exit Function
    End Select
    
    Set matches = regex.Execute(fileName)
    
    If matches.count > 0 Then
        If LCase(Trim(dateFormat)) = "mmddyy" Then
            extractedDate = matches(matches.count - 1).Value
            If Len(extractedDate) = 6 Then
                ExtractAndConvertDate_FIXED_FINAL = Left(extractedDate, 4) & "20" & Right(extractedDate, 2)
            Else
                ExtractAndConvertDate_FIXED_FINAL = extractedDate
            End If
        ElseIf LCase(Trim(dateFormat)) = "mmddyyyy" Then
            extractedDate = matches(matches.count - 1).Value
            ExtractAndConvertDate_FIXED_FINAL = extractedDate
        ElseIf LCase(Trim(dateFormat)) = "yyyymmdd" Then
            extractedDate = matches(matches.count - 1).Value
            If Len(extractedDate) = 8 And Left(extractedDate, 2) = "20" Then
                ExtractAndConvertDate_FIXED_FINAL = Mid(extractedDate, 5, 2) & Right(extractedDate, 2) & Left(extractedDate, 4)
            Else
                ExtractAndConvertDate_FIXED_FINAL = extractedDate
            End If
        End If
    Else
        ExtractAndConvertDate_FIXED_FINAL = ""
    End If
    Exit Function
    
ErrorHandler:
    ExtractAndConvertDate_FIXED_FINAL = ""
End Function

Function IsFileMatchPattern_ProcessRename_Fixed(fileName As String, csvPattern As String, dateFormat As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim regex As Object
    Dim regexPattern As String
    
    regexPattern = csvPattern
    
    Select Case LCase(Trim(dateFormat))
        Case "mmddyyyy"
            regexPattern = Replace(regexPattern, "mmddyyyy", "\d{8}")
        Case "mmddyy"
            regexPattern = Replace(regexPattern, "mmddyy", "\d{6}")
        Case "yyyymmdd"
            regexPattern = Replace(regexPattern, "yyyymmdd", "\d{8}")
    End Select
    
    regexPattern = Replace(regexPattern, " ", "[\s_]")
    regexPattern = Replace(regexPattern, "-", "[-_]")
    regexPattern = "^" & regexPattern & "$"
    
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .pattern = regexPattern
        .IgnoreCase = True
        .Global = False
    End With
    
    IsFileMatchPattern_ProcessRename_Fixed = regex.Test(fileName)
    Exit Function
    
ErrorHandler:
    IsFileMatchPattern_ProcessRename_Fixed = False
End Function

Function ExecuteSpecificMacro_ProcessRename_Fixed(ws As Worksheet, sftpName As String, specificMacroPath As String) As Boolean
    On Error GoTo MacroError
    
    Debug.Print "=== ExecuteSpecificMacro_ProcessRename_Fixed DEBUG ==="
    Debug.Print "sftpName: " & sftpName
    Debug.Print "specificMacroPath: " & specificMacroPath
    
    If specificMacroPath <> "" And Len(Trim(specificMacroPath)) > 0 Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim moduleName As String
        moduleName = fso.GetBaseName(specificMacroPath)
        
        Debug.Print "Extracted module name: " & moduleName
        
        On Error Resume Next
        Debug.Print "Attempting to call: " & moduleName & ".ProcessCSVFiles"
        Application.Run moduleName & ".ProcessCSVFiles"
        
        If Err.Number = 0 Then
            Debug.Print "SUCCESS: Called " & moduleName & ".ProcessCSVFiles"
            ExecuteSpecificMacro_ProcessRename_Fixed = True
        Else
            Debug.Print "FAILED: " & moduleName & ".ProcessCSVFiles - Error: " & Err.Description
            Err.Clear
            Debug.Print "Attempting to call: " & moduleName & ".Main"
            Application.Run moduleName & ".Main"
            If Err.Number = 0 Then
                Debug.Print "SUCCESS: Called " & moduleName & ".Main"
                ExecuteSpecificMacro_ProcessRename_Fixed = True
            Else
                Debug.Print "FAILED: " & moduleName & ".Main - Error: " & Err.Description
                Err.Clear
                Debug.Print "Attempting to call: " & moduleName & "." & moduleName
                Application.Run moduleName & "." & moduleName
                If Err.Number = 0 Then
                    Debug.Print "SUCCESS: Called " & moduleName & "." & moduleName
                    ExecuteSpecificMacro_ProcessRename_Fixed = True
                Else
                    Debug.Print "FAILED: All attempts failed for module: " & moduleName
                    ExecuteSpecificMacro_ProcessRename_Fixed = False
                End If
            End If
        End If
        On Error GoTo 0
        
    Else
        Debug.Print "No specific macro path provided, checking sftpName cases"
        Select Case sftpName
            Case "HealthKarma_ThinBlueLine"
                Debug.Print "Calling ColumnFix_Integrated for HealthKarma_ThinBlueLine"
                ColumnFix_Integrated ws
                ExecuteSpecificMacro_ProcessRename_Fixed = True
            Case Else
                Debug.Print "No matching sftpName case found"
                ExecuteSpecificMacro_ProcessRename_Fixed = False
        End Select
    End If
    
    Debug.Print "ExecuteSpecificMacro_ProcessRename_Fixed result: " & ExecuteSpecificMacro_ProcessRename_Fixed
    Debug.Print "=== END DEBUG ==="
    Exit Function
    
MacroError:
    Debug.Print "MacroError: " & Err.Description
    ExecuteSpecificMacro_ProcessRename_Fixed = False
End Function



' Helper function to find column by header name
Private Function GetColumnByHeader(ws As Worksheet, headerName As String) As Long
    Dim lastCol As Long
    Dim i As Long
    
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    
    For i = 1 To lastCol
        If Trim(UCase(ws.Cells(1, i).Value)) = Trim(UCase(headerName)) Then
            GetColumnByHeader = i
            Exit Function
        End If
    Next i
    
    GetColumnByHeader = 0 ' Not found
End Function

Sub CleanAndFormatLocationColumn_Integrated(ws As Worksheet)
    Dim headerRow As Range
    Dim cell As Range
    Dim locationCol As Integer
    Dim lastRow As Long
    Dim i As Long
    Dim originalVal As String
    Dim cleanedVal As String
    Dim regex As Object

    Set headerRow = ws.Rows(1)
    locationCol = 0

    For Each cell In headerRow.Cells
        If Trim(cell.Value) = "Location" Then
            locationCol = cell.Column
            Exit For
        End If
    Next cell

    If locationCol = 0 Then Exit Sub

    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = "[^A-Za-z0-9]"
    regex.Global = True

    lastRow = ws.Cells(ws.Rows.count, locationCol).End(xlUp).row

    For i = 2 To lastRow
        originalVal = ws.Cells(i, locationCol).Text
        cleanedVal = regex.Replace(originalVal, "")
        
        If IsNumeric(cleanedVal) Then
            ws.Cells(i, locationCol).Value = Val(cleanedVal)
        Else
            ws.Cells(i, locationCol).Value = ""
        End If
    Next i

    ws.Columns(locationCol).NumberFormat = "#000"
End Sub

Sub FixPEOIDToMatchGroupID_Integrated(ws As Worksheet, sftpName As String)
    Dim lastCol As Long, lastRow As Long, colPEO As Long
    Dim i As Long
    Dim expectedGroupID As String
    
    Select Case sftpName
        Case "PrismHR_AmericanBenefitCompany"
            expectedGroupID = "778182"
        Case "PrismHR_DeltaAdministrators"
            expectedGroupID = "757422"
        Case Else
            Exit Sub
    End Select

    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row

    colPEO = 0
    For i = 1 To lastCol
        If Trim(ws.Cells(1, i).Value) = "PEO ID" Then
            colPEO = i
            Exit For
        End If
    Next i

    If colPEO = 0 Then Exit Sub

    For i = 2 To lastRow
        If Trim(ws.Cells(i, colPEO).Value) <> expectedGroupID Then
            ws.Cells(i, colPEO).Value = expectedGroupID
        End If
    Next i
End Sub

Sub ColumnFix_Integrated(ws As Worksheet)
    Dim headerDict As Object
    Dim desiredOrder As Variant
    Dim i As Long
    Dim currentHeader As String
    Dim tempSheet As Worksheet

    Set headerDict = CreateObject("Scripting.Dictionary")

    For i = 1 To ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
        currentHeader = Trim(ws.Cells(1, i).Value)
        If LCase(currentHeader) = "internal code" Then
            currentHeader = "MetaTag1"
        End If
        If Not headerDict.Exists(LCase(currentHeader)) Then
            headerDict(LCase(currentHeader)) = i
        End If
    Next i

    desiredOrder = Array( _
        "LastName", "FirstName", "Gender", "DateOfBirth", "AddressLine1", "AddressLine2", _
        "City", "State", "ZipCode", "CountryCode", "MobilePhone", "EmailAddress", _
        "EffectiveStart", "EffectiveEnd", "MemberType", "ClientMemberID", _
        "SecondaryClientMemberID", "ClientPrimaryMemberID", "ServiceOffering", _
        "GroupID", "GroupName", "MetaTag1", "MetaTag2", "MetaTag3", "MetaTag4", "MetaTag5" _
    )

    Set tempSheet = ws.Parent.Worksheets.Add
    tempSheet.Name = "TempSheet12345"

    For i = 0 To UBound(desiredOrder)
        currentHeader = LCase(desiredOrder(i))
        If headerDict.Exists(currentHeader) Then
            ws.Columns(headerDict(currentHeader)).Copy Destination:=tempSheet.Columns(i + 1)
        Else
            tempSheet.Cells(1, i + 1).Value = desiredOrder(i)
        End If
    Next i

    ws.Cells.Clear
    tempSheet.UsedRange.Copy Destination:=ws.Range("A1")
    Application.DisplayAlerts = False
    tempSheet.Delete
    Application.DisplayAlerts = True
End Sub

Sub GenderFix_Integrated(ws As Worksheet)
    Dim genderCol As Long
    Dim lastRow As Long
    Dim i As Long

    genderCol = 0
    For i = 1 To ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
        If Trim(UCase(ws.Cells(1, i).Value)) = "GENDER" Then
            genderCol = i
            Exit For
        End If
    Next i

    If genderCol = 0 Then Exit Sub

    lastRow = ws.Cells(ws.Rows.count, genderCol).End(xlUp).row

    For i = 2 To lastRow
        If Trim(ws.Cells(i, genderCol).Value) = "" Then
            ws.Cells(i, genderCol).Value = "M"
        End If
    Next i
End Sub

Sub CreateFolderPath(fso As Object, fullPath As String)
    Dim pathParts() As String
    Dim currentPath As String
    Dim i As Integer
    
    pathParts = Split(fullPath, "\")
    currentPath = pathParts(0)
    
    For i = 1 To UBound(pathParts)
        currentPath = currentPath & "\" & pathParts(i)
        If Not fso.FolderExists(currentPath) Then
            On Error Resume Next
            fso.CreateFolder currentPath
            On Error GoTo 0
        End If
    Next i
End Sub

Function FormatDateForFolder_Part1(dateString As String, updatedDateFormat As String) As String
    On Error GoTo ErrorHandler
    
    Dim yearPart As String, monthPart As String, dayPart As String
    Dim monthNames As Variant
    Dim monthIndex As Integer
    
    monthNames = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", _
                      "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    
    If Len(dateString) <> 8 Then
        FormatDateForFolder_Part1 = ""
        Exit Function
    End If
    
    Dim actualFormat As String
    Dim firstTwo As Integer
    firstTwo = CInt(Left(dateString, 2))
    
    If firstTwo > 12 Then
        actualFormat = "yyyymmdd"
    Else
        actualFormat = "mmddyyyy"
    End If
    
    Select Case actualFormat
        Case "mmddyyyy"
            monthPart = Left(dateString, 2)
            dayPart = Mid(dateString, 3, 2)
            yearPart = Right(dateString, 2)
        Case "yyyymmdd"
            yearPart = Right(Left(dateString, 4), 2)
            monthPart = Mid(dateString, 5, 2)
            dayPart = Right(dateString, 2)
        Case Else
            FormatDateForFolder_Part1 = ""
            Exit Function
    End Select
    
    monthIndex = CInt(monthPart) - 1
    If monthIndex < 0 Or monthIndex > 11 Then
        FormatDateForFolder_Part1 = ""
        Exit Function
    End If
    
    FormatDateForFolder_Part1 = Format(monthPart, "00") & monthNames(monthIndex) & yearPart
    Exit Function
    
ErrorHandler:
    FormatDateForFolder_Part1 = ""
End Function
Function GetOrSetBaseFolderPath() As String
    Dim basePath As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Try OneDrive environment variable
    basePath = Environ("OneDrive")
    If basePath <> "" Then
        basePath = basePath & "\Documents - Ops"
        If fso.FolderExists(basePath) Then
            GetOrSetBaseFolderPath = basePath
            Exit Function
        End If
    End If
    
    ' Try OneDrive - Recuro Health
    basePath = Environ("USERPROFILE") & "\OneDrive - Recuro Health\Documents - Ops"
    If fso.FolderExists(basePath) Then
        GetOrSetBaseFolderPath = basePath
        Exit Function
    End If
    
    ' Try OneDriveCommercial
    basePath = Environ("OneDriveCommercial")
    If basePath <> "" Then
        basePath = basePath & "\Documents - Ops"
        If fso.FolderExists(basePath) Then
            GetOrSetBaseFolderPath = basePath
            Exit Function
        End If
    End If
    
    MsgBox "Could not find OneDrive folder at:" & vbCrLf & _
           Environ("USERPROFILE") & "\OneDrive - Recuro Health\Documents - Ops", _
           vbCritical, "OneDrive Path Not Found"
    GetOrSetBaseFolderPath = ""
End Function
Function GetColumnMappingForProcessing(sFileType As String) As ValidationEngine.ColumnMapping
    Dim oMapping As ValidationEngine.ColumnMapping
    Dim wsMapping As Worksheet
    Dim lLastRow As Long
    Dim lRow As Long
    
    On Error GoTo MappingError
    
    Set wsMapping = ThisWorkbook.Worksheets("Filetype Mapping")
    lLastRow = wsMapping.Cells(wsMapping.Rows.count, "A").End(xlUp).row
    
    ' Initialize with zero values
    oMapping.fileType = ""
    oMapping.memberID = 0
    oMapping.serviceOffering = 0
    oMapping.EffectiveDate = 0
    oMapping.effectiveEndDate = 0
    
    ' Find the FileType row in Filetype Mapping sheet
    For lRow = 2 To lLastRow
        If UCase(Trim(wsMapping.Cells(lRow, "A").Value)) = UCase(Trim(sFileType)) Then
            With oMapping
                .fileType = sFileType
                .memberID = wsMapping.Cells(lRow, "M").Value        ' Column M
                .serviceOffering = wsMapping.Cells(lRow, "L").Value ' Column L
                .EffectiveDate = wsMapping.Cells(lRow, "J").Value   ' Column J (EffectiveStart)
                .effectiveEndDate = wsMapping.Cells(lRow, "N").Value ' Column N
            End With
            
            GetColumnMappingForProcessing = oMapping
            Exit Function
        End If
    Next lRow
    
    ' FileType not found
    Debug.Print "WARNING: FileType '" & sFileType & "' not found in Filetype Mapping"
    GetColumnMappingForProcessing = oMapping
    Exit Function
    
MappingError:
    Debug.Print "ERROR in GetColumnMappingForProcessing: " & Err.Description
    GetColumnMappingForProcessing = oMapping
End Function
Function ApplyDuplicateRemovalLogic(ws As Worksheet, sFileType As String) As Boolean
    On Error GoTo DuplicateError
    
    Dim oMapping As ValidationEngine.ColumnMapping
    Dim dict As Object
    Dim cell As Range
    Dim i As Long, lastRow As Long
    Dim sMemberID As String
    Dim sServiceOffering As String
    Dim sCombinedKey As String
    Dim sEffectiveEndDate As String
    Dim dEffectiveStart As Double
    Dim colMemberID As Integer
    Dim colServiceOffering As Integer
    Dim colEffectiveStart As Integer
    Dim colEffectiveEndDate As Integer
    
    ' Get column mapping for this FileType
    oMapping = GetColumnMappingForProcessing(sFileType)
    
    ' Validate that we have the required column mappings
    If oMapping.fileType = "" Or oMapping.memberID = 0 Or oMapping.serviceOffering = 0 Then
        Debug.Print "WARNING: Incomplete mapping for FileType '" & sFileType & "' - skipping duplicate removal"
        ApplyDuplicateRemovalLogic = False
        Exit Function
    End If
    
    ' Get column numbers from mapping
    colMemberID = oMapping.memberID
    colServiceOffering = oMapping.serviceOffering
    colEffectiveStart = oMapping.EffectiveDate
    colEffectiveEndDate = oMapping.effectiveEndDate
    
    ' Verify columns exist in worksheet
    If colMemberID > ws.Cells(1, ws.Columns.count).End(xlToLeft).Column Or _
       colServiceOffering > ws.Cells(1, ws.Columns.count).End(xlToLeft).Column Then
        Debug.Print "WARNING: Required columns not found in file - skipping duplicate removal"
        ApplyDuplicateRemovalLogic = False
        Exit Function
    End If
    
    lastRow = ws.Cells(ws.Rows.count, colMemberID).End(xlUp).row
    If lastRow < 2 Then
        ApplyDuplicateRemovalLogic = False
        Exit Function
    End If
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' ==========================================================================
    ' STEP 1: Build dictionary of MemberID + ServiceOffering combinations
    ' Count duplicates
    ' ==========================================================================
    For i = 2 To lastRow
        sMemberID = Trim(CStr(ws.Cells(i, colMemberID).Value))
        sServiceOffering = Trim(CStr(ws.Cells(i, colServiceOffering).Value))
        
        ' Create combined key: MemberID + "|" + ServiceOffering
        sCombinedKey = sMemberID & "|" & sServiceOffering
        
        If Not dict.Exists(sCombinedKey) Then
            dict.Add sCombinedKey, 1
        Else
            dict(sCombinedKey) = dict(sCombinedKey) + 1
        End If
    Next i
    
    ' ==========================================================================
    ' STEP 2: Delete rows where combination is duplicated AND has EffectiveEndDate
    ' (i.e., remove inactive/terminated records when duplicates exist)
    ' ==========================================================================
    For i = lastRow To 2 Step -1
        sMemberID = Trim(CStr(ws.Cells(i, colMemberID).Value))
        sServiceOffering = Trim(CStr(ws.Cells(i, colServiceOffering).Value))
        sCombinedKey = sMemberID & "|" & sServiceOffering
        
        ' Check if this is a duplicate combination
        If dict.Exists(sCombinedKey) And dict(sCombinedKey) > 1 Then
            ' Check if this record has an EffectiveEndDate (is inactive)
            If colEffectiveEndDate > 0 Then
                sEffectiveEndDate = Trim(CStr(ws.Cells(i, colEffectiveEndDate).Value))
                
                ' If EffectiveEndDate has a value, delete this row
                If sEffectiveEndDate <> "" Then
                    ws.Rows(i).Delete
                End If
            End If
        End If
    Next i
    
    ' ==========================================================================
    ' STEP 3: For remaining duplicates, keep the one with the LARGER EffectiveStart
    ' (i.e., keep the more recent enrollment)
    ' ==========================================================================
    Set dict = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Rows.count, colMemberID).End(xlUp).row
    
    For i = 2 To lastRow
        sMemberID = Trim(CStr(ws.Cells(i, colMemberID).Value))
        sServiceOffering = Trim(CStr(ws.Cells(i, colServiceOffering).Value))
        sCombinedKey = sMemberID & "|" & sServiceOffering
        
        If Not dict.Exists(sCombinedKey) Then
            ' First occurrence - store row number
            dict.Add sCombinedKey, i
        Else
            ' Duplicate found - compare EffectiveStart dates
            Dim storedRow As Long
            Dim currentEffectiveStart As Double
            Dim storedEffectiveStart As Double
            
            storedRow = dict(sCombinedKey)
            
            ' Get EffectiveStart values (convert to numeric for comparison)
            On Error Resume Next
            currentEffectiveStart = CDbl(ws.Cells(i, colEffectiveStart).Value)
            storedEffectiveStart = CDbl(ws.Cells(storedRow, colEffectiveStart).Value)
            On Error GoTo 0
            
            ' Keep row with LARGER (more recent) EffectiveStart
            If currentEffectiveStart < storedEffectiveStart Then
                ' Current row is older - delete it
                ws.Rows(i).Delete
                lastRow = lastRow - 1
                i = i - 1
            Else
                ' Stored row is older - delete it and update dictionary
                ws.Rows(storedRow).Delete
                lastRow = lastRow - 1
                ' Update dictionary with current row (adjusted for deletion)
                If i > storedRow Then
                    dict(sCombinedKey) = i - 1
                Else
                    dict(sCombinedKey) = i
                End If
            End If
        End If
    Next i
    
    ApplyDuplicateRemovalLogic = True
    Exit Function
    
DuplicateError:
    Debug.Print "ERROR in ApplyDuplicateRemovalLogic: " & Err.Description
    ApplyDuplicateRemovalLogic = False
End Function


Function RemoveDuplicates_BasedOnMemberIDAndServiceOffering(ws As Worksheet, fileTypeCode As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' This function removes duplicate records based on:
    ' 1. MemberID (Column P) + ServiceOffering (Column S) matching
    ' 2. If one is inactive (has EffectiveEnd), remove it
    ' 3. If BOTH are active, DON'T remove - let validator catch it as error
    
    Dim lastRow As Long
    Dim i As Long
    Dim dict As Object
    Dim memberKey As String
    Dim memberID As String
    Dim serviceOffering As String
    Dim effectiveEnd As String
    Dim effectiveStart As String
    Dim removedCount As Long
    
    removedCount = 0
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Get column positions from header row
    Dim colMemberID As Long, colServiceOffering As Long
    Dim colEffectiveEnd As Long, colEffectiveStart As Long
    
    colMemberID = GetColumnByHeader(ws, "ClientMemberID")
    colServiceOffering = GetColumnByHeader(ws, "ServiceOffering")
    colEffectiveEnd = GetColumnByHeader(ws, "EffectiveEnd")
    colEffectiveStart = GetColumnByHeader(ws, "EffectiveStart")
    
    ' Validate that required columns exist
    If colMemberID = 0 Or colServiceOffering = 0 Then
        Debug.Print "RemoveDuplicates: Required columns not found. MemberID=" & colMemberID & ", ServiceOffering=" & colServiceOffering
        RemoveDuplicates_BasedOnMemberIDAndServiceOffering = False
        Exit Function
    End If
    
    lastRow = ws.Cells(ws.Rows.count, colMemberID).End(xlUp).row
    If lastRow < 2 Then
        RemoveDuplicates_BasedOnMemberIDAndServiceOffering = True
        Exit Function
    End If
    
    Debug.Print "=== DUPLICATE REMOVAL STARTED ==="
    Debug.Print "Total rows to process: " & (lastRow - 1)
    
    ' PASS 1: Build dictionary of all records with their active status
    ' Key format: "MemberID|ServiceOffering"
    ' Value format: Row number|IsActive|EffectiveStart
    
    For i = 2 To lastRow
        memberID = Trim(CStr(ws.Cells(i, colMemberID).Value))
        serviceOffering = Trim(CStr(ws.Cells(i, colServiceOffering).Value))
        
        If memberID <> "" And serviceOffering <> "" Then
            memberKey = memberID & "|" & serviceOffering
            
            effectiveEnd = ""
            effectiveStart = ""
            
            If colEffectiveEnd > 0 Then
                effectiveEnd = Trim(CStr(ws.Cells(i, colEffectiveEnd).Value))
            End If
            
            If colEffectiveStart > 0 Then
                effectiveStart = Trim(CStr(ws.Cells(i, colEffectiveStart).Value))
            End If
            
            ' Determine if record is active
            Dim isActive As Boolean
            isActive = IsRecordActive(effectiveEnd)
            
            ' Store: "RowNumber|IsActive|EffectiveStart"
            Dim rowData As String
            rowData = i & "|" & isActive & "|" & effectiveStart
            
            ' If key doesn't exist, add it
            If Not dict.Exists(memberKey) Then
                ' Create new collection for this memberKey
                Dim newCollection As Collection
                Set newCollection = New Collection
                newCollection.Add rowData
                dict.Add memberKey, newCollection
            Else
                ' Add to existing collection
                dict(memberKey).Add rowData
            End If
        End If
    Next i
    
    Debug.Print "Unique MemberID+ServiceOffering combinations found: " & dict.count
    
    ' PASS 2: Process duplicates
    ' Work backwards through rows to safely delete
    Dim rowsToDelete As Collection
    Set rowsToDelete = New Collection
    
    Dim key As Variant
    For Each key In dict.Keys
        Dim recordCollection As Collection
        Set recordCollection = dict(key)
        
        If recordCollection.count > 1 Then
            ' We have duplicates for this MemberID + ServiceOffering combo
            Debug.Print "Processing duplicates for: " & key & " (Count: " & recordCollection.count & ")"
            
            ' Analyze the records
            Dim activeCount As Long
            Dim inactiveCount As Long
            Dim activeRows As Collection
            Dim inactiveRows As Collection
            
            Set activeRows = New Collection
            Set inactiveRows = New Collection
            activeCount = 0
            inactiveCount = 0
            
            Dim j As Long
            For j = 1 To recordCollection.count
                Dim rowInfo As String
                rowInfo = recordCollection(j)
                
                Dim rowParts() As String
                rowParts = Split(rowInfo, "|")
                
                Dim rowNum As Long
                Dim rowIsActive As Boolean
                rowNum = CLng(rowParts(0))
                rowIsActive = CBool(rowParts(1))
                
                If rowIsActive Then
                    activeCount = activeCount + 1
                    activeRows.Add rowNum
                Else
                    inactiveCount = inactiveCount + 1
                    inactiveRows.Add rowNum
                End If
            Next j
            
            Debug.Print "  Active: " & activeCount & ", Inactive: " & inactiveCount
            
            ' DECISION LOGIC:
            ' 1. If multiple active records exist, DON'T remove any - let validator flag it
            ' 2. If only 1 active exists, remove all inactive records
            ' 3. If all are inactive, keep the one with most recent start date
            
            If activeCount >= 2 Then
                ' Multiple active records - DO NOT REMOVE
                ' Let the validator catch this as an error
                Debug.Print "  ? Multiple active records detected. Keeping all for validator to flag."
                
            ElseIf activeCount = 1 And inactiveCount > 0 Then
                ' One active, one or more inactive - Remove the inactive ones
                Debug.Print "  ? Removing " & inactiveCount & " inactive record(s)"
                
                For j = 1 To inactiveRows.count
                    rowsToDelete.Add inactiveRows(j)
                    removedCount = removedCount + 1
                Next j
                
            ElseIf activeCount = 0 And inactiveCount > 1 Then
                ' All inactive - Keep the one with most recent start date, remove others
                Debug.Print "  ? All records inactive. Keeping most recent, removing others."
                
                ' Find row with most recent start date
                Dim mostRecentRow As Long
                Dim mostRecentDate As Date
                mostRecentDate = DateSerial(1900, 1, 1) ' Very old date
                
                For j = 1 To recordCollection.count
                    rowInfo = recordCollection(j)
                    rowParts = Split(rowInfo, "|")
                    rowNum = CLng(rowParts(0))
                    
                    Dim startDateStr As String
                    startDateStr = rowParts(2)
                    
                    If IsDate(startDateStr) Then
                        Dim startDate As Date
                        startDate = CDate(startDateStr)
                        
                        If startDate > mostRecentDate Then
                            ' If we had a previous "most recent", mark it for deletion
                            If mostRecentRow > 0 Then
                                rowsToDelete.Add mostRecentRow
                            End If
                            
                            mostRecentDate = startDate
                            mostRecentRow = rowNum
                        Else
                            ' This row is older, mark for deletion
                            rowsToDelete.Add rowNum
                        End If
                    End If
                Next j
                
                removedCount = removedCount + (inactiveCount - 1)
            End If
        End If
    Next key
    
    ' PASS 3: Delete rows (working backwards to maintain row numbers)
    If rowsToDelete.count > 0 Then
        Debug.Print "Deleting " & rowsToDelete.count & " duplicate rows..."
        
        ' Sort rows in descending order for safe deletion
        Dim sortedRows() As Long
        ReDim sortedRows(1 To rowsToDelete.count)
        
        For i = 1 To rowsToDelete.count
            sortedRows(i) = rowsToDelete(i)
        Next i
        
        ' Simple bubble sort (descending)
        Dim temp As Long
        Dim swapped As Boolean
        Do
            swapped = False
            For i = 1 To UBound(sortedRows) - 1
                If sortedRows(i) < sortedRows(i + 1) Then
                    temp = sortedRows(i)
                    sortedRows(i) = sortedRows(i + 1)
                    sortedRows(i + 1) = temp
                    swapped = True
                End If
            Next i
        Loop While swapped
        
        ' Delete rows from highest to lowest
        For i = 1 To UBound(sortedRows)
            ws.Rows(sortedRows(i)).Delete
            Debug.Print "  Deleted row " & sortedRows(i)
        Next i
    End If
    
    Debug.Print "=== DUPLICATE REMOVAL COMPLETE ==="
    Debug.Print "Total rows removed: " & removedCount
    
    RemoveDuplicates_BasedOnMemberIDAndServiceOffering = True
    Exit Function
    
ErrorHandler:
    Debug.Print "ERROR in RemoveDuplicates: " & Err.Description
    RemoveDuplicates_BasedOnMemberIDAndServiceOffering = False
End Function

' Helper function to determine if a record is active
Private Function IsRecordActive(effectiveEndDate As String) As Boolean
    ' A record is active if:
    ' 1. EffectiveEnd is blank/empty, OR
    ' 2. EffectiveEnd is a future date (after today)
    
    If Trim(effectiveEndDate) = "" Then
        IsRecordActive = True
        Exit Function
    End If
    
    If IsDate(effectiveEndDate) Then
        Dim endDate As Date
        Dim todayDate As Date
        
        endDate = CDate(effectiveEndDate)
        todayDate = Date ' Today's date
        
        ' Active if end date is in the future
        IsRecordActive = (endDate >= todayDate)
    Else
        ' If we can't parse the date, assume active
        IsRecordActive = True
    End If
End Function

' Helper function to find column by header name
Private Function GetColumnByHeader(ws As Worksheet, headerName As String) As Long
    Dim lastCol As Long
    Dim i As Long
    
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    
    For i = 1 To lastCol
        If Trim(UCase(ws.Cells(1, i).Value)) = Trim(UCase(headerName)) Then
            GetColumnByHeader = i
            Exit Function
        End If
    Next i
    
    GetColumnByHeader = 0 ' Not found
End Function
