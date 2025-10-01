Attribute VB_Name = "XLSMB_SFTP_Part1_Process"

Sub RenameCSVFiles_ParsedVersion()
    Dim ws As Worksheet
    Dim csvPath As Variant
    Dim refData As Variant
    Dim fileDialog As fileDialog
    Dim selectedFiles As FileDialogSelectedItems
    Dim originalFileName As String
    Dim sftpName As String
    Dim GroupID As String
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
    Dim wb As Workbook
    Dim genderFixed As Boolean
    
    ' Summary tracking variables
    Dim zipFormattedList As String, apexAppliedList As String
    Dim macroExecutedList As String, renamedList As String
    Dim originalsBackedUpList As String
    Dim genderFixedList As String
    
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

    ' File System Object for renaming
    Set fso = CreateObject("Scripting.FileSystemObject")
    processedCount = 0
    ErrorCount = 0
    
    ' Initialize summary strings
    zipFormattedList = "ZIP Formatting Applied To:" & vbCrLf
    apexAppliedList = "APEX Logic Applied To:" & vbCrLf
    macroExecutedList = "Specific Macros Executed For:" & vbCrLf
    renamedList = "Successfully Renamed:" & vbCrLf
    originalsBackedUpList = "Original Files Backed Up:" & vbCrLf
    genderFixedList = "Gender Field Fixed In:" & vbCrLf

    ' Loop through selected files
    For Each csvPath In selectedFiles
        originalFileName = fso.GetFileName(csvPath)
        matchFound = False
        
        ' Find matching entry in worksheet data (skip header row)
        For i = 2 To UBound(refData, 1)
            ' Try pattern matching against Initial Filename Format (Column A)
            If IsFileMatchPattern_ProcessRename_Fixed(originalFileName, CStr(refData(i, 1)), CStr(refData(i, 6))) Then
                ' Extract data from parsed worksheet - CORRECTED COLUMN REFERENCES FOR 14 COLUMNS
                sftpName = CStr(refData(i, 10))    ' Adjusted GroupName (Column J)
                GroupID = CStr(refData(i, 11))     ' Adjusted Group ID (Column K)
                newFileName = CStr(refData(i, 13)) ' Adjusted Filename Final Save Format (Column M)
                
                ' Extract date from original filename and convert properly
                originalDateFormat = CStr(refData(i, 6)) ' Initial Date Format (Parsed) from Column F
                
                extractedDate = ExtractAndConvertDate_FIXED_FINAL(originalFileName, originalDateFormat)
                
                If extractedDate <> "" Then
                    ' Replace mmddyyyy in Final Save Format with actual date
                    newFileName = Replace(newFileName, "mmddyyyy", extractedDate)
                End If
                
                ' === BACKUP ORIGINAL FILE TO DATE-BASED FOLDER ===
                Dim saveFolder As String
                Dim originalsFolder As String
                Dim originalBackupPath As String
                Dim updatedDateFormat As String
                Dim folderDateFormat As String
                Dim dateBasedFolder As String
                
                saveFolder = CStr(refData(i, 14)) ' Column N - Final Save Folder
                updatedDateFormat = CStr(refData(i, 12)) ' Column L - Updated file date format
                
                ' Replace path placeholders for cross-user compatibility
                saveFolder = ResolveFolderPath(saveFolder)
                saveFolder = Replace(saveFolder, "[Adjusted GroupName]", sftpName)
                saveFolder = Replace(saveFolder, "[Adjusted groupID]", GroupID)
                saveFolder = Replace(saveFolder, "[GroupName]", sftpName)
                saveFolder = Replace(saveFolder, "[groupID]", GroupID)
                
                ' Create date-based folder structure matching Part2_MoveToFolder logic
                If saveFolder <> "" And saveFolder <> "**NOT INCLUDED**" And InStr(saveFolder, "[") = 0 And extractedDate <> "" Then
                    ' Format the date for the folder name (e.g., "09Sep25")
                    folderDateFormat = FormatDateForFolder_Part1(extractedDate, updatedDateFormat)
                    
                    If folderDateFormat <> "" Then
                        ' Create the full path: [SaveFolder]/[DateFolder]/ORIGINALS/
                        dateBasedFolder = fso.BuildPath(saveFolder, folderDateFormat)
                        originalsFolder = fso.BuildPath(dateBasedFolder, "ORIGINALS")
                        
                        ' Create folders if they don't exist
                        If Not fso.FolderExists(saveFolder) Then
                            CreateFolderPath fso, saveFolder
                        End If
                        If Not fso.FolderExists(dateBasedFolder) Then
                            fso.CreateFolder dateBasedFolder
                        End If
                        If Not fso.FolderExists(originalsFolder) Then
                            fso.CreateFolder originalsFolder
                        End If
                        
                        ' Copy original file to ORIGINALS folder
                        originalBackupPath = fso.BuildPath(originalsFolder, originalFileName)
                        On Error Resume Next
                        fso.CopyFile csvPath, originalBackupPath, True ' True = overwrite if exists
                        If Err.Number = 0 Then
                            originalsBackedUpList = originalsBackedUpList & originalFileName & " -> " & folderDateFormat & "\ORIGINALS\" & vbCrLf
                        End If
                        On Error GoTo 0
                    End If
                End If
                
                ' STEP 1: OPEN FILE AND APPLY FORMATTING
                On Error GoTo FileOpenError
                
                ' Handle different file types
                If LCase(fso.GetExtensionName(csvPath)) = "csv" Then
                    Application.DisplayAlerts = False
                    Workbooks.OpenText FileName:=csvPath, _
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
                
                ' Apply ZIP formatting with improved method
                zipFormatted = ApplyZipFormatting_ProcessRename_Fixed(ws)
                If zipFormatted Then zipFormattedList = zipFormattedList & originalFileName & vbCrLf
                
                ' Apply UNIVERSAL GENDER FIX
                genderFixed = ApplyUniversalGenderFix(ws)
                If genderFixed Then genderFixedList = genderFixedList & originalFileName & vbCrLf
                
                ' Apply APEX logic if applicable
                apexApplied = ApplyApexLogic_ProcessRename_Fixed(ws, wb.Name)
                If apexApplied Then apexAppliedList = apexAppliedList & originalFileName & vbCrLf
                
                ' STEP 2: EXECUTE SPECIFIC MACRO BASED ON SFTPNAME
                specificMacroExecuted = ExecuteSpecificMacro_ProcessRename_Fixed(ws, sftpName, CStr(refData(i, 8))) ' Column H - Specific Macro
                If specificMacroExecuted Then macroExecutedList = macroExecutedList & originalFileName & " (" & sftpName & " specific macro)" & vbCrLf
                
                ' STEP 3: SAVE THE FORMATTED FILE AS PROPER CSV
                On Error Resume Next
                Application.DisplayAlerts = False
                
                ' Create the target file name with date-based subfolder
                currentFolder = fso.GetParentFolderName(csvPath)
                
                ' Create today's date in mmddyyyy format
                Dim todayDate As String
                todayDate = Format(Date, "mmddyyyy")
                
                ' Create subfolder name: "mmddyyyy SFTP files"
                Dim subfolderName As String
                subfolderName = todayDate & " SFTP files"
                
                ' Create full path to subfolder
                Dim subfolderPath As String
                subfolderPath = fso.BuildPath(currentFolder, subfolderName)
                
                ' Create subfolder if it doesn't exist
                If Not fso.FolderExists(subfolderPath) Then
                    fso.CreateFolder (subfolderPath)
                End If
                
                ' Build target path in the subfolder
                targetPath = fso.BuildPath(subfolderPath, newFileName)
                
                ' Save as CSV, not XLSX
                wb.SaveAs FileName:=targetPath, FileFormat:=xlCSV
                
                Application.DisplayAlerts = True
                On Error GoTo 0
                wb.Close
                
                ' Skip the rename section since we saved directly with the final name
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

    ' Show comprehensive summary
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

' UPDATED: Function to read from internal worksheet instead of external CSV
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
                ' Apply text format and convert existing values to text with leading zeros
                ws.Columns(j).NumberFormat = "@"  ' Text format
                
                ' Convert existing values to text with leading zeros
                Dim lastRow As Long
                lastRow = ws.Cells(ws.Rows.count, j).End(xlUp).Row
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
    
    ' Find last row and column
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    
    ' Find Gender column (check for variations)
    genderCol = 0
    For j = 1 To lastCol
        Dim headerText As String
        headerText = UCase(Trim(ws.Cells(1, j).Value))
        
        ' Check various possible header names
        If headerText = "GENDER" Or _
           headerText = "SEX" Or _
           headerText = "MEMBER GENDER" Or _
           headerText = "MEMBERGENDER" Or _
           InStr(headerText, "GENDER") > 0 Then
            genderCol = j
            Exit For
        End If
    Next j
    
    ' If Gender column found, fill blanks with "M"
    If genderCol > 0 Then
        fixedCount = 0
        For i = 2 To lastRow ' Start from row 2 to skip header
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

Function ExtractAndConvertDate_FIXED_FINAL(FileName As String, dateFormat As String) As String
    On Error GoTo ErrorHandler
    
    Dim regex As Object
    Dim matches As Object
    Dim extractedDate As String
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True
    regex.Global = True  ' Get ALL matches, not just first
    
    ' Set pattern based on expected date format
    Select Case LCase(Trim(dateFormat))
        Case "mmddyy"
            regex.Pattern = "\d{6}"
        Case "mmddyyyy"
            regex.Pattern = "\d{8}"
        Case "yyyymmdd"
            regex.Pattern = "\d{8}"
        Case Else
            ExtractAndConvertDate_FIXED_FINAL = ""
            Exit Function
    End Select
    
    Set matches = regex.Execute(FileName)
    
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

Function IsFileMatchPattern_ProcessRename_Fixed(FileName As String, csvPattern As String, dateFormat As String) As Boolean
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
        .Pattern = regexPattern
        .IgnoreCase = True
        .Global = False
    End With
    
    IsFileMatchPattern_ProcessRename_Fixed = regex.Test(FileName)
    Exit Function
    
ErrorHandler:
    IsFileMatchPattern_ProcessRename_Fixed = False
End Function

Function ExecuteSpecificMacro_ProcessRename_Fixed(ws As Worksheet, sftpName As String, specificMacroPath As String) As Boolean
    On Error GoTo MacroError
    
    ' DEBUG: Show what we received
    Debug.Print "=== ExecuteSpecificMacro_ProcessRename_Fixed DEBUG ==="
    Debug.Print "sftpName: " & sftpName
    Debug.Print "specificMacroPath: " & specificMacroPath
    
    If specificMacroPath <> "" And Len(Trim(specificMacroPath)) > 0 Then
        ' Extract module name from file path
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim moduleName As String
        moduleName = fso.GetBaseName(specificMacroPath)  ' Gets "Specifics_PF1" from "Specifics_PF1.bas"
        
        Debug.Print "Extracted module name: " & moduleName
        
        ' Try to run the ProcessCSVFiles procedure from the specified module
        On Error Resume Next
        Debug.Print "Attempting to call: " & moduleName & ".ProcessCSVFiles"
        Application.Run moduleName & ".ProcessCSVFiles"
        
        If Err.Number = 0 Then
            Debug.Print "SUCCESS: Called " & moduleName & ".ProcessCSVFiles"
            ExecuteSpecificMacro_ProcessRename_Fixed = True
        Else
            Debug.Print "FAILED: " & moduleName & ".ProcessCSVFiles - Error: " & Err.Description
            ' If ProcessCSVFiles doesn't exist, try other common procedure names
            Err.Clear
            Debug.Print "Attempting to call: " & moduleName & ".Main"
            Application.Run moduleName & ".Main"
            If Err.Number = 0 Then
                Debug.Print "SUCCESS: Called " & moduleName & ".Main"
                ExecuteSpecificMacro_ProcessRename_Fixed = True
            Else
                Debug.Print "FAILED: " & moduleName & ".Main - Error: " & Err.Description
                ' Try procedure with same name as module
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
        ' Handle cases where no specific macro path is provided (existing logic)
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

Function ApplyApexLogic_ProcessRename_Fixed(ws As Worksheet, FileName As String) As Boolean
    If InStr(1, UCase(FileName), "APEX") = 0 Then
        ApplyApexLogic_ProcessRename_Fixed = False
        Exit Function
    End If
    
    On Error GoTo ApexError
    Dim rngP As Range, cell As Range
    Dim dict As Object
    Dim i As Long, lastRow As Long
    
    If ws.Cells(ws.Rows.count, "P").End(xlUp).Row < 2 Then
        ApplyApexLogic_ProcessRename_Fixed = False
        Exit Function
    End If
    
    Set rngP = ws.Range("P2:P" & ws.Cells(ws.Rows.count, "P").End(xlUp).Row)
    Set dict = CreateObject("Scripting.Dictionary")
    
    For Each cell In rngP
        If Not dict.Exists(cell.Value) Then
            dict.Add cell.Value, 1
        Else
            dict(cell.Value) = dict(cell.Value) + 1
        End If
    Next cell

    For i = ws.Cells(ws.Rows.count, "P").End(xlUp).Row To 2 Step -1
        If dict.Exists(ws.Cells(i, "P").Value) And dict(ws.Cells(i, "P").Value) > 1 _
            And ws.Cells(i, "N").Value <> "" Then
            ws.Rows(i).Delete
        End If
    Next i

    Set dict = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Rows.count, "P").End(xlUp).Row
    For Each cell In ws.Range("P2:P" & lastRow)
        If Not dict.Exists(cell.Value) Then
            dict.Add cell.Value, cell.Row
        ElseIf ws.Cells(cell.Row, "M").Value < ws.Cells(dict(cell.Value), "M").Value Then
            ws.Rows(cell.Row).Delete
        Else
            ws.Rows(dict(cell.Value)).Delete
            dict(cell.Value) = cell.Row
        End If
    Next cell
    
    ApplyApexLogic_ProcessRename_Fixed = True
    Exit Function
    
ApexError:
    ApplyApexLogic_ProcessRename_Fixed = False
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
    regex.Pattern = "[^A-Za-z0-9]"
    regex.Global = True

    lastRow = ws.Cells(ws.Rows.count, locationCol).End(xlUp).Row

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
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row

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

    lastRow = ws.Cells(ws.Rows.count, genderCol).End(xlUp).Row

    For i = 2 To lastRow
        If Trim(ws.Cells(i, genderCol).Value) = "" Then
            ws.Cells(i, genderCol).Value = "M"
        End If
    Next i
End Sub

' Create folder path recursively
Sub CreateFolderPath(fso As Object, fullPath As String)
    Dim pathParts() As String
    Dim currentPath As String
    Dim i As Integer
    
    ' Split the path into parts
    pathParts = Split(fullPath, "\")
    
    ' Start with drive letter or network path
    currentPath = pathParts(0)
    
    ' Build path step by step
    For i = 1 To UBound(pathParts)
        currentPath = currentPath & "\" & pathParts(i)
        If Not fso.FolderExists(currentPath) Then
            On Error Resume Next
            fso.CreateFolder currentPath
            On Error GoTo 0
        End If
    Next i
End Sub

' Resolve folder path placeholders
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
    
    oneDrivePath = Environ("OneDriveCommercial")
    If oneDrivePath = "" Then
        oneDrivePath = Environ("OneDrive")
    End If
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

' Format Date for Folder (Part1 version) - matches Part2_MoveToFolder logic
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
    
    ' Determine format of the date string
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
            yearPart = Right(dateString, 2) ' Get last 2 digits of year
        Case "yyyymmdd"
            yearPart = Right(Left(dateString, 4), 2) ' Get last 2 digits of year
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
    
    ' Format: "09Sep25"
    FormatDateForFolder_Part1 = Format(monthPart, "00") & monthNames(monthIndex) & yearPart
    Exit Function
    
ErrorHandler:
    FormatDateForFolder_Part1 = ""
End Function

