Attribute VB_Name = "FileValidationMain"

Option Explicit

' Main entry point for file validation system
Public Sub StartFileValidation()
    Const sPROC_NAME As String = "StartFileValidation"
    
    On Error GoTo ErrorHandler
    
    ' Initialize application settings
    Call InitializeApplication
    
    ' Show file selection dialog for MULTIPLE files
    Dim selectedFiles As Collection
    Set selectedFiles = SelectMultipleValidationFiles()
    
    If selectedFiles Is Nothing Or selectedFiles.count = 0 Then
        MsgBox "No files selected. Validation cancelled.", vbInformation, "File Validation"
        GoTo Cleanup
    End If
    
    ' Create Excel report workbook
    Dim wbReport As Workbook
    Set wbReport = CreateValidationReport(selectedFiles)
    
    ' Save the report
    Dim sReportPath As String
    sReportPath = Environ("USERPROFILE") & "\Downloads\ValidationReport_" & _
                  Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
    wbReport.SaveAs sReportPath
    
    MsgBox "Validation complete!" & vbCrLf & vbCrLf & _
           "Files processed: " & selectedFiles.count & vbCrLf & _
           "Report saved to: " & sReportPath, vbInformation, "Validation Complete"
    
Cleanup:
    Call RestoreApplication
    Exit Sub
    
ErrorHandler:
    Call RestoreApplication
    Call ErrorHandler_Central(sPROC_NAME, Err.Number, Err.Description)
End Sub

' Main entry point for manual file validation with user-selected filetype
Public Sub StartManualFileTypeValidation()
    Const sPROC_NAME As String = "StartManualFileTypeValidation"
    
    On Error GoTo ErrorHandler
    
    ' Initialize application settings
    Call InitializeApplication
    
    ' STEP 1: Get file type FIRST
    ' Get list of available file types from Filetype Mapping sheet
    Dim colFileTypes As Collection
    Set colFileTypes = GetAvailableFileTypes()
    
    If colFileTypes.count = 0 Then
        MsgBox "No file types found in Filetype Mapping sheet.", vbExclamation, "File Validation"
        GoTo Cleanup
    End If
    
    ' Show form to select filetype
    Dim sSelectedFileType As String
    sSelectedFileType = ShowFileTypeSelectionDialog(colFileTypes)
    
    If sSelectedFileType = "" Then
        MsgBox "No file type selected. Validation cancelled.", vbInformation, "File Validation"
        GoTo Cleanup
    End If
    
    ' STEP 2: Now select MULTIPLE files for that file type
    MsgBox "You selected: " & sSelectedFileType & vbCrLf & vbCrLf & _
           "Now select one or more CSV files to validate with this file type.", _
           vbInformation, "File Type Selected"
           
    Dim selectedFiles As Collection
    Set selectedFiles = SelectMultipleFilesForManualValidation()
    
    If selectedFiles Is Nothing Or selectedFiles.count = 0 Then
        MsgBox "No files selected. Validation cancelled.", vbInformation, "File Validation"
        GoTo Cleanup
    End If
    
    ' Store the selected file type with each file path for processing
    Dim colFilesWithType As New Collection
    Dim sFilePath As Variant
    For Each sFilePath In selectedFiles
        colFilesWithType.Add sFilePath & "|" & sSelectedFileType
    Next sFilePath
    
    Application.StatusBar = "Validating " & selectedFiles.count & " file(s) with FileType: " & sSelectedFileType
    
    ' Create Excel report workbook (same as automatic validation)
    Dim wbReport As Workbook
    Set wbReport = CreateValidationReportWithManualType(colFilesWithType)
    
    ' Save the report
    Dim sReportPath As String
    sReportPath = Environ("USERPROFILE") & "\Downloads\ValidationReport_Manual_" & _
                  Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
    wbReport.SaveAs sReportPath
    
    MsgBox "Validation complete!" & vbCrLf & vbCrLf & _
           "Files processed: " & selectedFiles.count & vbCrLf & _
           "File Type: " & sSelectedFileType & vbCrLf & _
           "Report saved to: " & sReportPath, vbInformation, "Validation Complete"
    
Cleanup:
    Call RestoreApplication
    Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    Call RestoreApplication
    Application.StatusBar = False
    Call ErrorHandler_Central(sPROC_NAME, Err.Number, Err.Description)
End Sub

' Add this new function for selecting multiple files for manual validation
Private Function SelectMultipleFilesForManualValidation() As Collection
    Dim colFiles As New Collection
    Dim fd As fileDialog
    Dim vrtSelectedItem As Variant
    
    Set fd = Application.fileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "Select CSV Files to Validate"
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        .Filters.Add "All Files", "*.*"
        .FilterIndex = 1
        .AllowMultiSelect = True  ' Enable multiple file selection
        .ButtonName = "Validate"
        
        If .Show = -1 Then
            For Each vrtSelectedItem In .SelectedItems
                colFiles.Add CStr(vrtSelectedItem)
            Next vrtSelectedItem
        End If
    End With
    
    Set SelectMultipleFilesForManualValidation = colFiles
End Function

Private Function ValidateSelectedFile(sFilePath As String) As ValidationResult
    Const sPROC_NAME As String = "ValidateSelectedFile"
    
    Dim oResult As ValidationResult
    Set oResult = New ValidationResult
    
    On Error GoTo ErrorHandler
    
    ' Show progress form
    Load frmProgress
    frmProgress.Show vbModeless
    
    ' Stage 1: Extract filename and match pattern
    Call frmProgress.UpdateProgress("Analyzing filename pattern...", 5)
    
    Dim sFileName As String
    sFileName = GetFileNameFromPath(sFilePath)
           
    Dim oFileInfo As FileInfo
    oFileInfo = MatchFilenamePattern(sFileName)
    
    If Not oFileInfo.IsValid Then
        oResult.AddError 0, "Filename", "No matching pattern found for filename: " & sFileName
        GoTo Cleanup
    End If
    ' Populate ValidationResult with file information
    oResult.FileName = sFileName
    oResult.FilePath = sFilePath
    oResult.FileType = oFileInfo.FileType
    oResult.GroupID = oFileInfo.GroupID
    oResult.ProcessedDate = Now
    
    ' Get Group Name from lookup sheet
    oResult.GroupName = GetGroupName(oFileInfo.GroupID)
    
    ' Record validation checks performed
    oResult.AddValidationCheck "Filename Pattern", "Matched: " & sFileName
    oResult.AddValidationCheck "FileType Identified", oFileInfo.FileType
    oResult.AddValidationCheck "GroupID Extracted", oFileInfo.GroupID
    
    ' Stage 2: Get column mapping for FileType
    Call frmProgress.UpdateProgress("Loading column mappings...", 15)
    
    Dim oMapping As ColumnMapping
    oMapping = GetColumnMapping(oFileInfo.FileType)
    
    If oMapping.FileType = "" Then
        oResult.AddError 0, "FileType", "No column mapping found for FileType: " & oFileInfo.FileType
        GoTo Cleanup
    End If
    
    ' Stage 3: Load validation rules
    Call frmProgress.UpdateProgress("Loading validation rules...", 25)
    
    Dim colRules As Collection
    Set colRules = LoadValidationRules()
    
    ' Stage 4: Read and parse CSV file
    Call frmProgress.UpdateProgress("Reading CSV file...", 35)
    
    Dim vCSVData As Variant
    vCSVData = ReadCSVToArray(sFilePath)
    
    If IsEmpty(vCSVData) Then
        oResult.AddError 0, "File", "Failed to read CSV file or file is empty"
        GoTo Cleanup
    End If
    
    ' Stage 5: Apply validation rules
    Call frmProgress.UpdateProgress("Validating data...", 50)
    
    Call ValidateCSVDataIntoResult(vCSVData, oMapping, colRules, oFileInfo, oResult)
    
    ' Stage 6: Complete
    Call frmProgress.UpdateProgress("Validation complete!", 100)
    
Cleanup:
    On Error Resume Next
    Unload frmProgress
    On Error GoTo 0
    Set ValidateSelectedFile = oResult
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    Unload frmProgress
    On Error GoTo 0
    Call ErrorHandler_Central(sPROC_NAME, Err.Number, Err.Description, sFilePath)
    If oResult Is Nothing Then Set oResult = New ValidationResult
    oResult.AddError 0, "System", "Validation failed due to system error: " & Err.Description
    GoTo Cleanup
End Function

' Validate file with manually specified filetype
Private Function ValidateFileWithManualType(sFilePath As String, sFileType As String) As ValidationResult
    Const sPROC_NAME As String = "ValidateFileWithManualType"
    
    Dim oResult As ValidationResult
    Set oResult = New ValidationResult
    Dim sFileName As String
    
    On Error GoTo ErrorHandler
    
    ' Show progress form
    Load frmProgress
    frmProgress.Show vbModeless
    
    ' Extract filename from path
    Call frmProgress.UpdateProgress("Processing file...", 5)
    sFileName = GetFileNameFromPath(sFilePath)
    
    ' Populate ValidationResult with file information
    oResult.FileName = sFileName
    oResult.FilePath = sFilePath
    oResult.FileType = sFileType
    oResult.ProcessedDate = Now
    oResult.GroupID = "MANUAL"  ' Use MANUAL as placeholder for manual validations
    oResult.GroupName = "Manual Validation - " & sFileType
    
    ' Record validation checks performed
    oResult.AddValidationCheck "Filename", sFileName
    oResult.AddValidationCheck "FileType (Manual)", sFileType
    
    ' Stage 1: Get column mapping for selected FileType
    Call frmProgress.UpdateProgress("Loading column mappings...", 15)
    
    Dim oMapping As ColumnMapping
    oMapping = GetColumnMapping(sFileType)
    
    If oMapping.FileType = "" Then
        oResult.AddError 0, "FileType", "No column mapping found for FileType: " & sFileType
        GoTo Cleanup
    End If
    
    ' Stage 2: Load validation rules
    Call frmProgress.UpdateProgress("Loading validation rules...", 25)
    
    Dim colRules As Collection
    Set colRules = LoadValidationRules()
    
    ' Stage 3: Read and parse CSV file
    Call frmProgress.UpdateProgress("Reading CSV file...", 35)
    
    Dim vCSVData As Variant
    vCSVData = ReadCSVToArray(sFilePath)
    
    If IsEmpty(vCSVData) Then
        oResult.AddError 0, "File", "Failed to read CSV file or file is empty"
        GoTo Cleanup
    End If
    
    ' Stage 4: Apply validation rules (without FileInfo since no pattern matching)
    Call frmProgress.UpdateProgress("Validating data...", 50)
    
    ' Create a dummy FileInfo for compatibility
    Dim oFileInfo As FileInfo
    oFileInfo.FileType = sFileType
    oFileInfo.GroupID = "MANUAL"  ' Use MANUAL for manual validations
    oFileInfo.IsValid = True
    
    Call ValidateCSVDataIntoResult(vCSVData, oMapping, colRules, oFileInfo, oResult)
    
    Call frmProgress.UpdateProgress("Validation complete!", 100)
    
Cleanup:
    On Error Resume Next
    Unload frmProgress
    On Error GoTo 0
    Set ValidateFileWithManualType = oResult
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    Unload frmProgress
    On Error GoTo 0
    Call ErrorHandler_Central(sPROC_NAME, Err.Number, Err.Description)
    oResult.AddError 0, "System", "Validation failed: " & Err.Description
    GoTo Cleanup
End Function

Private Sub ValidateCSVDataIntoResult(vData As Variant, oMapping As ColumnMapping, colRules As Collection, oFileInfo As FileInfo, oResult As ValidationResult)
    Dim lRow As Long
    Dim lTotalRows As Long
    Dim colMemberIDs As Collection
    Set colMemberIDs = New Collection
    
    lTotalRows = UBound(vData, 1)
    oResult.TotalRecords = lTotalRows - 1 ' Subtract header row
    
    ' Validate each data row (skip header)
    For lRow = 2 To lTotalRows
        ' Update progress occasionally
        If lRow Mod 1000 = 0 Then
            Dim lPercent As Long
            lPercent = 50 + ((lRow / lTotalRows) * 30) ' Progress from 50% to 80%
            Call frmProgress.UpdateProgress("Validating record " & (lRow) & " of " & (lTotalRows - 1), lPercent)
        End If
        
        ' Validate individual fields
        Call ValidateRowFields(vData, lRow, oMapping, colRules, oResult)
        
        ' Check MemberID uniqueness (only for active records)
        If oMapping.MemberID > 0 And oMapping.MemberID <= UBound(vData, 2) Then
            Dim sMemberID As String
            Dim sEffectiveEndDate As String
            Dim bCurrentIsActive As Boolean
            
            ' Get MemberID for current row
            sMemberID = CStr(vData(lRow, oMapping.MemberID))
            
            ' Get EffectiveEndDate if column exists
            If oMapping.EffectiveEndDate > 0 And oMapping.EffectiveEndDate <= UBound(vData, 2) Then
                sEffectiveEndDate = CStr(vData(lRow, oMapping.EffectiveEndDate))
            Else
                sEffectiveEndDate = "" ' Default to blank if column not mapped
            End If
            
            ' Check if current record is active
            bCurrentIsActive = ActiveChecker(sEffectiveEndDate)
            
            If sMemberID <> "" Then
                On Error Resume Next
                ' Try to add to collection with active status
                colMemberIDs.Add sMemberID & "|" & bCurrentIsActive, sMemberID
                
                If Err.Number <> 0 Then
                    ' Duplicate MemberID found - but only error if both are active
                    Err.Clear
                    On Error GoTo 0
                    
                    ' Get the previously stored value
                    Dim sStoredValue As String
                    Dim bPreviousIsActive As Boolean
                    
                    sStoredValue = colMemberIDs.Item(sMemberID)
                    
                    ' Extract the active status from stored value (format: "MemberID|True" or "MemberID|False")
                    Dim vParts As Variant
                    vParts = Split(sStoredValue, "|")
                    
                    If UBound(vParts) >= 1 Then
                        bPreviousIsActive = CBool(vParts(1))
                    Else
                        bPreviousIsActive = True ' Default to active if can't parse
                    End If
                    
                    ' CRITICAL: Only flag as error if BOTH occurrences are active
                    If bCurrentIsActive And bPreviousIsActive Then
                        oResult.AddError lRow, "MemberID", "Duplicate active MemberID found: " & sMemberID & _
                                        " (both records have blank or future EffectiveEndDate)"
                    End If
                    ' If one is inactive, NO ERROR is added - the duplicate is allowed
                End If
                On Error GoTo 0
            End If
        End If
                
        ' Validate GroupID matches expected (skip for manual validations)
        If oFileInfo.GroupID <> "MANUAL" Then
            If oMapping.GroupID > 0 And oMapping.GroupID <= UBound(vData, 2) Then
                If lRow = 2 Then oResult.AddValidationCheck "Group ID Match", "Verifying against expected: " & oFileInfo.GroupID
                Dim sGroupID As String
                sGroupID = CStr(vData(lRow, oMapping.GroupID))
                
                If sGroupID <> oFileInfo.GroupID Then
                    oResult.AddError lRow, "GroupID", "GroupID mismatch. Expected: " & oFileInfo.GroupID & ", Found: " & sGroupID
                End If
            End If
        End If
    Next lRow
    
    oResult.ValidationComplete = True
End Sub

Private Sub InitializeApplication()
    ' Store current settings
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
End Sub

Private Sub RestoreApplication()
    ' Restore Excel settings
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub

Private Sub ValidateRowFields(vData As Variant, lRow As Long, oMapping As ColumnMapping, colRules As Collection, oResult As ValidationResult)
    ' Validate First Name
    If oMapping.FirstName > 0 And oMapping.FirstName <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.FirstName), "FirstName", lRow, colRules, oResult)
    End If
    
    ' Validate Last Name
    If oMapping.LastName > 0 And oMapping.LastName <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.LastName), "LastName", lRow, colRules, oResult)
    End If
    
    ' Validate DOB
    If oMapping.DOB > 0 And oMapping.DOB <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.DOB), "DOB", lRow, colRules, oResult)
    End If
    
    ' Validate Gender
    If oMapping.Gender > 0 And oMapping.Gender <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.Gender), "Gender", lRow, colRules, oResult)
    End If
    
    ' Validate Zip Code
    If oMapping.ZipCode > 0 And oMapping.ZipCode <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.ZipCode), "ZipCode", lRow, colRules, oResult)
    End If
    
    ' Validate Address1
    If oMapping.Address1 > 0 And oMapping.Address1 <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.Address1), "Address1", lRow, colRules, oResult)
    End If
    
    ' Validate City
    If oMapping.City > 0 And oMapping.City <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.City), "City", lRow, colRules, oResult)
    End If
    
    ' Validate State
    If oMapping.State > 0 And oMapping.State <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.State), "State", lRow, colRules, oResult)
    End If
    
    ' Validate Effective Date
    If oMapping.EffectiveDate > 0 And oMapping.EffectiveDate <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.EffectiveDate), "EffectiveDate", lRow, colRules, oResult)
    End If
    
    ' Validate Service Offering
    If oMapping.ServiceOffering > 0 And oMapping.ServiceOffering <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.ServiceOffering), "ServiceOffering", lRow, colRules, oResult)
    End If
End Sub

Public Function GetValidationRule(colRules As Collection, sFieldType As String) As ValidationRule
    Dim oRule As ValidationRule
    Dim emptyRule As ValidationRule
    Dim sRuleData As String
    Dim vParts As Variant
    
    ' Initialize empty rule
    emptyRule.FieldType = ""
    emptyRule.Required = False
    emptyRule.MaxLength = 0
    emptyRule.MinLength = 0
    emptyRule.FormatPattern = ""
    emptyRule.CustomFunction = ""
    
    On Error GoTo NotFound
    
    ' Try to get the rule data string from collection using the field type as key
    sRuleData = colRules.Item(sFieldType)
    
    ' Parse the delimited string back into ValidationRule Type
    vParts = Split(sRuleData, "|")
    
    If UBound(vParts) >= 5 Then
        oRule.FieldType = vParts(0)
        oRule.Required = (UCase(vParts(1)) = "TRUE")
        oRule.MaxLength = Val(vParts(2))
        oRule.MinLength = Val(vParts(3))
        oRule.FormatPattern = vParts(4)
        oRule.CustomFunction = vParts(5)
        
        GetValidationRule = oRule
        Exit Function
    End If
    
NotFound:
    ' Return empty rule if not found or error
    GetValidationRule = emptyRule
End Function

Private Sub ValidateField(vFieldValue As Variant, sFieldType As String, lRowNumber As Long, colRules As Collection, oResult As ValidationResult)
    ' Record that we're checking this field (only once per field type)
    Static checkedFields As Collection
    If checkedFields Is Nothing Then Set checkedFields = New Collection
    
    On Error Resume Next
    checkedFields.Add sFieldType, sFieldType
    If Err.Number = 0 Then  ' First time checking this field type
        oResult.AddValidationCheck sFieldType & " Field", "Validated across all records"
    End If
    On Error GoTo 0
    
    ' Get the validation rule for max/min length and format patterns
    Dim oRule As ValidationRule
    oRule = GetValidationRule(colRules, sFieldType)
    
    Dim sValue As String
    sValue = CStr(vFieldValue)
    
    ' AUTOMATIC BLANK FIELD CHECK FOR ALL CORE FIELDS
    Dim isRequiredField As Boolean
    
    Select Case sFieldType
        Case "FirstName", "LastName", "DOB", "Gender", _
             "ZipCode", "Address1", "City", "State", _
             "EffectiveDate", "ServiceOffering", "MemberID", "GroupID"
            isRequiredField = True
        Case Else
            isRequiredField = False
    End Select
    
    ' Perform the blank check for all required fields
    If isRequiredField Then
        If sValue = "" Or IsNull(vFieldValue) Or Trim(sValue) = "" Then
            oResult.AddError lRowNumber, sFieldType, "Required field is blank"
            ' Log this validation check
            oResult.AddValidationCheck "Blank Check - " & sFieldType, "Row " & lRowNumber & ": Found blank"
            Exit Sub  ' No need to check other validations if field is blank
        End If
    End If
    
    ' OTHER VALIDATION CHECKS (Max Length, Format, etc.)
    ' Skip other validations if field is empty (and not required)
    If sValue = "" Then Exit Sub
    
    ' Max character check (still uses Column Checks sheet)
    If oRule.MaxLength > 0 And Len(sValue) > oRule.MaxLength Then
        oResult.AddError lRowNumber, sFieldType, "Exceeds maximum length of " & oRule.MaxLength & " characters (found " & Len(sValue) & ")"
        oResult.AddValidationCheck "Max Length Check - " & sFieldType, "Row " & lRowNumber & ": Exceeds " & oRule.MaxLength & " chars"
    End If
    
    ' Min character check (still uses Column Checks sheet)
    If oRule.MinLength > 0 And Len(sValue) < oRule.MinLength Then
        oResult.AddError lRowNumber, sFieldType, "Below minimum length of " & oRule.MinLength & " characters (found " & Len(sValue) & ")"
        oResult.AddValidationCheck "Min Length Check - " & sFieldType, "Row " & lRowNumber & ": Below " & oRule.MinLength & " chars"
    End If
    
    ' Format validation (still uses Column Checks sheet patterns)
    If oRule.FormatPattern <> "" Then
        If Not ValidateFieldFormat(sValue, sFieldType, oRule.FormatPattern) Then
            oResult.AddError lRowNumber, sFieldType, "Invalid format for " & sFieldType
            oResult.AddValidationCheck "Format Check - " & sFieldType, "Row " & lRowNumber & ": Invalid format"
        End If
    End If
End Sub

Private Function ValidateFieldFormat(sValue As String, sFieldType As String, sPattern As String) As Boolean
    Select Case UCase(sFieldType)
        Case "DOB", "EFFECTIVEDATE"
            ValidateFieldFormat = ValidateDateFormat(sValue)
        Case "GENDER"
            ValidateFieldFormat = ValidateGenderCode(sValue)
        Case "ZIPCODE"
            ValidateFieldFormat = ValidateZipCode(sValue)
        Case "FIRSTNAME", "LASTNAME", "CITY"
            ValidateFieldFormat = ValidateNameFormat(sValue)
        Case "STATE"
            ValidateFieldFormat = ValidateStateCode(sValue)
        Case Else
            ' Use regex pattern if provided
            If sPattern <> "" Then
                ValidateFieldFormat = ValidateWithRegex(sValue, sPattern)
            Else
                ValidateFieldFormat = True ' No specific validation
            End If
    End Select
End Function

Private Function ValidateDateFormat(sValue As String) As Boolean
    On Error Resume Next
    Dim dtTest As Date
    dtTest = CDate(sValue)
    ValidateDateFormat = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function ValidateGenderCode(sValue As String) As Boolean
    Dim vValidCodes As Variant
    vValidCodes = Array("M", "F", "MALE", "FEMALE", "1", "2", "0", "U", "UNKNOWN")
    
    Dim i As Long
    For i = 0 To UBound(vValidCodes)
        If UCase(Trim(sValue)) = UCase(CStr(vValidCodes(i))) Then
            ValidateGenderCode = True
            Exit Function
        End If
    Next i
    
    ValidateGenderCode = False
End Function

Private Function ValidateZipCode(sValue As String) As Boolean
    Dim oRegex As Object
    Set oRegex = CreateObject("VBScript.RegExp")
    
    ' US: 12345 or 12345-6789
    oRegex.Pattern = "^\d{5}(-\d{4})?$"
    ValidateZipCode = oRegex.Test(sValue)
End Function

Private Function ValidateNameFormat(sValue As String) As Boolean
    Dim oRegex As Object
    Set oRegex = CreateObject("VBScript.RegExp")
    
    ' Allow letters, spaces, hyphens, apostrophes, periods
    oRegex.Pattern = "^[a-zA-Z][a-zA-Z\s\-'\.]{1,49}$"
    oRegex.IgnoreCase = True
    
    ValidateNameFormat = oRegex.Test(Trim(sValue)) And Len(Trim(sValue)) >= 2
End Function

Private Function ValidateStateCode(sValue As String) As Boolean
    ' This could be expanded with a full list of state codes
    ValidateStateCode = (Len(Trim(sValue)) = 2)
End Function

Private Function ValidateWithRegex(sValue As String, sPattern As String) As Boolean
    Dim oRegex As Object
    Set oRegex = CreateObject("VBScript.RegExp")
    
    oRegex.Pattern = sPattern
    ValidateWithRegex = oRegex.Test(sValue)
End Function

Private Function SelectMultipleValidationFiles() As Collection
    Dim colFiles As New Collection
    Dim fd As fileDialog
    Dim vrtSelectedItem As Variant
    
    Set fd = Application.fileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "Select CSV Files for Validation"
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        .Filters.Add "All Files", "*.*"
        .FilterIndex = 1
        .AllowMultiSelect = True  ' Enable multiple file selection
        
        If .Show = -1 Then
            For Each vrtSelectedItem In .SelectedItems
                colFiles.Add CStr(vrtSelectedItem)
            Next vrtSelectedItem
        End If
    End With
    
    Set SelectMultipleValidationFiles = colFiles
End Function

Private Function CreateValidationReport(selectedFiles As Collection) As Workbook
    Dim wbReport As Workbook
    Dim wsSummary As Worksheet
    Dim wsDetail As Worksheet
    Dim oResult As ValidationResult
    Dim sFilePath As Variant
    Dim lRow As Long
    Dim i As Long
    Dim bAnyFailures As Boolean
    
    ' Initialize
    bAnyFailures = False
    
    ' Create new workbook for report
    Set wbReport = Workbooks.Add
    Set wsSummary = wbReport.Sheets(1)
    wsSummary.Name = "Summary"
    
    ' Set up summary sheet headers
    With wsSummary
        .Range("A1").Value = "VALIDATION SUMMARY REPORT"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        
        .Range("A3").Value = "Generated:"
        .Range("B3").Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
        
        .Range("A5").Value = "File Name"
        .Range("B5").Value = "Status"
        .Range("C5").Value = "Records"
        .Range("D5").Value = "Errors"
        .Range("E5").Value = "Warnings"
        .Range("F5").Value = "File Type"
        .Range("G5").Value = "Group ID"
        .Range("H5").Value = "Details"
        
        ' Format headers
        .Range("A5:H5").Font.Bold = True
        .Range("A5:H5").Interior.Color = RGB(200, 200, 200)
        .Range("A5:H5").Borders.LineStyle = xlContinuous
    End With
    
    ' Process each file
    lRow = 6
    i = 1
    For Each sFilePath In selectedFiles
        ' Show progress
        Application.StatusBar = "Validating file " & i & " of " & selectedFiles.count & "..."
        
        ' Validate the file
        Set oResult = ValidateSelectedFile(CStr(sFilePath))
        
        ' Track if any files failed
        If Not oResult.IsValid Then
            bAnyFailures = True
        End If
        
        ' Add to summary
        With wsSummary
            .Range("A" & lRow).Value = oResult.FileName
            
            ' Status with color coding
            .Range("B" & lRow).Value = IIf(oResult.IsValid, "PASSED", "FAILED")
            If oResult.IsValid Then
                .Range("B" & lRow).Interior.Color = RGB(200, 255, 200) ' Light green
            Else
                .Range("B" & lRow).Interior.Color = RGB(255, 200, 200) ' Light red
            End If
            
            .Range("C" & lRow).Value = oResult.TotalRecords
            .Range("D" & lRow).Value = oResult.ErrorCount
            .Range("E" & lRow).Value = oResult.WarningCount
            .Range("F" & lRow).Value = oResult.FileType
            .Range("G" & lRow).Value = oResult.GroupID
            
            Dim sUniqueSheetName As String
            sUniqueSheetName = GetUniqueSheetName(wbReport, oResult.GroupName)
            .Range("H" & lRow).Value = "See Sheet: " & sUniqueSheetName
        End With
        
        ' Create detailed sheet for this file
        Set wsDetail = wbReport.Sheets.Add(After:=wbReport.Sheets(wbReport.Sheets.count))
        wsDetail.Name = sUniqueSheetName
        
        ' COLOR THE DETAIL SHEET TAB
        If oResult.IsValid Then
            ' Green for passed files
            wsDetail.Tab.Color = RGB(0, 176, 80)  ' Green
        Else
            ' Red for failed files
            wsDetail.Tab.Color = RGB(192, 0, 0)   ' Red
        End If
        
        ' Fill detailed sheet
        Call FillDetailedSheet(wsDetail, oResult, CStr(sFilePath))
        
        lRow = lRow + 1
        i = i + 1
    Next sFilePath
    
    ' COLOR THE SUMMARY TAB BASED ON OVERALL RESULTS
    If bAnyFailures Then
        ' Red if any files failed
        wsSummary.Tab.Color = RGB(192, 0, 0)  ' Red
    Else
        ' Green if all files passed
        wsSummary.Tab.Color = RGB(0, 176, 80) ' Green
    End If
    
    ' Auto-fit columns in summary
    wsSummary.Columns("A:H").AutoFit
    wsSummary.Activate
    
    Application.StatusBar = False
    Call ReorganizeSheetsByStatus(wbReport)
    Set CreateValidationReport = wbReport
End Function

' Create validation report for manual file type selection
Private Function CreateValidationReportWithManualType(selectedFiles As Collection) As Workbook
    Dim wbReport As Workbook
    Dim wsSummary As Worksheet
    Dim wsDetail As Worksheet
    Dim oResult As ValidationResult
    Dim sFileData As Variant
    Dim vParts As Variant
    Dim sFilePath As String
    Dim sFileType As String
    Dim lRow As Long
    Dim i As Long
    Dim bAnyFailures As Boolean
    
    ' Initialize
    bAnyFailures = False
    
    ' Create new workbook for report
    Set wbReport = Workbooks.Add
    Set wsSummary = wbReport.Sheets(1)
    wsSummary.Name = "Summary"
    
    ' Set up summary sheet headers
    With wsSummary
        .Range("A1").Value = "VALIDATION SUMMARY REPORT (Manual FileType)"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        
        .Range("A3").Value = "Generated:"
        .Range("B3").Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
        
        .Range("A5").Value = "File Name"
        .Range("B5").Value = "Status"
        .Range("C5").Value = "Records"
        .Range("D5").Value = "Errors"
        .Range("E5").Value = "Warnings"
        .Range("F5").Value = "File Type"
        .Range("G5").Value = "Validation Mode"
        .Range("H5").Value = "Details"
        
        ' Format headers
        .Range("A5:H5").Font.Bold = True
        .Range("A5:H5").Interior.Color = RGB(200, 200, 200)
        .Range("A5:H5").Borders.LineStyle = xlContinuous
    End With
    
    ' Process each file
    lRow = 6
    i = 1
    For Each sFileData In selectedFiles
        ' Parse file path and file type
        vParts = Split(sFileData, "|")
        sFilePath = vParts(0)
        sFileType = vParts(1)
        
        ' Show progress
        Application.StatusBar = "Validating file with FileType: " & sFileType & "..."
        
        ' Validate the file with manual type
        Set oResult = ValidateFileWithManualType(sFilePath, sFileType)
        
        ' Track if any files failed
        If Not oResult.IsValid Then
            bAnyFailures = True
        End If
        
        ' Add to summary
        With wsSummary
            .Range("A" & lRow).Value = oResult.FileName
            
            ' Status with color coding
            .Range("B" & lRow).Value = IIf(oResult.IsValid, "PASSED", "FAILED")
            If oResult.IsValid Then
                .Range("B" & lRow).Interior.Color = RGB(200, 255, 200) ' Light green
            Else
                .Range("B" & lRow).Interior.Color = RGB(255, 200, 200) ' Light red
            End If
            
            .Range("C" & lRow).Value = oResult.TotalRecords
            .Range("D" & lRow).Value = oResult.ErrorCount
            .Range("E" & lRow).Value = oResult.WarningCount
            .Range("F" & lRow).Value = oResult.FileType
            .Range("G" & lRow).Value = "Manual"
            
            Dim sUniqueSheetName As String
            sUniqueSheetName = GetUniqueSheetName(wbReport, oResult.FileType)
            .Range("H" & lRow).Value = "See Sheet: " & sUniqueSheetName
        End With
        
        ' Create detailed sheet for this file
        Set wsDetail = wbReport.Sheets.Add(After:=wbReport.Sheets(wbReport.Sheets.count))
        wsDetail.Name = sUniqueSheetName
        
        ' COLOR THE DETAIL SHEET TAB
        If oResult.IsValid Then
            wsDetail.Tab.Color = RGB(0, 176, 80)  ' Green
        Else
            wsDetail.Tab.Color = RGB(192, 0, 0)   ' Red
        End If
        
        ' Fill detailed sheet
        Call FillDetailedSheet(wsDetail, oResult, sFilePath)
        
        lRow = lRow + 1
        i = i + 1
    Next sFileData
    
    ' COLOR THE SUMMARY TAB
    If bAnyFailures Then
        wsSummary.Tab.Color = RGB(192, 0, 0)  ' Red
    Else
        wsSummary.Tab.Color = RGB(0, 176, 80) ' Green
    End If
    
    ' Auto-fit columns in summary
    wsSummary.Columns("A:H").AutoFit
    wsSummary.Activate
    
    Application.StatusBar = False
    Call ReorganizeSheetsByStatus(wbReport)
    Set CreateValidationReportWithManualType = wbReport
End Function

Private Sub FillDetailedSheet(ws As Worksheet, oResult As ValidationResult, sFilePath As String)
    Dim lRow As Long
    Dim oError As ValidationError
    Dim i As Long
    
    With ws
        ' Header information
        .Range("A1").Value = "FILE VALIDATION DETAIL"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        
        ' File information
        lRow = 3
        .Range("A" & lRow).Value = "File Path:"
        .Range("B" & lRow).Value = sFilePath
        lRow = lRow + 1
        
        .Range("A" & lRow).Value = "File Name:"
        .Range("B" & lRow).Value = oResult.FileName
        lRow = lRow + 1
        
        .Range("A" & lRow).Value = "File Type:"
        .Range("B" & lRow).Value = oResult.FileType
        lRow = lRow + 1
        
        .Range("A" & lRow).Value = "Group ID:"
        .Range("B" & lRow).Value = oResult.GroupID
        lRow = lRow + 1
        
        .Range("A" & lRow).Value = "Group Name:"
        .Range("B" & lRow).Value = oResult.GroupName
        lRow = lRow + 1
        
        .Range("A" & lRow).Value = "Processed Date:"
        .Range("B" & lRow).Value = oResult.ProcessedDate
        lRow = lRow + 1
        
        .Range("A" & lRow).Value = "Total Records:"
        .Range("B" & lRow).Value = oResult.TotalRecords
        lRow = lRow + 1
        
        .Range("A" & lRow).Value = "Validation Status:"
        .Range("B" & lRow).Value = IIf(oResult.IsValid, "PASSED", "FAILED")
        If oResult.IsValid Then
            .Range("B" & lRow).Interior.Color = RGB(200, 255, 200)
        Else
            .Range("B" & lRow).Interior.Color = RGB(255, 200, 200)
        End If
        
        ' New way - summary matrix
        Call CreateValidationSummaryMatrix(ws, oResult, lRow)
        lRow = lRow + 15  ' Adjust based on how many rows the matrix uses
        
        ' Errors section
        If oResult.ErrorCount > 0 Then
            lRow = lRow + 2
            .Range("A" & lRow).Value = "ERRORS (" & oResult.ErrorCount & " found)"
            .Range("A" & lRow).Font.Bold = True
            .Range("A" & lRow).Interior.Color = RGB(255, 200, 200)
            
            lRow = lRow + 1
            .Range("A" & lRow).Value = "Row"
            .Range("B" & lRow).Value = "Field"
            .Range("C" & lRow).Value = "Error Message"
            .Range("A" & lRow & ":C" & lRow).Font.Bold = True
            .Range("A" & lRow & ":C" & lRow).Borders.LineStyle = xlContinuous
            
            ' Add each error
            For i = 1 To oResult.Errors.count
                Set oError = oResult.Errors(i)
                lRow = lRow + 1
                .Range("A" & lRow).Value = oError.RowNumber
                .Range("B" & lRow).Value = oError.fieldName
                .Range("C" & lRow).Value = oError.ErrorMessage
            Next i
        End If
        
        ' Warnings section
        If oResult.WarningCount > 0 Then
            lRow = lRow + 2
            .Range("A" & lRow).Value = "WARNINGS (" & oResult.WarningCount & " found)"
            .Range("A" & lRow).Font.Bold = True
            .Range("A" & lRow).Interior.Color = RGB(255, 255, 200)
            
            lRow = lRow + 1
            .Range("A" & lRow).Value = "Row"
            .Range("B" & lRow).Value = "Field"
            .Range("C" & lRow).Value = "Warning Message"
            .Range("A" & lRow & ":C" & lRow).Font.Bold = True
            .Range("A" & lRow & ":C" & lRow).Borders.LineStyle = xlContinuous
            
            ' Add each warning
            For i = 1 To oResult.Warnings.count
                Set oError = oResult.Warnings(i)
                lRow = lRow + 1
                .Range("A" & lRow).Value = oError.RowNumber
                .Range("B" & lRow).Value = oError.fieldName
                .Range("C" & lRow).Value = oError.ErrorMessage
            Next i
        End If
        
        ' Auto-fit columns
        .Columns("A:C").AutoFit
    End With
End Sub

' Rest of the functions remain the same...
Private Function GetGroupName(sGroupID As String) As String
    Dim wsGroups As Worksheet
    Dim lLastRow As Long
    Dim lRow As Long
    
    On Error Resume Next
    Set wsGroups = ThisWorkbook.Worksheets("Parsed_SFTPfiles")
    
    If wsGroups Is Nothing Then
        GetGroupName = "Unknown"
        Exit Function
    End If
    
    lLastRow = wsGroups.Cells(wsGroups.Rows.count, "K").End(xlUp).Row
    
    ' Look for matching GroupID in column K
    For lRow = 2 To lLastRow
        If CStr(wsGroups.Cells(lRow, "K").Value) = sGroupID Then
            ' Get Group Name from column L (adjust if it's different)
            GetGroupName = CStr(wsGroups.Cells(lRow, "J").Value)
            Exit Function
        End If
    Next lRow
    
    GetGroupName = "Not Found"
End Function

Private Sub CreateValidationSummaryMatrix(ws As Worksheet, oResult As ValidationResult, lStartRow As Long)
    Dim lRow As Long
    lRow = lStartRow
    
    ' Add title for summary section
    ws.Range("A" & lRow).Value = "VALIDATION SUMMARY"
    ws.Range("A" & lRow).Font.Bold = True
    ws.Range("A" & lRow).Font.Size = 12
    ws.Range("A" & lRow & ":E" & lRow).Merge
    ws.Range("A" & lRow & ":E" & lRow).Interior.Color = RGB(200, 200, 200)
    ws.Range("A" & lRow & ":E" & lRow).HorizontalAlignment = xlCenter
    lRow = lRow + 2
    
    ' Create the matrix headers
    ws.Range("A" & lRow).Value = "Field Name"
    ws.Range("B" & lRow).Value = "Blank Field Check"
    ws.Range("C" & lRow).Value = "Max Character Check"
    ws.Range("D" & lRow).Value = "Format Check"
    
    ' Format headers
    ws.Range("A" & lRow & ":D" & lRow).Font.Bold = True
    ws.Range("A" & lRow & ":D" & lRow).Interior.Color = RGB(220, 220, 220)
    ws.Range("A" & lRow & ":D" & lRow).Borders.LineStyle = xlContinuous
    
    ' Define the 12 fields to summarize
    Dim vFields As Variant
    vFields = Array("FirstName", "LastName", "DOB", "Gender", _
                    "ZipCode", "Address1", "City", "State", _
                    "EffectiveDate", "GroupID", "ServiceOffering", "MemberID")
    
    ' Analyze errors to determine status for each field/check combination
    Dim dictStatus As Object
    Set dictStatus = CreateObject("Scripting.Dictionary")
    
    ' Initialize all as PASS
    Dim i As Integer
    For i = 0 To UBound(vFields)
        dictStatus(vFields(i) & "_Blank") = "PASS"
        dictStatus(vFields(i) & "_MaxChar") = "PASS"
        dictStatus(vFields(i) & "_Format") = "PASS"
    Next i
    
    ' Check errors collection to update status
    Dim oError As ValidationError
    For Each oError In oResult.Errors
        Dim sFieldName As String
        sFieldName = oError.fieldName
        
        ' Determine which type of check failed
        If InStr(oError.ErrorMessage, "blank") > 0 Or _
           InStr(oError.ErrorMessage, "Required") > 0 Then
            dictStatus(sFieldName & "_Blank") = "REVIEW REQUIRED"
            
        ElseIf InStr(oError.ErrorMessage, "maximum length") > 0 Or _
               InStr(oError.ErrorMessage, "Exceeds") > 0 Then
            dictStatus(sFieldName & "_MaxChar") = "REVIEW REQUIRED"
            
        ElseIf InStr(oError.ErrorMessage, "format") > 0 Or _
               InStr(oError.ErrorMessage, "Invalid") > 0 Then
            dictStatus(sFieldName & "_Format") = "REVIEW REQUIRED"
        End If
    Next
    
    ' Populate the matrix
    For i = 0 To UBound(vFields)
        lRow = lRow + 1
        ws.Range("A" & lRow).Value = vFields(i)
        
        ' Blank Field Check status
        ws.Range("B" & lRow).Value = dictStatus(vFields(i) & "_Blank")
        If dictStatus(vFields(i) & "_Blank") = "REVIEW REQUIRED" Then
            ws.Range("B" & lRow).Interior.Color = RGB(255, 255, 200) ' Light yellow
            ws.Range("B" & lRow).Font.Color = RGB(200, 0, 0) ' Red text
        Else
            ws.Range("B" & lRow).Interior.Color = RGB(200, 255, 200) ' Light green
            ws.Range("B" & lRow).Font.Color = RGB(0, 128, 0) ' Green text
        End If
        
        ' Max Character Check status
        ws.Range("C" & lRow).Value = dictStatus(vFields(i) & "_MaxChar")
        If dictStatus(vFields(i) & "_MaxChar") = "REVIEW REQUIRED" Then
            ws.Range("C" & lRow).Interior.Color = RGB(255, 255, 200)
            ws.Range("C" & lRow).Font.Color = RGB(200, 0, 0)
        Else
            ws.Range("C" & lRow).Interior.Color = RGB(200, 255, 200)
            ws.Range("C" & lRow).Font.Color = RGB(0, 128, 0)
        End If
        
        ' Format Check status
        ws.Range("D" & lRow).Value = dictStatus(vFields(i) & "_Format")
        If dictStatus(vFields(i) & "_Format") = "REVIEW REQUIRED" Then
            ws.Range("D" & lRow).Interior.Color = RGB(255, 255, 200)
            ws.Range("D" & lRow).Font.Color = RGB(200, 0, 0)
        Else
            ws.Range("D" & lRow).Interior.Color = RGB(200, 255, 200)
            ws.Range("D" & lRow).Font.Color = RGB(0, 128, 0)
        End If
        
        ' Add borders to the row
        ws.Range("A" & lRow & ":D" & lRow).Borders.LineStyle = xlContinuous
    Next i
    
    ' Auto-fit columns
    ws.Columns("A:D").AutoFit
    
    ' Add a note about details below
    lRow = lRow + 2
    ws.Range("A" & lRow).Value = "See detailed validation results below"
    ws.Range("A" & lRow).Font.Italic = True
    ws.Range("A" & lRow).Font.Color = RGB(100, 100, 100)
End Sub

Private Function SheetExists(wb As Workbook, sSheetName As String) As Boolean
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = wb.Worksheets(sSheetName)
    On Error GoTo 0
    
    SheetExists = Not ws Is Nothing
End Function

Private Function GetUniqueSheetName(wb As Workbook, sBaseName As String) As String
    Dim sSheetName As String
    Dim iCounter As Integer
    
    ' Clean the base name first
    sSheetName = CleanSheetName(sBaseName, 0)
    
    ' If this name doesn't exist, use it
    If Not SheetExists(wb, sSheetName) Then
        GetUniqueSheetName = sSheetName
        Exit Function
    End If
    
    ' Otherwise, add a counter until we find a unique name
    iCounter = 2
    Do While SheetExists(wb, sSheetName & "_" & iCounter)
        iCounter = iCounter + 1
    Loop
    
    ' Make sure we don't exceed 31 characters
    If Len(sSheetName & "_" & iCounter) > 31 Then
        sSheetName = Left(sSheetName, 31 - Len("_" & iCounter))
    End If
    
    GetUniqueSheetName = sSheetName & "_" & iCounter
End Function

Private Function CleanSheetName(sGroupName As String, Optional iFileIndex As Long = 0) As String
    ' Clean and truncate GroupName for Excel sheet naming rules
    Dim sCleanName As String
    
    ' Start with the GroupName
    sCleanName = sGroupName
    
    ' Remove invalid characters for Excel sheet names
    ' Note: Underscores are ALLOWED in sheet names, so we keep them
    sCleanName = Replace(sCleanName, "/", "-")
    sCleanName = Replace(sCleanName, "\", "-")
    sCleanName = Replace(sCleanName, "?", "")
    sCleanName = Replace(sCleanName, "*", "")
    sCleanName = Replace(sCleanName, "[", "(")
    sCleanName = Replace(sCleanName, "]", ")")
    sCleanName = Replace(sCleanName, ":", "-")
    
    ' Truncate to 31 characters (Excel's limit)
    If Len(sCleanName) > 31 Then
        ' If we have a file index, leave room for it
        If iFileIndex > 0 Then
            ' Leave room for something like "_1" at the end
            sCleanName = Left(sCleanName, 28) & "_" & iFileIndex
        Else
            sCleanName = Left(sCleanName, 31)
        End If
    End If
    
    ' Handle empty or whitespace-only results
    If Trim(sCleanName) = "" Then
        sCleanName = "File" & iFileIndex
    End If
    
    CleanSheetName = sCleanName
End Function

Private Function ActiveChecker(sEndDate As String) As Boolean
    ' Handle blank dates (blank = active forever)
    If Trim(sEndDate) = "" Then
        ActiveChecker = True
        Exit Function
    End If
    
    ' Check if date is in the future or past
    On Error GoTo DateError
    If CDate(sEndDate) > Date Then
        ActiveChecker = True  ' Future date = still active
    Else
        ActiveChecker = False ' Past date = inactive
    End If
    Exit Function
    
DateError:
    ' If date can't be parsed, treat as active
    ActiveChecker = True
End Function

Private Sub ReorganizeSheetsByStatus(wb As Workbook)
    Dim ws As Worksheet
    Dim i As Integer
    Dim FailedSheets As Collection
    Dim PassedSheets As Collection
    
    Set FailedSheets = New Collection
    Set PassedSheets = New Collection
    
    ' Categorize sheets (skip Summary sheet)
    For Each ws In wb.Worksheets
        If ws.Name <> "Summary" Then
            If ws.Tab.Color = RGB(192, 0, 0) Then  ' Red = failed
                FailedSheets.Add ws.Name
            Else
                PassedSheets.Add ws.Name
            End If
        End If
    Next ws
    
    ' Move failed sheets to be right after Summary
    i = 2  ' Start position after Summary
    Dim sheetName As Variant
    
    ' Move failed sheets first
    For Each sheetName In FailedSheets
        wb.Worksheets(sheetName).Move After:=wb.Worksheets(i - 1)
        i = i + 1
    Next sheetName
    
    ' Passed sheets will naturally be after failed ones
End Sub

' Get list of available file types from Filetype Mapping sheet
Private Function GetAvailableFileTypes() As Collection
    Const sPROC_NAME As String = "GetAvailableFileTypes"
    
    Dim colFileTypes As Collection
    Set colFileTypes = New Collection
    
    Dim wsMapping As Worksheet
    Dim lLastRow As Long
    Dim lRow As Long
    Dim sFileType As String
    
    On Error GoTo ErrorHandler
    
    Set wsMapping = ThisWorkbook.Worksheets("Filetype Mapping")
    lLastRow = wsMapping.Cells(wsMapping.Rows.count, "A").End(xlUp).Row
    
    ' Collect all file types (skip header row)
    For lRow = 2 To lLastRow
        sFileType = Trim(wsMapping.Cells(lRow, "A").Value)
        If sFileType <> "" Then
            colFileTypes.Add sFileType
        End If
    Next lRow
    
    Set GetAvailableFileTypes = colFileTypes
    Exit Function
    
ErrorHandler:
    Call ErrorHandler_Central(sPROC_NAME, Err.Number, Err.Description)
    Set GetAvailableFileTypes = colFileTypes
End Function

' Show dialog to select file type
Private Function ShowFileTypeSelectionDialog(colFileTypes As Collection) As String
    Dim frm As frmFileTypeSelector
    
    ' Create new instance of the form
    Set frm = New frmFileTypeSelector
    
    ' Initialize the form with file types
    frm.InitializeForm colFileTypes
    
    ' Show the form modally (waits for user action)
    frm.Show vbModal
    
    ' Get the result
    If frm.Cancelled Then
        ShowFileTypeSelectionDialog = ""
    Else
        ShowFileTypeSelectionDialog = frm.SelectedFileType
    End If
    
    ' Clean up
    Unload frm
    Set frm = Nothing
End Function

' Select a single file for validation
Private Function SelectSingleValidationFile() As String
    Dim fd As fileDialog
    Dim sFilePath As String
    
    Set fd = Application.fileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "Select File to Validate"
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        .Filters.Add "All Files", "*.*"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            sFilePath = .SelectedItems(1)
        Else
            sFilePath = ""
        End If
    End With
    
    Set fd = Nothing
    SelectSingleValidationFile = sFilePath
End Function

