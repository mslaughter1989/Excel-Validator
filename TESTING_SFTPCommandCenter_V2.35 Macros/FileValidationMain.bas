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
    
    If Not oFileInfo.isValid Then
        oResult.AddError 0, "Filename", "No matching pattern found for filename: " & sFileName
        GoTo Cleanup
    End If
    ' Populate ValidationResult with file information
    oResult.fileName = sFileName
    oResult.filePath = sFilePath
    oResult.fileType = oFileInfo.fileType
    oResult.groupID = oFileInfo.groupID
    oResult.ProcessedDate = Now
    
    ' Get Group Name from lookup sheet
    oResult.groupName = GetGroupName(oFileInfo.groupID)
    
    ' Record validation checks performed
    oResult.AddValidationCheck "Filename Pattern", "Matched: " & sFileName
    oResult.AddValidationCheck "FileType Identified", oFileInfo.fileType
    oResult.AddValidationCheck "GroupID Extracted", oFileInfo.groupID
    
    ' Stage 2: Get column mapping for FileType
    Call frmProgress.UpdateProgress("Loading column mappings...", 15)
    
    Dim oMapping As ColumnMapping
    oMapping = GetColumnMapping(oFileInfo.fileType)
    
    If oMapping.fileType = "" Then
        oResult.AddError 0, "FileType", "No column mapping found for FileType: " & oFileInfo.fileType
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
    oResult.fileName = sFileName
    oResult.filePath = sFilePath
    oResult.fileType = sFileType
    oResult.ProcessedDate = Now
    oResult.groupID = "MANUAL"  ' Use MANUAL as placeholder for manual validations
    oResult.groupName = "Manual Validation - " & sFileType
    
    ' Record validation checks performed
    oResult.AddValidationCheck "Filename", sFileName
    oResult.AddValidationCheck "FileType (Manual)", sFileType
    
    ' Stage 1: Get column mapping for selected FileType
    Call frmProgress.UpdateProgress("Loading column mappings...", 15)
    
    Dim oMapping As ColumnMapping
    oMapping = GetColumnMapping(sFileType)
    
    If oMapping.fileType = "" Then
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
    oFileInfo.fileType = sFileType
    oFileInfo.groupID = "MANUAL"  ' Use MANUAL for manual validations
    oFileInfo.isValid = True
    
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
    Dim colMemberIDs As Object ' Changed from Collection to Dictionary for better tracking
    Set colMemberIDs = CreateObject("Scripting.Dictionary")
    
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
        
        ' ===================================================================
        ' ENHANCED DUPLICATE CHECK: MemberID + ServiceOffering
        ' ===================================================================
        If oMapping.memberID > 0 And oMapping.memberID <= UBound(vData, 2) Then
            Dim sMemberID As String
            Dim sServiceOffering As String
            Dim sEffectiveEndDate As String
            Dim bCurrentIsActive As Boolean
            Dim sComboKey As String
            
            ' Get MemberID for current row
            sMemberID = Trim(CStr(vData(lRow, oMapping.memberID)))
            
            ' Get ServiceOffering if column exists
            If oMapping.serviceOffering > 0 And oMapping.serviceOffering <= UBound(vData, 2) Then
                sServiceOffering = Trim(CStr(vData(lRow, oMapping.serviceOffering)))
            Else
                sServiceOffering = "UNKNOWN" ' Default if not mapped
            End If
            
            ' Get EffectiveEndDate if column exists
            If oMapping.effectiveEndDate > 0 And oMapping.effectiveEndDate <= UBound(vData, 2) Then
                sEffectiveEndDate = Trim(CStr(vData(lRow, oMapping.effectiveEndDate)))
            Else
                sEffectiveEndDate = "" ' Default to blank if column not mapped
            End If
            
            ' Check if current record is active
            bCurrentIsActive = ActiveChecker(sEffectiveEndDate)
            
            ' Create combo key: "MemberID|ServiceOffering"
            sComboKey = sMemberID & "|" & sServiceOffering
            
            If sMemberID <> "" And sServiceOffering <> "" Then
                ' Check if this combo key already exists
                If colMemberIDs.Exists(sComboKey) Then
                    ' Duplicate found - retrieve previous record info
                    Dim sPreviousData As String
                    sPreviousData = colMemberIDs(sComboKey)
                    
                    ' Parse previous data: "RowNumber|IsActive"
                    Dim vPrevParts As Variant
                    vPrevParts = Split(sPreviousData, "|")
                    
                    Dim lPreviousRow As Long
                    Dim bPreviousIsActive As Boolean
                    
                    lPreviousRow = CLng(vPrevParts(0))
                    bPreviousIsActive = CBool(vPrevParts(1))
                    
                    ' CRITICAL: Only flag as error if BOTH occurrences are active
                    If bCurrentIsActive And bPreviousIsActive Then
                        ' Both records are active - this is an ERROR
                        oResult.AddError lRow, "MemberID", _
                            "Duplicate active record found: MemberID=" & sMemberID & _
                            ", ServiceOffering=" & sServiceOffering & _
                            " (both records have blank or future EffectiveEndDate). " & _
                            "Previous occurrence at row " & lPreviousRow & "."
                        
                        Debug.Print "DUPLICATE ERROR: Row " & lRow & " - " & sComboKey & " (both active)"
                    ElseIf bCurrentIsActive And Not bPreviousIsActive Then
                        ' Current is active, previous is inactive - Update to track current as the active one
                        colMemberIDs(sComboKey) = lRow & "|" & bCurrentIsActive
                        Debug.Print "DUPLICATE OK: Row " & lRow & " - " & sComboKey & " (current active, previous inactive)"
                    ElseIf Not bCurrentIsActive And bPreviousIsActive Then
                        ' Current is inactive, previous is active - Don't update, keep tracking the active one
                        Debug.Print "DUPLICATE OK: Row " & lRow & " - " & sComboKey & " (current inactive, previous active)"
                    Else
                        ' Both inactive - Update to track most recent
                        colMemberIDs(sComboKey) = lRow & "|" & bCurrentIsActive
                        Debug.Print "DUPLICATE OK: Row " & lRow & " - " & sComboKey & " (both inactive)"
                    End If
                Else
                    ' First occurrence of this combo - add to dictionary
                    ' Store: "RowNumber|IsActive"
                    colMemberIDs.Add sComboKey, lRow & "|" & bCurrentIsActive
                End If
            End If
        End If
        
        ' Validate GroupID matches expected (skip for manual validations)
        If oFileInfo.groupID <> "MANUAL" Then
            If oMapping.groupID > 0 And oMapping.groupID <= UBound(vData, 2) Then
                ' Only log once at the beginning
                If lRow = 2 Then
                    oResult.AddValidationCheck "Group ID Match", "Verifying against expected: " & oFileInfo.groupID
                    Debug.Print "===== GROUP ID VALIDATION DEBUG ====="
                    Debug.Print "Expected GroupID(s): '" & oFileInfo.groupID & "'"
                    Debug.Print "Length of expected: " & Len(oFileInfo.groupID)
                End If
                
                Dim sGroupID As String
                sGroupID = Trim(CStr(vData(lRow, oMapping.groupID)))
                
                ' Debug output for first few rows
                If lRow <= 5 Then
                    Debug.Print "Row " & lRow & " - Raw GroupID value: '" & vData(lRow, oMapping.groupID) & "'"
                    Debug.Print "Row " & lRow & " - After Trim/CStr: '" & sGroupID & "'"
                End If
                
                ' Check if GroupID matches (can be comma-separated list)
                Dim bGroupIDValid As Boolean
                bGroupIDValid = False
                
                If InStr(oFileInfo.groupID, "_") > 0 Then
                    ' Multiple GroupIDs separated by underscore
                    Dim vGroupIDs As Variant
                    vGroupIDs = Split(oFileInfo.groupID, "_")
                    
                    Dim vGroupID As Variant
                    For Each vGroupID In vGroupIDs
                        If Trim(CStr(vGroupID)) = sGroupID Then
                            bGroupIDValid = True
                            Exit For
                        End If
                    Next vGroupID
                Else
                    ' Single GroupID
                    bGroupIDValid = (sGroupID = oFileInfo.groupID)
                End If
                
                If lRow <= 5 Then
                    Debug.Print "Row " & lRow & " - GroupID Valid: " & bGroupIDValid
                End If
                
                ' Only add error if GroupID is not valid AND not blank
                If Not bGroupIDValid And sGroupID <> "" Then
                    oResult.AddError lRow, "GroupID", "GroupID mismatch. Expected: " & oFileInfo.groupID & ", Found: " & sGroupID
                    
                    ' Extra debug for errors
                    Debug.Print "ERROR at Row " & lRow & ": GroupID '" & sGroupID & "' not found in '" & oFileInfo.groupID & "'"
                End If
                
                ' End debug output after first few rows
                If lRow = 5 Then
                    Debug.Print "===== END GROUP ID DEBUG (only showing first few rows) ====="
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
    If oMapping.serviceOffering > 0 And oMapping.serviceOffering <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.serviceOffering), "ServiceOffering", lRow, colRules, oResult)
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
    Static blankCheckCount As Object
    Static maxCharCheckCount As Object
    Static formatCheckCount As Object
    
    If checkedFields Is Nothing Then Set checkedFields = New Collection
    If blankCheckCount Is Nothing Then Set blankCheckCount = CreateObject("Scripting.Dictionary")
    If maxCharCheckCount Is Nothing Then Set maxCharCheckCount = CreateObject("Scripting.Dictionary")
    If formatCheckCount Is Nothing Then Set formatCheckCount = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    checkedFields.Add sFieldType, sFieldType
    If Err.Number = 0 Then  ' First time checking this field type
        oResult.AddValidationCheck sFieldType & " Field", "Validating across all records"
        ' Initialize counters
        blankCheckCount(sFieldType) = 0
        maxCharCheckCount(sFieldType) = 0
        formatCheckCount(sFieldType) = 0
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
        blankCheckCount(sFieldType) = blankCheckCount(sFieldType) + 1
        
        If sValue = "" Or IsNull(vFieldValue) Or Trim(sValue) = "" Then
            oResult.AddError lRowNumber, sFieldType, "Required field is blank"
            ' Only log first occurrence to avoid spam
            If blankCheckCount(sFieldType) = 1 Then
                oResult.AddValidationCheck "Blank Check - " & sFieldType, "PERFORMED - Found blank value(s)"
            End If
            Exit Sub  ' No need to check other validations if field is blank
        Else
            ' Log success on first valid check
            If blankCheckCount(sFieldType) = 1 Then
                oResult.AddValidationCheck "Blank Check - " & sFieldType, "PERFORMED - Field populated"
            End If
        End If
    End If
    
    ' OTHER VALIDATION CHECKS (Max Length, Format, etc.)
    ' Skip other validations if field is empty (and not required)
    If sValue = "" Then Exit Sub
    
    ' Max character check (still uses Column Checks sheet)
    If oRule.MaxLength > 0 Then
        maxCharCheckCount(sFieldType) = maxCharCheckCount(sFieldType) + 1
        
        If Len(sValue) > oRule.MaxLength Then
            oResult.AddError lRowNumber, sFieldType, "Exceeds maximum length of " & oRule.MaxLength & " characters (found " & Len(sValue) & ")"
            ' Only log first occurrence
            If maxCharCheckCount(sFieldType) = 1 Then
                oResult.AddValidationCheck "Max Length Check - " & sFieldType, "PERFORMED - Found violation(s) (Max: " & oRule.MaxLength & ")"
            End If
        Else
            ' Log success on first valid check
            If maxCharCheckCount(sFieldType) = 1 Then
                oResult.AddValidationCheck "Max Length Check - " & sFieldType, "PERFORMED - Within limit (Max: " & oRule.MaxLength & ")"
            End If
        End If
    End If
    
    ' Min character check (still uses Column Checks sheet)
    If oRule.MinLength > 0 Then
        If Len(sValue) < oRule.MinLength Then
            oResult.AddError lRowNumber, sFieldType, "Below minimum length of " & oRule.MinLength & " characters (found " & Len(sValue) & ")"
        End If
    End If
    
    ' Format validation (still uses Column Checks sheet patterns)
    If oRule.FormatPattern <> "" Then
        formatCheckCount(sFieldType) = formatCheckCount(sFieldType) + 1
        
        If Not ValidateFieldFormat(sValue, sFieldType, oRule.FormatPattern) Then
            oResult.AddError lRowNumber, sFieldType, "Invalid format for " & sFieldType
            ' Only log first occurrence
            If formatCheckCount(sFieldType) = 1 Then
                oResult.AddValidationCheck "Format Check - " & sFieldType, "PERFORMED - Found invalid format(s)"
            End If
        Else
            ' Log success on first valid check
            If formatCheckCount(sFieldType) = 1 Then
                oResult.AddValidationCheck "Format Check - " & sFieldType, "PERFORMED - Valid format"
            End If
        End If
    End If
    
    ' DEBUG: Log to Immediate Window
    If lRowNumber Mod 100 = 0 Then
        Debug.Print "Processing row " & lRowNumber & " - Field: " & sFieldType & " - Value Length: " & Len(sValue)
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
    oRegex.pattern = "^\d{5}(-\d{4})?$"
    ValidateZipCode = oRegex.Test(sValue)
End Function

Private Function ValidateNameFormat(sValue As String) As Boolean
    Dim oRegex As Object
    Set oRegex = CreateObject("VBScript.RegExp")
    
    ' Allow letters, spaces, hyphens, apostrophes, periods
    oRegex.pattern = "^[a-zA-Z][a-zA-Z\s\-'\.]{1,49}$"
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
    
    oRegex.pattern = sPattern
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
        If Not oResult.isValid Then
            bAnyFailures = True
        End If
        
        ' Add to summary
        With wsSummary
            .Range("A" & lRow).Value = oResult.fileName
            
            ' Status with color coding
            .Range("B" & lRow).Value = IIf(oResult.isValid, "PASSED", "FAILED")
            If oResult.isValid Then
                .Range("B" & lRow).Interior.Color = RGB(200, 255, 200) ' Light green
            Else
                .Range("B" & lRow).Interior.Color = RGB(255, 200, 200) ' Light red
            End If
            
            .Range("C" & lRow).Value = oResult.TotalRecords
            .Range("D" & lRow).Value = oResult.ErrorCount
            .Range("E" & lRow).Value = oResult.WarningCount
            .Range("F" & lRow).Value = oResult.fileType
            .Range("G" & lRow).Value = oResult.groupID
            
            Dim sUniqueSheetName As String
            sUniqueSheetName = GetUniqueSheetName(wbReport, oResult.groupName)
            .Range("H" & lRow).Value = "See Sheet: " & sUniqueSheetName
        End With
        
        ' Create detailed sheet for this file
        Set wsDetail = wbReport.Sheets.Add(After:=wbReport.Sheets(wbReport.Sheets.count))
        wsDetail.Name = sUniqueSheetName
        
        ' COLOR THE DETAIL SHEET TAB
        If oResult.isValid Then
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
        If Not oResult.isValid Then
            bAnyFailures = True
        End If
        
        ' Add to summary
        With wsSummary
            .Range("A" & lRow).Value = oResult.fileName
            
            ' Status with color coding
            .Range("B" & lRow).Value = IIf(oResult.isValid, "PASSED", "FAILED")
            If oResult.isValid Then
                .Range("B" & lRow).Interior.Color = RGB(200, 255, 200) ' Light green
            Else
                .Range("B" & lRow).Interior.Color = RGB(255, 200, 200) ' Light red
            End If
            
            .Range("C" & lRow).Value = oResult.TotalRecords
            .Range("D" & lRow).Value = oResult.ErrorCount
            .Range("E" & lRow).Value = oResult.WarningCount
            .Range("F" & lRow).Value = oResult.fileType
            .Range("G" & lRow).Value = "Manual"
            
            Dim sUniqueSheetName As String
            sUniqueSheetName = GetUniqueSheetName(wbReport, oResult.fileType)
            .Range("H" & lRow).Value = "See Sheet: " & sUniqueSheetName
        End With
        
        ' Create detailed sheet for this file
        Set wsDetail = wbReport.Sheets.Add(After:=wbReport.Sheets(wbReport.Sheets.count))
        wsDetail.Name = sUniqueSheetName
        
        ' COLOR THE DETAIL SHEET TAB
        If oResult.isValid Then
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
        .Range("B" & lRow).Value = oResult.fileName
        lRow = lRow + 1
        
        .Range("A" & lRow).Value = "File Type:"
        .Range("B" & lRow).Value = oResult.fileType
        lRow = lRow + 1
        
        .Range("A" & lRow).Value = "Group ID:"
        .Range("B" & lRow).Value = oResult.groupID
        lRow = lRow + 1
        
        .Range("A" & lRow).Value = "Group Name:"
        .Range("B" & lRow).Value = oResult.groupName
        lRow = lRow + 1
        
        .Range("A" & lRow).Value = "Processed Date:"
        .Range("B" & lRow).Value = oResult.ProcessedDate
        lRow = lRow + 1
        
        .Range("A" & lRow).Value = "Total Records:"
        .Range("B" & lRow).Value = oResult.TotalRecords
        lRow = lRow + 1
        
        .Range("A" & lRow).Value = "Validation Status:"
        .Range("B" & lRow).Value = IIf(oResult.isValid, "PASSED", "FAILED")
        If oResult.isValid Then
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
    
    lLastRow = wsGroups.Cells(wsGroups.Rows.count, "K").End(xlUp).row
    
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

Private Function ActiveChecker(sEffectiveEndDate As String) As Boolean
    ' A record is active if:
    ' 1. EffectiveEndDate is blank/empty, OR
    ' 2. EffectiveEndDate is a future date (greater than or equal to today)
    
    Dim todayDate As Date
    todayDate = Date ' Current date
    
    ' Trim whitespace
    sEffectiveEndDate = Trim(sEffectiveEndDate)
    
    ' If blank or empty, record is active
    If sEffectiveEndDate = "" Or sEffectiveEndDate = "0" Then
        ActiveChecker = True
        Exit Function
    End If
    
    ' Try to parse as date
    If IsDate(sEffectiveEndDate) Then
        Dim endDate As Date
        endDate = CDate(sEffectiveEndDate)
        
        ' Active if end date is today or in the future
        ActiveChecker = (endDate >= todayDate)
    Else
        ' If we can't parse the date, assume active (conservative approach)
        ' Could also log this as a warning
        ActiveChecker = True
    End If
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
    lLastRow = wsMapping.Cells(wsMapping.Rows.count, "A").End(xlUp).row
    
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

' Function to verify all validation checks are completing
Public Sub DebugValidationChecks(oResult As ValidationResult)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Debug_" & Format(Now, "hhmmss")
    
    Dim lRow As Long
    lRow = 1
    
    ' Header
    ws.Range("A" & lRow).Value = "VALIDATION DEBUG REPORT"
    ws.Range("A" & lRow).Font.Bold = True
    ws.Range("A" & lRow).Font.Size = 14
    lRow = lRow + 2
    
    ' Summary stats
    ws.Range("A" & lRow).Value = "Total Records Processed:"
    ws.Range("B" & lRow).Value = oResult.TotalRecords
    lRow = lRow + 1
    
    ws.Range("A" & lRow).Value = "Total Errors:"
    ws.Range("B" & lRow).Value = oResult.ErrorCount
    lRow = lRow + 1
    
    ws.Range("A" & lRow).Value = "Total Warnings:"
    ws.Range("B" & lRow).Value = oResult.WarningCount
    lRow = lRow + 1
    
    ws.Range("A" & lRow).Value = "Validation Checks Performed:"
    ws.Range("B" & lRow).Value = oResult.ValidationChecks.count
    lRow = lRow + 2
    
    ' List all validation checks performed
    ws.Range("A" & lRow).Value = "VALIDATION CHECKS PERFORMED:"
    ws.Range("A" & lRow).Font.Bold = True
    lRow = lRow + 1
    
    Dim i As Long
    For i = 1 To oResult.ValidationChecks.count
        ws.Range("A" & lRow).Value = i
        ws.Range("B" & lRow).Value = oResult.ValidationChecks(i)
        lRow = lRow + 1
    Next i
    
    lRow = lRow + 1
    
    ' Expected validation checks for all 12 fields
    ws.Range("A" & lRow).Value = "EXPECTED VALIDATION CHECKS:"
    ws.Range("A" & lRow).Font.Bold = True
    lRow = lRow + 1
    
    Dim vFields As Variant
    vFields = Array("FirstName", "LastName", "DOB", "Gender", _
                    "ZipCode", "Address1", "City", "State", _
                    "EffectiveDate", "GroupID", "ServiceOffering", "MemberID")
    
    Dim vCheckTypes As Variant
    vCheckTypes = Array("Field", "Blank Check", "Max Length Check", "Format Check")
    
    Dim sField As Variant, sCheckType As Variant
    Dim bFound As Boolean
    Dim sCheckName As String
    
    For Each sField In vFields
        For Each sCheckType In vCheckTypes
            sCheckName = sCheckType & " - " & sField
            bFound = False
            
            ' Check if this validation was performed
            For i = 1 To oResult.ValidationChecks.count
                If InStr(oResult.ValidationChecks(i), sField) > 0 And _
                   InStr(oResult.ValidationChecks(i), Left(sCheckType, 5)) > 0 Then
                    bFound = True
                    Exit For
                End If
            Next i
            
            ws.Range("A" & lRow).Value = sCheckName
            ws.Range("B" & lRow).Value = IIf(bFound, "? PERFORMED", "? MISSING")
            
            If Not bFound Then
                ws.Range("B" & lRow).Font.Color = RGB(255, 0, 0)
                ws.Range("B" & lRow).Font.Bold = True
            Else
                ws.Range("B" & lRow).Font.Color = RGB(0, 128, 0)
            End If
            
            lRow = lRow + 1
        Next sCheckType
    Next sField
    
    ' Auto-fit columns
    ws.Columns("A:B").AutoFit
    
    MsgBox "Debug report created in sheet: " & ws.Name, vbInformation, "Debug Complete"
End Sub

' Function to test validation rules loading
Public Sub TestValidationRulesLoading()
    Dim colRules As Collection
    Set colRules = LoadValidationRules()
    
    Debug.Print "=== VALIDATION RULES LOADED ==="
    Debug.Print "Total rules: " & colRules.count
    
    Dim vFields As Variant
    vFields = Array("FirstName", "LastName", "DOB", "Gender", _
                    "ZipCode", "Address1", "City", "State", _
                    "EffectiveDate", "GroupID", "ServiceOffering", "MemberID")
    
    Dim sField As Variant
    Dim sRuleData As String
    Dim vParts As Variant
    
    For Each sField In vFields
        On Error Resume Next
        sRuleData = ""
        sRuleData = colRules.Item(CStr(sField))
        
        If sRuleData <> "" Then
            vParts = Split(sRuleData, "|")
            Debug.Print sField & ":"
            Debug.Print "  Required: " & vParts(1)
            Debug.Print "  MaxLength: " & vParts(2)
            Debug.Print "  MinLength: " & vParts(3)
            Debug.Print "  Pattern: " & vParts(4)
        Else
            Debug.Print sField & ": NO RULE FOUND"
        End If
        On Error GoTo 0
    Next sField
    
    Debug.Print "=== END RULES CHECK ==="
End Sub
Public Sub TestGroupIDValidation()
    Dim wsPatterns As Worksheet
    Dim wsMapping As Worksheet
    Dim lRow As Long
    
    Debug.Print "========================================="
    Debug.Print "GROUP ID VALIDATION DIAGNOSTIC TEST"
    Debug.Print "========================================="
    Debug.Print ""
    
    ' 1. Check Parsed_SFTPfiles sheet
    Debug.Print "1. CHECKING PARSED_SFTPFILES SHEET:"
    Debug.Print "------------------------------------"
    
    On Error Resume Next
    Set wsPatterns = ThisWorkbook.Worksheets("Parsed_SFTPfiles")
    If wsPatterns Is Nothing Then
        Debug.Print "ERROR: Parsed_SFTPfiles sheet not found!"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Check a few rows from column K
    For lRow = 2 To Application.Min(10, wsPatterns.Cells(wsPatterns.Rows.count, "K").End(xlUp).row)
        Dim sGroupIDValue As String
        sGroupIDValue = wsPatterns.Cells(lRow, "K").Value
        
        Debug.Print "Row " & lRow & " Column K:"
        Debug.Print "  Raw Value: '" & sGroupIDValue & "'"
        Debug.Print "  Length: " & Len(sGroupIDValue)
        Debug.Print "  Has underscore: " & (InStr(sGroupIDValue, "_") > 0)
        
        ' Check for hidden characters
        Dim iChar As Integer
        Dim sCharInfo As String
        sCharInfo = ""
        For iChar = 1 To Len(sGroupIDValue)
            If Mid(sGroupIDValue, iChar, 1) = "_" Then
                sCharInfo = sCharInfo & "[_]"
            ElseIf Asc(Mid(sGroupIDValue, iChar, 1)) < 32 Or Asc(Mid(sGroupIDValue, iChar, 1)) > 126 Then
                sCharInfo = sCharInfo & "[ASCII:" & Asc(Mid(sGroupIDValue, iChar, 1)) & "]"
            Else
                sCharInfo = sCharInfo & Mid(sGroupIDValue, iChar, 1)
            End If
        Next iChar
        Debug.Print "  Character analysis: " & sCharInfo
        
        ' If it contains underscores, split and show parts
        If InStr(sGroupIDValue, "_") > 0 Then
            Dim vParts As Variant
            vParts = Split(sGroupIDValue, "_")
            Debug.Print "  Split into " & UBound(vParts) + 1 & " parts:"
            Dim i As Integer
            For i = 0 To UBound(vParts)
                Debug.Print "    Part " & i & ": '" & vParts(i) & "' (Len: " & Len(vParts(i)) & ")"
            Next i
        End If
        Debug.Print ""
    Next lRow
    
    ' 2. Check Column_Mapping sheet
    Debug.Print "2. CHECKING COLUMN_MAPPING SHEET:"
    Debug.Print "------------------------------------"
    
    On Error Resume Next
    Set wsMapping = ThisWorkbook.Worksheets("Filetype Mapping")
    If wsMapping Is Nothing Then
        Debug.Print "ERROR: Column_Mapping sheet not found!"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Find GroupID column mapping for different file types
    Dim lLastRow As Long
    lLastRow = wsMapping.Cells(wsMapping.Rows.count, "A").End(xlUp).row
    
    For lRow = 2 To lLastRow
        Dim sFileType As String
        Dim lGroupIDCol As Long
        
        sFileType = wsMapping.Cells(lRow, "A").Value
        
        ' Find which column has GroupID
        Dim lCol As Long
        For lCol = 2 To 20  ' Check first 20 columns
            If UCase(Trim(wsMapping.Cells(lRow, lCol).Value)) = "GROUPID" Or _
               UCase(Trim(wsMapping.Cells(lRow, lCol).Value)) = "GROUP ID" Or _
               UCase(Trim(wsMapping.Cells(lRow, lCol).Value)) = "GID" Then
                lGroupIDCol = lCol - 1  ' Convert to 0-based for column index
                Exit For
            End If
        Next lCol
        
        If lGroupIDCol > 0 Then
            Debug.Print "FileType '" & sFileType & "': GroupID is in column " & lGroupIDCol
        End If
    Next lRow
    
    Debug.Print ""
    Debug.Print "3. TEST COMPARISONS:"
    Debug.Print "------------------------------------"
    
    ' Test the specific case mentioned
    Dim sTestExpected As String
    Dim sTestFound As String
    
    sTestExpected = "566484_808737_566484_689449"
    sTestFound = "689449"
    
    Debug.Print "Testing the reported issue:"
    Debug.Print "  Expected: '" & sTestExpected & "'"
    Debug.Print "  Found: '" & sTestFound & "'"
    
    Dim vTestParts As Variant
    vTestParts = Split(sTestExpected, "_")
    Debug.Print "  Split expected into " & UBound(vTestParts) + 1 & " parts"
    
    Dim bTestMatch As Boolean
    bTestMatch = False
    
    Dim j As Integer
    For j = 0 To UBound(vTestParts)
        Debug.Print "    Comparing '" & sTestFound & "' with part " & j & " '" & vTestParts(j) & "'"
        If Trim(CStr(vTestParts(j))) = Trim(sTestFound) Then
            Debug.Print "      *** MATCH! ***"
            bTestMatch = True
        End If
    Next j
    
    If Not bTestMatch Then
        Debug.Print "  WARNING: No match found - there may be hidden characters!"
    End If
    
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "END DIAGNOSTIC TEST"
    Debug.Print "========================================="
End Sub

' Also add this function to help clean GroupIDs if needed
Public Function CleanGroupID(sGroupID As String) As String
    ' Remove any non-printable characters
    Dim sClean As String
    Dim i As Integer
    
    sClean = ""
    For i = 1 To Len(sGroupID)
        Dim sChar As String
        sChar = Mid(sGroupID, i, 1)
        
        ' Only keep numbers, letters, and underscores
        If (sChar >= "0" And sChar <= "9") Or _
           (sChar >= "A" And sChar <= "Z") Or _
           (sChar >= "a" And sChar <= "z") Or _
           sChar = "_" Then
            sClean = sClean & sChar
        End If
    Next i
    
    CleanGroupID = sClean
End Function


