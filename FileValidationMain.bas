Attribute VB_Name = "FileValidationMain"
Option Explicit

' Main entry point for file validation system
Public Sub StartFileValidation()
    Const sPROC_NAME As String = "StartFileValidation"
    
    On Error GoTo ErrorHandler
    
    ' Initialize application settings
    Call InitializeApplication
    
    ' Show file selection dialog
    Dim sSelectedFile As String
    sSelectedFile = SelectValidationFile()
    
    If sSelectedFile = "" Then
        MsgBox "No file selected. Validation cancelled.", vbInformation, "File Validation"
        GoTo Cleanup
    End If
    
    ' Start validation process
    Dim oResult As ValidationResult
    Set oResult = ValidateSelectedFile(sSelectedFile)
    
    ' Display results
    Call DisplayValidationResults(oResult)
    
Cleanup:
    Call RestoreApplication
    Exit Sub
    
ErrorHandler:
    Call RestoreApplication
    Call ErrorHandler_Central(sPROC_NAME, err.Number, err.description)
End Sub

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
    
    ' Stage 6: Generate validation report
    Call frmProgress.UpdateProgress("Generating report...", 85)
    
    Dim sReportPath As String
    sReportPath = GenerateValidationReport(oResult, oFileInfo, sFilePath)
    oResult.ReportPath = sReportPath
    
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
    Call ErrorHandler_Central(sPROC_NAME, err.Number, err.description, sFilePath)
    If oResult Is Nothing Then Set oResult = New ValidationResult
    oResult.AddError 0, "System", "Validation failed due to system error: " & err.description
    GoTo Cleanup
End Function

Private Sub ValidateCSVDataIntoResult(vData As Variant, oMapping As ColumnMapping, colRules As Collection, oFileInfo As FileInfo, oResult As ValidationResult)
    Dim lRow As Long
    Dim lTotalRows As Long
    Dim colCMIDs As Collection
    Set colCMIDs = New Collection
    
    lTotalRows = UBound(vData, 1)
    oResult.TotalRecords = lTotalRows - 1 ' Subtract header row
    
    ' Validate each data row (skip header)
    For lRow = 2 To lTotalRows
        ' Update progress occasionally
        If lRow Mod 1000 = 0 Then
            Dim lPercent As Long
            lPercent = 50 + ((lRow / lTotalRows) * 30) ' Progress from 50% to 80%
            Call frmProgress.UpdateProgress("Validating record " & (lRow - 1) & " of " & (lTotalRows - 1), lPercent)
        End If
        
        ' Validate individual fields
        Call ValidateRowFields(vData, lRow, oMapping, colRules, oResult)
        
        ' Check CMID uniqueness
        If oMapping.CMID > 0 And oMapping.CMID <= UBound(vData, 2) Then
            Dim sCMID As String
            sCMID = CStr(vData(lRow, oMapping.CMID))
            
            If sCMID <> "" Then
                On Error Resume Next
                colCMIDs.Add sCMID, sCMID
                If err.Number <> 0 Then
                    ' Duplicate found
                    oResult.AddError lRow - 1, "CMID", "Duplicate CMID found: " & sCMID
                    err.Clear
                End If
                On Error GoTo 0
            End If
        End If
        
        ' Validate GID matches expected
        If oMapping.GID > 0 And oMapping.GID <= UBound(vData, 2) Then
            Dim sGID As String
            sGID = CStr(vData(lRow, oMapping.GID))
            
            If sGID <> oFileInfo.GroupID Then
                oResult.AddError lRow - 1, "GID", "GID mismatch. Expected: " & oFileInfo.GroupID & ", Found: " & sGID
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
        Call ValidateField(vData(lRow, oMapping.FirstName), "FirstName", lRow - 1, colRules, oResult)
    End If
    
    ' Validate Last Name
    If oMapping.LastName > 0 And oMapping.LastName <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.LastName), "LastName", lRow - 1, colRules, oResult)
    End If
    
    ' Validate DOB
    If oMapping.DOB > 0 And oMapping.DOB <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.DOB), "DOB", lRow - 1, colRules, oResult)
    End If
    
    ' Validate Gender
    If oMapping.Gender > 0 And oMapping.Gender <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.Gender), "Gender", lRow - 1, colRules, oResult)
    End If
    
    ' Validate Zip Code
    If oMapping.ZipCode > 0 And oMapping.ZipCode <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.ZipCode), "ZipCode", lRow - 1, colRules, oResult)
    End If
    
    ' Validate Address1
    If oMapping.Address1 > 0 And oMapping.Address1 <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.Address1), "Address1", lRow - 1, colRules, oResult)
    End If
    
    ' Validate City
    If oMapping.City > 0 And oMapping.City <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.City), "City", lRow - 1, colRules, oResult)
    End If
    
    ' Validate State
    If oMapping.State > 0 And oMapping.State <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.State), "State", lRow - 1, colRules, oResult)
    End If
    
    ' Validate Effective Date
    If oMapping.EffectiveDate > 0 And oMapping.EffectiveDate <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.EffectiveDate), "EffectiveDate", lRow - 1, colRules, oResult)
    End If
    
    ' Validate Service Offering
    If oMapping.ServiceOffering > 0 And oMapping.ServiceOffering <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.ServiceOffering), "ServiceOffering", lRow - 1, colRules, oResult)
    End If
End Sub

Private Sub ValidateField(vFieldValue As Variant, sFieldType As String, lRowNumber As Long, colRules As Collection, oResult As ValidationResult)
    On Error Resume Next
    
    Dim oRule As ValidationRule
    oRule = GetValidationRule(colRules, sFieldType)
    
    If err.Number <> 0 Then
        ' No rule found for this field type
        err.Clear
        Exit Sub
    End If
    
    On Error GoTo 0
    
    Dim sValue As String
    sValue = CStr(vFieldValue)
    
    ' Required field check
    If oRule.Required And (sValue = "" Or IsNull(vFieldValue)) Then
        oResult.AddError lRowNumber, sFieldType, "Required field is blank"
        Exit Sub
    End If
    
    ' Skip other validations if field is empty and not required
    If sValue = "" Then Exit Sub
    
    ' Length validations
    If oRule.MaxLength > 0 And Len(sValue) > oRule.MaxLength Then
        oResult.AddError lRowNumber, sFieldType, "Exceeds maximum length of " & oRule.MaxLength & " characters"
    End If
    
    If oRule.MinLength > 0 And Len(sValue) < oRule.MinLength Then
        oResult.AddError lRowNumber, sFieldType, "Below minimum length of " & oRule.MinLength & " characters"
    End If
    
    ' Format validation
    If oRule.FormatPattern <> "" Then
        If Not ValidateFieldFormat(sValue, sFieldType, oRule.FormatPattern) Then
            oResult.AddError lRowNumber, sFieldType, "Invalid format for " & sFieldType
        End If
    End If
End Sub

Private Function GetValidationRule(colRules As Collection, sFieldType As String) As ValidationRule
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
    ValidateDateFormat = (err.Number = 0)
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
