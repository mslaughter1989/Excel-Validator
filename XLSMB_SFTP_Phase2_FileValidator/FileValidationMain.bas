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
        GoTo CleanUp
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
    
CleanUp:
    Call RestoreApplication
    Exit Sub
    
ErrorHandler:
    Call RestoreApplication
    Call ErrorHandler_Central(sPROC_NAME, Err.Number, Err.Description)
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
    
        ' ADD THE DEBUG HERE - RIGHT AFTER THE GetFileNameFromPath LINE
    MsgBox "Full file path: " & sFilePath & vbCrLf & _
           "Extracted filename: " & sFileName
           
    Dim oFileInfo As FileInfo
    oFileInfo = MatchFilenamePattern(sFileName)
    
    If Not oFileInfo.IsValid Then
        oResult.AddError 0, "Filename", "No matching pattern found for filename: " & sFileName
        GoTo CleanUp
    End If
    
    ' Stage 2: Get column mapping for FileType
    Call frmProgress.UpdateProgress("Loading column mappings...", 15)
    
    Dim oMapping As ColumnMapping
    oMapping = GetColumnMapping(oFileInfo.FileType)
    
    If oMapping.FileType = "" Then
        oResult.AddError 0, "FileType", "No column mapping found for FileType: " & oFileInfo.FileType
        GoTo CleanUp
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
        GoTo CleanUp
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
    
CleanUp:
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
    GoTo CleanUp
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
                If Err.Number <> 0 Then
                    ' Duplicate found
                    oResult.AddError lRow - 1, "CMID", "Duplicate CMID found: " & sCMID
                    Err.Clear
                End If
                On Error GoTo 0
            End If
        End If
        
        ' Validate GID matches expected
        If oMapping.GID > 0 And oMapping.GID <= UBound(vData, 2) Then
            Dim sGID As String
            sGID = CStr(vData(lRow, oMapping.GID))
            
            If sGID <> oFileInfo.groupID Then
                oResult.AddError lRow - 1, "GID", "GID mismatch. Expected: " & oFileInfo.groupID & ", Found: " & sGID
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
    If oMapping.zipCode > 0 And oMapping.zipCode <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.zipCode), "ZipCode", lRow - 1, colRules, oResult)
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
    
    If Err.Number <> 0 Then
        ' No rule found for this field type
        Err.Clear
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
        oRule.MaxLength = val(vParts(2))
        oRule.MinLength = val(vParts(3))
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
        
        ' Add to summary
        With wsSummary
            .Range("A" & lRow).Value = oResult.fileName
            
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
            .Range("G" & lRow).Value = oResult.groupID
            .Range("H" & lRow).Value = "See Sheet: " & "File" & i
        End With
        
        ' Create detailed sheet for this file
        Set wsDetail = wbReport.Sheets.Add(After:=wbReport.Sheets(wbReport.Sheets.count))
        wsDetail.Name = "File" & i
        
        ' Fill detailed sheet
        Call FillDetailedSheet(wsDetail, oResult, CStr(sFilePath))
        
        lRow = lRow + 1
        i = i + 1
    Next sFilePath
    
    ' Auto-fit columns in summary
    wsSummary.Columns("A:H").AutoFit
    wsSummary.Activate
    
    Application.StatusBar = False
    Set CreateValidationReport = wbReport
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
                .Range("B" & lRow).Value = oError.FieldName
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
                .Range("B" & lRow).Value = oError.FieldName
                .Range("C" & lRow).Value = oError.ErrorMessage
            Next i
        End If
        
        ' Auto-fit columns
        .Columns("A:C").AutoFit
    End With
End Sub

