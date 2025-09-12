Attribute VB_Name = "ValidationEngine"
Option Explicit

Public Type ColumnMapping
    FileType As String
    FirstName As Long
    LastName As Long
    DOB As Long
    Gender As Long
    ZipCode As Long
    Address1 As Long
    City As Long
    State As Long
    EffectiveDate As Long
    GID As Long
    ServiceOffering As Long
    CMID As Long
End Type

Public Type ValidationRule
    FieldType As String
    Required As Boolean
    MaxLength As Long
    MinLength As Long
    FormatPattern As String
    CustomFunction As String
End Type

Public Function GetColumnMapping(sFileType As String) As ColumnMapping
    Const sPROC_NAME As String = "GetColumnMapping"
    
    Dim oMapping As ColumnMapping
    Dim wsMapping As Worksheet
    Dim lLastRow As Long
    Dim lRow As Long
    
    On Error GoTo ErrorHandler
    
    Set wsMapping = ThisWorkbook.Worksheets("Filetype Mapping")
    lLastRow = wsMapping.Cells(wsMapping.Rows.Count, "A").End(xlUp).row
    
    ' Find the FileType row
    For lRow = 2 To lLastRow
        If UCase(wsMapping.Cells(lRow, "A").Value) = UCase(sFileType) Then
            With oMapping
                .FileType = sFileType
                .FirstName = wsMapping.Cells(lRow, "B").Value
                .LastName = wsMapping.Cells(lRow, "C").Value
                .DOB = wsMapping.Cells(lRow, "D").Value
                .Gender = wsMapping.Cells(lRow, "E").Value
                .ZipCode = wsMapping.Cells(lRow, "F").Value
                .Address1 = wsMapping.Cells(lRow, "G").Value
                .City = wsMapping.Cells(lRow, "H").Value
                .State = wsMapping.Cells(lRow, "I").Value
                .EffectiveDate = wsMapping.Cells(lRow, "J").Value
                .GID = wsMapping.Cells(lRow, "K").Value
                .ServiceOffering = wsMapping.Cells(lRow, "L").Value
                .CMID = wsMapping.Cells(lRow, "M").Value
            End With
            
            Set GetColumnMapping = CreateColumnMappingObject(oMapping)
            Exit Function
        End If
    Next lRow
    
    ' FileType not found
    Set GetColumnMapping = Nothing
    Exit Function
    
ErrorHandler:
    Call ErrorHandler_Central(sPROC_NAME, err.Number, err.description, sFileType)
    Set GetColumnMapping = Nothing
End Function

Public Function LoadValidationRules() As Collection
    Const sPROC_NAME As String = "LoadValidationRules"
    
    Dim colRules As Collection
    Dim wsRules As Worksheet
    Dim lLastRow As Long
    Dim lRow As Long
    
    On Error GoTo ErrorHandler
    
    Set colRules = New Collection
    Set wsRules = ThisWorkbook.Worksheets("Column Checks")
    lLastRow = wsRules.Cells(wsRules.Rows.Count, "A").End(xlUp).row
    
    For lRow = 2 To lLastRow
        Dim oRule As ValidationRule
        
        With oRule
            .FieldType = wsRules.Cells(lRow, "A").Value
            .Required = (UCase(wsRules.Cells(lRow, "B").Value) = "TRUE")
            .MaxLength = wsRules.Cells(lRow, "C").Value
            .MinLength = wsRules.Cells(lRow, "D").Value
            .FormatPattern = wsRules.Cells(lRow, "E").Value
            .CustomFunction = wsRules.Cells(lRow, "F").Value
        End With
        
        colRules.Add CreateValidationRuleObject(oRule), oRule.FieldType
    Next lRow
    
    Set LoadValidationRules = colRules
    Exit Function
    
ErrorHandler:
    Call ErrorHandler_Central(sPROC_NAME, err.Number, err.description)
    If colRules Is Nothing Then Set colRules = New Collection
    Set LoadValidationRules = colRules
End Function

Public Function ValidateCSVData(vData As Variant, oMapping As ColumnMapping, colRules As Collection, oFileInfo As FileInfo) As ValidationResult
    Const sPROC_NAME As String = "ValidateCSVData"
    
    Dim oResult As ValidationResult
    Set oResult = New ValidationResult
    
    Dim lRow As Long
    Dim lTotalRows As Long
    Dim colCMIDs As Collection
    Set colCMIDs = New Collection
    
    On Error GoTo ErrorHandler
    
    lTotalRows = UBound(vData, 1)
    oResult.TotalRecords = lTotalRows - 1 ' Subtract header row
    
    ' Validate each data row (skip header)
    For lRow = 2 To lTotalRows
        ' Update progress
        If lRow Mod 1000 = 0 Then
            Dim lPercent As Long
            lPercent = 50 + ((lRow / lTotalRows) * 30) ' Progress from 50% to 80%
            Call frmProgress.UpdateProgress("Validating record " & (lRow - 1) & " of " & (lTotalRows - 1), lPercent)
        End If
        
        ' Validate individual fields
        Call ValidateRowFields(vData, lRow, oMapping, colRules, oResult)
        
        ' Check CMID uniqueness
        If oMapping.CMID > 0 Then
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
                On Error GoTo ErrorHandler
            End If
        End If
        
        ' Validate GID matches expected
        If oMapping.GID > 0 Then
            Dim sGID As String
            sGID = CStr(vData(lRow, oMapping.GID))
            
            If sGID <> oFileInfo.GroupID Then
                oResult.AddError lRow - 1, "GID", "GID mismatch. Expected: " & oFileInfo.GroupID & ", Found: " & sGID
            End If
        End If
    Next lRow
    
    oResult.ValidationComplete = True
    Set ValidateCSVData = oResult
    Exit Function
    
ErrorHandler:
    Call ErrorHandler_Central(sPROC_NAME, err.Number, err.description)
    oResult.AddError 0, "System", "Validation failed due to system error: " & err.description
    Set ValidateCSVData = oResult
End Function

Private Sub ValidateRowFields(vData As Variant, lRow As Long, oMapping As ColumnMapping, colRules As Collection, oResult As ValidationResult)
    ' Validate First Name
    If oMapping.FirstName > 0 Then
        Call ValidateField vData(lRow, oMapping.FirstName), "FirstName", lRow - 1, colRules, oResult
    End If
    
    ' Validate Last Name
    If oMapping.LastName > 0 Then
        Call ValidateField vData(lRow, oMapping.LastName), "LastName", lRow - 1, colRules, oResult
    End If
    
    ' Validate DOB
    If oMapping.DOB > 0 Then
        Call ValidateField vData(lRow, oMapping.DOB), "DOB", lRow - 1, colRules, oResult
    End If
    
    ' Validate Gender
    If oMapping.Gender > 0 Then
        Call ValidateField vData(lRow, oMapping.Gender), "Gender", lRow - 1, colRules, oResult
    End If
    
    ' Validate Zip Code
    If oMapping.ZipCode > 0 Then
        Call ValidateField vData(lRow, oMapping.ZipCode), "ZipCode", lRow - 1, colRules, oResult
    End If
    
    ' Validate Address1
    If oMapping.Address1 > 0 Then
        Call ValidateField vData(lRow, oMapping.Address1), "Address1", lRow - 1, colRules, oResult
    End If
    
    ' Validate City
    If oMapping.City > 0 Then
        Call ValidateField vData(lRow, oMapping.City), "City", lRow - 1, colRules, oResult
    End If
    
    ' Validate State
    If oMapping.State > 0 Then
        Call ValidateField vData(lRow, oMapping.State), "State", lRow - 1, colRules, oResult
    End If
    
    ' Validate Effective Date
    If oMapping.EffectiveDate > 0 Then
        Call ValidateField vData(lRow, oMapping.EffectiveDate), "EffectiveDate", lRow - 1, colRules, oResult
    End If
    
    ' Validate Service Offering
    If oMapping.ServiceOffering > 0 Then
        Call ValidateField vData(lRow, oMapping.ServiceOffering), "ServiceOffering", lRow - 1, colRules, oResult
    End If
End Sub

Private Sub ValidateField(vFieldValue As Variant, sFieldType As String, lRowNumber As Long, colRules As Collection, oResult As ValidationResult)
    On Error Resume Next
    
    Dim oRule As ValidationRule
    Set oRule = colRules.item(sFieldType)
    
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
    
    ' Custom function validation
    If oRule.CustomFunction <> "" Then
        If Not CallCustomValidationFunction(sValue, oRule.CustomFunction) Then
            oResult.AddError lRowNumber, sFieldType, "Failed custom validation: " & oRule.CustomFunction
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
