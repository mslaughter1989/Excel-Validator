Attribute VB_Name = "FileValidationMain"
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
    Dim i As Long
    Dim oRule As ValidationRule
    
    ' Search through collection for matching field type
    For i = 1 To colRules.Count
        oRule = colRules(i)
        If UCase(oRule.FieldType) = UCase(sFieldType) Then
            GetValidationRule = oRule
            Exit Function
        End If
    Next i
    
    ' Return empty rule if not found
    Dim emptyRule As ValidationRule
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
