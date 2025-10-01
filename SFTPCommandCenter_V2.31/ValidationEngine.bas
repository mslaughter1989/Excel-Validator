Attribute VB_Name = "ValidationEngine"
Option Explicit

Type ColumnMapping
    FileType As String
    FirstName As Integer
    LastName As Integer
    DOB As Integer
    Gender As Integer
    ZipCode As Integer
    Address1 As Integer
    Address2 As Integer
    City As Integer
    State As Integer
    EffectiveDate As Integer
    ServiceOffering As Integer
    MemberID As Integer
    GroupID As Integer
    EffectiveEndDate As Integer
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
    lLastRow = wsMapping.Cells(wsMapping.Rows.count, "A").End(xlUp).Row
    
    ' Initialize with empty values
    oMapping.FileType = ""
    oMapping.FirstName = 0
    oMapping.LastName = 0
    oMapping.DOB = 0
    oMapping.Gender = 0
    oMapping.ZipCode = 0
    oMapping.Address1 = 0
    oMapping.City = 0
    oMapping.State = 0
    oMapping.EffectiveDate = 0
    oMapping.GroupID = 0
    oMapping.ServiceOffering = 0
    oMapping.MemberID = 0
    
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
                .GroupID = wsMapping.Cells(lRow, "K").Value
                .ServiceOffering = wsMapping.Cells(lRow, "L").Value
                .MemberID = wsMapping.Cells(lRow, "M").Value
                .EffectiveEndDate = wsMapping.Cells(lRow, "N").Value
            End With
            
            ' Return the Type directly (no Set keyword)
            GetColumnMapping = oMapping
            Exit Function
        End If
    Next lRow
    
    ' FileType not found - return empty mapping
    GetColumnMapping = oMapping
    Exit Function
    
ErrorHandler:
    Call ErrorHandler_Central(sPROC_NAME, Err.Number, Err.Description, sFileType)
    GetColumnMapping = oMapping
End Function

Public Function LoadValidationRules() As Collection
    Const sPROC_NAME As String = "LoadValidationRules"
    
    Dim colRules As New Collection
    Dim wsRules As Worksheet
    Dim lRow As Long
    Dim lLastRow As Long
    
    On Error GoTo ErrorHandler
    
    ' Reference the Column Checks sheet
    Set wsRules = ThisWorkbook.Worksheets("Column Checks")
    lLastRow = wsRules.Cells(wsRules.Rows.count, "A").End(xlUp).Row
    
    ' Expected Column Layout in "Column Checks" sheet:
    ' Column A: Field Name (e.g., FirstName, LastName, DOB, Gender, ZipCode, Address1, City, State, EffectiveDate, ServiceOffering, etc.)
    ' Column B: Required (TRUE/FALSE)
    ' Column C: Max Length
    ' Column D: Min Length
    ' Column E: Format Pattern (regex or format type)
    ' Column F: Custom Function (optional)
    
    ' Process each row (assuming row 1 has headers)
    For lRow = 2 To lLastRow
        Dim sFieldType As String
        Dim sRuleData As String
        
        sFieldType = Trim(wsRules.Cells(lRow, "A").Value)
        
        If sFieldType <> "" Then
            ' Create a delimited string with rule data
            sRuleData = sFieldType & "|" & _
                       wsRules.Cells(lRow, "B").Value & "|" & _
                       wsRules.Cells(lRow, "C").Value & "|" & _
                       wsRules.Cells(lRow, "D").Value & "|" & _
                       wsRules.Cells(lRow, "E").Value & "|" & _
                       wsRules.Cells(lRow, "F").Value
            
            ' Add the string data to collection using FieldType as key
            On Error Resume Next
            colRules.Add sRuleData, sFieldType
            If Err.Number <> 0 Then
                Debug.Print "Duplicate field type found: " & sFieldType
                Err.Clear
            End If
            On Error GoTo ErrorHandler
            
            ' Log the loaded rule
            Debug.Print "Loaded rule for: " & sFieldType & " - Required: " & wsRules.Cells(lRow, "B").Value & _
                       ", MaxLen: " & wsRules.Cells(lRow, "C").Value & ", MinLen: " & wsRules.Cells(lRow, "D").Value
        End If
    Next lRow
    
    Debug.Print "Total validation rules loaded: " & colRules.count
    
    Set LoadValidationRules = colRules
    Exit Function
    
ErrorHandler:
    MsgBox "Error loading validation rules from 'Column Checks' sheet: " & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "LoadValidationRules Error"
    
    If colRules Is Nothing Then Set colRules = New Collection
    Set LoadValidationRules = colRules
End Function

Private Sub ValidateRowFields(vData As Variant, lRow As Long, oMapping As ColumnMapping, colRules As Collection, oResult As ValidationResult)
    ' This validates all 12 fields in your system
    
    Debug.Print ">>> EXECUTING: ValidationEngine.ValidateRowFields at row " & lRow
    
    ' 1. First Name
    If oMapping.FirstName > 0 And oMapping.FirstName <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.FirstName), "FirstName", lRow, colRules, oResult)
    End If
    
    ' 2. Last Name
    If oMapping.LastName > 0 And oMapping.LastName <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.LastName), "LastName", lRow, colRules, oResult)
    End If
    
    ' 3. DOB
    If oMapping.DOB > 0 And oMapping.DOB <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.DOB), "DOB", lRow, colRules, oResult)
    End If
    
    ' 4. Gender
    If oMapping.Gender > 0 And oMapping.Gender <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.Gender), "Gender", lRow, colRules, oResult)
    End If
    
    ' 5. Zip Code
    If oMapping.ZipCode > 0 And oMapping.ZipCode <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.ZipCode), "ZipCode", lRow, colRules, oResult)
    End If
    
    ' 6. Address1
    If oMapping.Address1 > 0 And oMapping.Address1 <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.Address1), "Address1", lRow, colRules, oResult)
    End If
    
    ' 7. Address2 (if you have this field)
    If oMapping.Address2 > 0 And oMapping.Address2 <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.Address2), "Address2", lRow, colRules, oResult)
    End If
    
    ' 8. City
    If oMapping.City > 0 And oMapping.City <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.City), "City", lRow, colRules, oResult)
    End If
    
    ' 9. State
    If oMapping.State > 0 And oMapping.State <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.State), "State", lRow, colRules, oResult)
    End If
    
    ' 10. Effective Date
    If oMapping.EffectiveDate > 0 And oMapping.EffectiveDate <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.EffectiveDate), "EffectiveDate", lRow, colRules, oResult)
    End If
    
    ' 11. Service Offering
    If oMapping.ServiceOffering > 0 And oMapping.ServiceOffering <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.ServiceOffering), "ServiceOffering", lRow, colRules, oResult)
    End If
    
    ' 12. Member ID (or whatever your 12th field is)
    If oMapping.MemberID > 0 And oMapping.MemberID <= UBound(vData, 2) Then
        Call ValidateField(vData(lRow, oMapping.MemberID), "MemberID", lRow, colRules, oResult)
    End If
End Sub

Private Sub ValidateField(vFieldValue As Variant, sFieldType As String, lRowNumber As Long, colRules As Collection, oResult As ValidationResult)

    ' Track which field types have been validated (only report once per field type)
    Static checkedFields As Collection
    If checkedFields Is Nothing Then Set checkedFields = New Collection
    
    On Error Resume Next
    checkedFields.Add sFieldType, sFieldType  ' <-- FIXED: Use sFieldType
    If Err.Number = 0 Then  ' First time checking this field type
        oResult.AddValidationCheck sFieldType & " Field", "Checking across all records"
    End If
    On Error GoTo 0
    
    ' Get the validation rule for this field
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
    
    ' Required field check (blank check)
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

Private Function FindValidationRule(colRules As Collection, sFieldType As String) As ValidationRule
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
    
    ' Get the rule data string from collection
    sRuleData = colRules.Item(sFieldType)
    
    ' Parse the delimited string back into ValidationRule Type
    vParts = Split(sRuleData, "|")
    
    If UBound(vParts) >= 5 Then
        oRule.FieldType = vParts(0)
        oRule.Required = (UCase(vParts(1)) = "Y" Or UCase(vParts(1)) = "TRUE" Or UCase(vParts(1)) = "y")
        oRule.MaxLength = Val(vParts(2))
        oRule.MinLength = Val(vParts(3))
        oRule.FormatPattern = vParts(4)
        oRule.CustomFunction = vParts(5)
        
        FindValidationRule = oRule
        Exit Function
    End If
    
NotFound:
    ' Return empty rule if not found or error
    FindValidationRule = emptyRule
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
