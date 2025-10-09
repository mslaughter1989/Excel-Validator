Attribute VB_Name = "ValidationTestGenerator"
Option Explicit

' ==============================================================================
' VALIDATION TEST FILE GENERATOR - FIXED VERSION
' Generates files with CORRECT column order per FileType mapping
' ==============================================================================

Sub GenerateValidationTestFiles()
    Dim wsParsed As Worksheet
    Dim wsMapping As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim sFileType As String
    Dim sFileName As String
    Dim sGroupName As String
    Dim sGroupID As String
    Dim outputFolder As String
    Dim fso As Object
    Dim filesCreated As Long
    Dim reportMsg As String
    
    On Error GoTo ErrorHandler
    
    Set wsParsed = ThisWorkbook.Worksheets("Parsed_SFTPfiles")
    Set wsMapping = ThisWorkbook.Worksheets("Filetype Mapping")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    outputFolder = SelectOutputFolder()
    If outputFolder = "" Then
        MsgBox "No folder selected. Operation cancelled.", vbInformation
        Exit Sub
    End If
    
    Dim testFolder As String
    testFolder = fso.BuildPath(outputFolder, "Validation_Test_Files_" & Format(Now, "yyyymmdd_hhnnss"))
    If Not fso.FolderExists(testFolder) Then
        fso.CreateFolder testFolder
    End If
    
    filesCreated = 0
    reportMsg = "VALIDATION TEST FILES GENERATED:" & vbCrLf & vbCrLf
    
    lastRow = wsParsed.Cells(wsParsed.Rows.count, "A").End(xlUp).row
    
    For i = 2 To lastRow
        sFileType = Trim(CStr(wsParsed.Cells(i, "O").Value))
        sGroupName = Trim(CStr(wsParsed.Cells(i, "J").Value))
        sGroupID = Trim(CStr(wsParsed.Cells(i, "K").Value))
        
        If sFileType <> "" Then
            sFileName = GenerateValidationTestFileName(wsParsed, i)
            
            If CreateValidationTestCSVFile(testFolder, sFileName, sFileType, sGroupName, sGroupID, wsMapping) Then
                filesCreated = filesCreated + 1
                reportMsg = reportMsg & sFileName & " (" & sFileType & ")" & vbCrLf
            End If
        End If
    Next i
    
    reportMsg = reportMsg & vbCrLf & "Total files created: " & filesCreated & vbCrLf
    reportMsg = reportMsg & "Location: " & testFolder & vbCrLf & vbCrLf
    reportMsg = reportMsg & "Each file contains ~30 records with various validation errors." & vbCrLf
    reportMsg = reportMsg & "Run FileValidationMain to test all validation rules!"
    
    MsgBox reportMsg, vbInformation, "Validation Test File Generation Complete"
    
    Shell "explorer.exe " & Chr(34) & testFolder & Chr(34), vbNormalFocus
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating validation test files: " & Err.Description, vbCritical
End Sub

' ==============================================================================
' Generate filename based on pattern
' ==============================================================================
Private Function GenerateValidationTestFileName(wsParsed As Worksheet, rowNum As Long) As String
    Dim pattern As String
    Dim dateFormat As String
    Dim testDate As String
    
    pattern = Trim(CStr(wsParsed.Cells(rowNum, "A").Value))
    dateFormat = Trim(CStr(wsParsed.Cells(rowNum, "F").Value))
    
    Select Case UCase(dateFormat)
        Case "MMDDYYYY"
            testDate = Format(Date, "mmddyyyy")
        Case "YYYYMMDD"
            testDate = Format(Date, "yyyymmdd")
        Case "MMDDYY"
            testDate = Format(Date, "mmddyy")
        Case Else
            testDate = Format(Date, "mmddyyyy")
    End Select
    
    pattern = Replace(pattern, "mmddyyyy", testDate)
    pattern = Replace(pattern, "yyyymmdd", testDate)
    pattern = Replace(pattern, "mmddyy", testDate)
    
    GenerateValidationTestFileName = pattern
End Function

' ==============================================================================
' Create CSV with CORRECT column mapping per FileType
' ==============================================================================
Private Function CreateValidationTestCSVFile(folderPath As String, fileName As String, _
                                            fileType As String, groupName As String, _
                                            groupID As String, wsMapping As Worksheet) As Boolean
    On Error GoTo FileError
    
    Dim filePath As String
    Dim fileNum As Integer
    Dim oMapping As ValidationEngine.ColumnMapping
    
    ' Get column mapping
    oMapping = GetColumnMappingForValidation(fileType, wsMapping)
    
    If oMapping.fileType = "" Then
        Debug.Print "WARNING: No mapping found for FileType: " & fileType
        CreateValidationTestCSVFile = False
        Exit Function
    End If
    
    filePath = folderPath & "\" & fileName
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    
    ' Determine max columns needed for this FileType
    Dim maxCol As Integer
    maxCol = GetMaxColumnForFileType(oMapping)
    
    ' Write header row with correct column names
    Print #fileNum, BuildHeaderRow(oMapping, maxCol)
    
    ' Generate test records
    Print #fileNum, GenerateValidationRow(1, "VALID", groupID, groupName, oMapping, maxCol)
    
    ' Blank required fields
    Print #fileNum, GenerateValidationRow(2, "BLANK_FIRSTNAME", groupID, groupName, oMapping, maxCol)
    Print #fileNum, GenerateValidationRow(3, "BLANK_LASTNAME", groupID, groupName, oMapping, maxCol)
    Print #fileNum, GenerateValidationRow(4, "BLANK_ADDRESS1", groupID, groupName, oMapping, maxCol)
    Print #fileNum, GenerateValidationRow(5, "BLANK_CITY", groupID, groupName, oMapping, maxCol)
    Print #fileNum, GenerateValidationRow(6, "BLANK_ZIPCODE", groupID, groupName, oMapping, maxCol)
    
    ' Max length exceeded
    Print #fileNum, GenerateValidationRow(8, "MAXLEN_FIRSTNAME", groupID, groupName, oMapping, maxCol)
    Print #fileNum, GenerateValidationRow(9, "MAXLEN_LASTNAME", groupID, groupName, oMapping, maxCol)
    Print #fileNum, GenerateValidationRow(10, "MAXLEN_ADDRESS1", groupID, groupName, oMapping, maxCol)
    Print #fileNum, GenerateValidationRow(11, "MAXLEN_CITY", groupID, groupName, oMapping, maxCol)
    
    ' Invalid formats
    Print #fileNum, GenerateValidationRow(14, "INVALID_ZIPCODE", groupID, groupName, oMapping, maxCol)
    Print #fileNum, GenerateValidationRow(15, "INVALID_DOB", groupID, groupName, oMapping, maxCol)
    Print #fileNum, GenerateValidationRow(16, "INVALID_GENDER", groupID, groupName, oMapping, maxCol)
    Print #fileNum, GenerateValidationRow(17, "INVALID_STATE", groupID, groupName, oMapping, maxCol)
    
    ' Invalid characters
    Print #fileNum, GenerateValidationRow(18, "INVALID_CHARS_FIRSTNAME", groupID, groupName, oMapping, maxCol)
    Print #fileNum, GenerateValidationRow(19, "INVALID_CHARS_LASTNAME", groupID, groupName, oMapping, maxCol)
    Print #fileNum, GenerateValidationRow(20, "INVALID_CHARS_CITY", groupID, groupName, oMapping, maxCol)
    
    ' Duplicate scenarios
    Print #fileNum, GenerateValidationRow(900, "DUP_BOTH_ACTIVE_1", groupID, groupName, oMapping, maxCol)
    Print #fileNum, GenerateValidationRow(900, "DUP_BOTH_ACTIVE_2", groupID, groupName, oMapping, maxCol)
    Print #fileNum, GenerateValidationRow(901, "DUP_ONE_ACTIVE", groupID, groupName, oMapping, maxCol)
    Print #fileNum, GenerateValidationRow(901, "DUP_ONE_INACTIVE", groupID, groupName, oMapping, maxCol)
    
    ' Invalid GroupID
    Print #fileNum, GenerateValidationRow(24, "INVALID_GROUPID", groupID, groupName, oMapping, maxCol)
    
    ' Combination errors
    Print #fileNum, GenerateValidationRow(25, "COMBO_BLANK_AND_LENGTH", groupID, groupName, oMapping, maxCol)
    Print #fileNum, GenerateValidationRow(26, "COMBO_FORMAT_AND_CHARS", groupID, groupName, oMapping, maxCol)
    Print #fileNum, GenerateValidationRow(27, "COMBO_MULTIPLE_BLANKS", groupID, groupName, oMapping, maxCol)
    
    ' Edge cases
    Print #fileNum, GenerateValidationRow(28, "EDGE_ZIP_PLUS4", groupID, groupName, oMapping, maxCol)
    Print #fileNum, GenerateValidationRow(29, "EDGE_ZIP_TOO_SHORT", groupID, groupName, oMapping, maxCol)
    Print #fileNum, GenerateValidationRow(30, "EDGE_FUTURE_DATE", groupID, groupName, oMapping, maxCol)
    
    Close #fileNum
    CreateValidationTestCSVFile = True
    Exit Function
    
FileError:
    CreateValidationTestCSVFile = False
    If fileNum > 0 Then Close #fileNum
End Function

' ==============================================================================
' Build header row based on column mapping
' ==============================================================================
Private Function BuildHeaderRow(oMapping As ValidationEngine.ColumnMapping, maxCol As Integer) As String
    Dim headers() As String
    ReDim headers(1 To maxCol)
    Dim i As Integer
    
    ' Initialize all as empty
    For i = 1 To maxCol
        headers(i) = "Column" & i
    Next i
    
    ' Place headers in correct positions per mapping
    If oMapping.FirstName > 0 And oMapping.FirstName <= maxCol Then
        headers(oMapping.FirstName) = "FirstName"
    End If
    
    If oMapping.LastName > 0 And oMapping.LastName <= maxCol Then
        headers(oMapping.LastName) = "LastName"
    End If
    
    If oMapping.DOB > 0 And oMapping.DOB <= maxCol Then
        headers(oMapping.DOB) = "DateOfBirth"
    End If
    
    If oMapping.Gender > 0 And oMapping.Gender <= maxCol Then
        headers(oMapping.Gender) = "Gender"
    End If
    
    If oMapping.ZipCode > 0 And oMapping.ZipCode <= maxCol Then
        headers(oMapping.ZipCode) = "ZipCode"
    End If
    
    If oMapping.Address1 > 0 And oMapping.Address1 <= maxCol Then
        headers(oMapping.Address1) = "AddressLine1"
    End If
    
    If oMapping.City > 0 And oMapping.City <= maxCol Then
        headers(oMapping.City) = "City"
    End If
    
    If oMapping.State > 0 And oMapping.State <= maxCol Then
        headers(oMapping.State) = "State"
    End If
    
    If oMapping.EffectiveDate > 0 And oMapping.EffectiveDate <= maxCol Then
        headers(oMapping.EffectiveDate) = "EffectiveStart"
    End If
    
    If oMapping.effectiveEndDate > 0 And oMapping.effectiveEndDate <= maxCol Then
        headers(oMapping.effectiveEndDate) = "EffectiveEnd"
    End If
    
    If oMapping.groupID > 0 And oMapping.groupID <= maxCol Then
        headers(oMapping.groupID) = "GroupID"
    End If
    
    If oMapping.serviceOffering > 0 And oMapping.serviceOffering <= maxCol Then
        headers(oMapping.serviceOffering) = "ServiceOffering"
    End If
    
    If oMapping.memberID > 0 And oMapping.memberID <= maxCol Then
        headers(oMapping.memberID) = "ClientMemberID"
    End If
    
    BuildHeaderRow = Join(headers, ",")
End Function

' ==============================================================================
' Generate data row with values in CORRECT column positions
' ==============================================================================
Private Function GenerateValidationRow(recordID As Long, errorType As String, _
                                       groupID As String, groupName As String, _
                                       oMapping As ValidationEngine.ColumnMapping, _
                                       maxCol As Integer) As String
    Dim row() As String
    ReDim row(1 To maxCol)
    Dim i As Integer
    
    ' Initialize all columns as empty
    For i = 1 To maxCol
        row(i) = ""
    Next i
    
    ' Get default valid values
    Dim vFirstName As String, vLastName As String, vDOB As String
    Dim vGender As String, vZipCode As String, vAddress1 As String
    Dim vCity As String, vState As String, vEffectiveStart As String
    Dim vEffectiveEnd As String, vGroupID As String, vServiceOffering As String
    Dim vMemberID As String
    
    ' Set defaults
    vFirstName = "Firstname" & recordID
    vLastName = "Lastname" & recordID
    vDOB = "01/15/1990"
    vGender = "M"
    vZipCode = "77801"
    vAddress1 = recordID & " Main Street"
    vCity = "TestCity"
    vState = "TX"
    vEffectiveStart = Format(DateAdd("m", -6, Date), "mm/dd/yyyy")
    vEffectiveEnd = ""
    vGroupID = groupID  ' ? USE ACTUAL GroupID parameter!
    vServiceOffering = "MEDICAL"
    vMemberID = "VALID" & Format(recordID, "000000")
    
    ' Apply error scenarios
    Select Case errorType
        Case "VALID"
            ' Keep all defaults
            
        Case "BLANK_FIRSTNAME"
            vFirstName = ""
            
        Case "BLANK_LASTNAME"
            vLastName = ""
            
        Case "BLANK_ADDRESS1"
            vAddress1 = ""
            
        Case "BLANK_CITY"
            vCity = ""
            
        Case "BLANK_ZIPCODE"
            vZipCode = ""
            
        Case "MAXLEN_FIRSTNAME"
            vFirstName = String(75, "A") & "TooLong"
            
        Case "MAXLEN_LASTNAME"
            vLastName = String(75, "B") & "TooLong"
            
        Case "MAXLEN_ADDRESS1"
            vAddress1 = String(175, "X") & "TooLongAddress"
            
        Case "MAXLEN_CITY"
            vCity = String(175, "C") & "TooLongCity"
            
        Case "INVALID_ZIPCODE"
            vZipCode = "ABC12"
            
        Case "INVALID_DOB"
            vDOB = "13/32/2020"
            
        Case "INVALID_GENDER"
            vGender = "X"
            
        Case "INVALID_STATE"
            vState = "ZZ"
            
        Case "INVALID_CHARS_FIRSTNAME"
            vFirstName = "Test@Name#123$"
            
        Case "INVALID_CHARS_LASTNAME"
            vLastName = "Last&Name%^"
            
        Case "INVALID_CHARS_CITY"
            vCity = "City@#$%"
            
        Case "DUP_BOTH_ACTIVE_1", "DUP_BOTH_ACTIVE_2"
            vMemberID = "DUP900000"
            vEffectiveEnd = ""
            
        Case "DUP_ONE_ACTIVE"
            vMemberID = "DUP901000"
            vEffectiveEnd = ""
            
        Case "DUP_ONE_INACTIVE"
            vMemberID = "DUP901000"
            vEffectiveEnd = Format(DateAdd("m", -1, Date), "mm/dd/yyyy")
            
        Case "INVALID_GROUPID"
            vGroupID = "WRONGGROUP123"
            
        Case "COMBO_BLANK_AND_LENGTH"
            vLastName = ""
            vFirstName = String(75, "Z") & "TooLong"
            
        Case "COMBO_FORMAT_AND_CHARS"
            vCity = "City@#$"
            
        Case "COMBO_MULTIPLE_BLANKS"
            vLastName = ""
            vCity = ""
            
        Case "EDGE_ZIP_PLUS4"
            vZipCode = "77801-1234"
            
        Case "EDGE_ZIP_TOO_SHORT"
            vZipCode = "778"
            
        Case "EDGE_FUTURE_DATE"
            vEffectiveStart = Format(DateAdd("yyyy", 1, Date), "mm/dd/yyyy")
    End Select
    
    ' Place values in CORRECT column positions
    If oMapping.FirstName > 0 And oMapping.FirstName <= maxCol Then
        row(oMapping.FirstName) = vFirstName
    End If
    
    If oMapping.LastName > 0 And oMapping.LastName <= maxCol Then
        row(oMapping.LastName) = vLastName
    End If
    
    If oMapping.DOB > 0 And oMapping.DOB <= maxCol Then
        row(oMapping.DOB) = vDOB
    End If
    
    If oMapping.Gender > 0 And oMapping.Gender <= maxCol Then
        row(oMapping.Gender) = vGender
    End If
    
    If oMapping.ZipCode > 0 And oMapping.ZipCode <= maxCol Then
        row(oMapping.ZipCode) = vZipCode
    End If
    
    If oMapping.Address1 > 0 And oMapping.Address1 <= maxCol Then
        row(oMapping.Address1) = vAddress1
    End If
    
    If oMapping.City > 0 And oMapping.City <= maxCol Then
        row(oMapping.City) = vCity
    End If
    
    If oMapping.State > 0 And oMapping.State <= maxCol Then
        row(oMapping.State) = vState
    End If
    
    If oMapping.EffectiveDate > 0 And oMapping.EffectiveDate <= maxCol Then
        row(oMapping.EffectiveDate) = vEffectiveStart
    End If
    
    If oMapping.effectiveEndDate > 0 And oMapping.effectiveEndDate <= maxCol Then
        row(oMapping.effectiveEndDate) = vEffectiveEnd
    End If
    
    If oMapping.groupID > 0 And oMapping.groupID <= maxCol Then
        row(oMapping.groupID) = vGroupID
    End If
    
    If oMapping.serviceOffering > 0 And oMapping.serviceOffering <= maxCol Then
        row(oMapping.serviceOffering) = vServiceOffering
    End If
    
    If oMapping.memberID > 0 And oMapping.memberID <= maxCol Then
        row(oMapping.memberID) = vMemberID
    End If
    
    GenerateValidationRow = Join(row, ",")
End Function

' ==============================================================================
' Get max column number needed for this FileType
' ==============================================================================
Private Function GetMaxColumnForFileType(oMapping As ValidationEngine.ColumnMapping) As Integer
    Dim maxCol As Integer
    maxCol = 0
    
    If oMapping.FirstName > maxCol Then maxCol = oMapping.FirstName
    If oMapping.LastName > maxCol Then maxCol = oMapping.LastName
    If oMapping.DOB > maxCol Then maxCol = oMapping.DOB
    If oMapping.Gender > maxCol Then maxCol = oMapping.Gender
    If oMapping.ZipCode > maxCol Then maxCol = oMapping.ZipCode
    If oMapping.Address1 > maxCol Then maxCol = oMapping.Address1
    If oMapping.City > maxCol Then maxCol = oMapping.City
    If oMapping.State > maxCol Then maxCol = oMapping.State
    If oMapping.EffectiveDate > maxCol Then maxCol = oMapping.EffectiveDate
    If oMapping.effectiveEndDate > maxCol Then maxCol = oMapping.effectiveEndDate
    If oMapping.groupID > maxCol Then maxCol = oMapping.groupID
    If oMapping.serviceOffering > maxCol Then maxCol = oMapping.serviceOffering
    If oMapping.memberID > maxCol Then maxCol = oMapping.memberID
    
    ' Add buffer for unmapped columns
    GetMaxColumnForFileType = maxCol + 5
End Function

' ==============================================================================
' Get column mapping
' ==============================================================================
Private Function GetColumnMappingForValidation(fileType As String, _
                                              wsMapping As Worksheet) As ValidationEngine.ColumnMapping
    Dim oMapping As ValidationEngine.ColumnMapping
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = wsMapping.Cells(wsMapping.Rows.count, "A").End(xlUp).row
    oMapping.fileType = ""
    
    For i = 2 To lastRow
        If UCase(Trim(wsMapping.Cells(i, "A").Value)) = UCase(Trim(fileType)) Then
            With oMapping
                .fileType = fileType
                .FirstName = wsMapping.Cells(i, "B").Value
                .LastName = wsMapping.Cells(i, "C").Value
                .DOB = wsMapping.Cells(i, "D").Value
                .Gender = wsMapping.Cells(i, "E").Value
                .ZipCode = wsMapping.Cells(i, "F").Value
                .Address1 = wsMapping.Cells(i, "G").Value
                .City = wsMapping.Cells(i, "H").Value
                .State = wsMapping.Cells(i, "I").Value
                .EffectiveDate = wsMapping.Cells(i, "J").Value
                .groupID = wsMapping.Cells(i, "K").Value
                .serviceOffering = wsMapping.Cells(i, "L").Value
                .memberID = wsMapping.Cells(i, "M").Value
                .effectiveEndDate = wsMapping.Cells(i, "N").Value
            End With
            Exit For
        End If
    Next i
    
    GetColumnMappingForValidation = oMapping
End Function

' ==============================================================================
' Folder selection
' ==============================================================================
Private Function SelectOutputFolder() As String
    Dim folderPicker As fileDialog
    
    Set folderPicker = Application.fileDialog(msoFileDialogFolderPicker)
    
    With folderPicker
        .Title = "Select Folder for Validation Test Files"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            SelectOutputFolder = .SelectedItems(1)
        Else
            SelectOutputFolder = ""
        End If
    End With
End Function

