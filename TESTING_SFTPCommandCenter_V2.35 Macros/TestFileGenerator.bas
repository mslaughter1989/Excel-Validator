Attribute VB_Name = "TestFileGenerator"

Option Explicit

' ==============================================================================
' TEST FILE GENERATOR - FIXED VERSION v2.0
' Generates sample CSV files with CORRECT column order per FileType
' Includes duplicate test scenarios for Part1_Process testing
' ==============================================================================

Sub GenerateTestFiles()
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
    testFolder = fso.BuildPath(outputFolder, "Test_Files_" & Format(Now, "yyyymmdd_hhnnss"))
    If Not fso.FolderExists(testFolder) Then
        fso.CreateFolder testFolder
    End If
    
    filesCreated = 0
    reportMsg = "TEST FILES GENERATED:" & vbCrLf & vbCrLf
    
    lastRow = wsParsed.Cells(wsParsed.Rows.count, "A").End(xlUp).row
    
    For i = 2 To lastRow
        sFileType = Trim(CStr(wsParsed.Cells(i, "O").Value))
        sGroupName = Trim(CStr(wsParsed.Cells(i, "J").Value))
        sGroupID = Trim(CStr(wsParsed.Cells(i, "K").Value))
        
        If sFileType <> "" Then
            sFileName = GenerateTestFileName(wsParsed, i)
            
            If CreateTestCSVFile(testFolder, sFileName, sFileType, sGroupName, sGroupID, wsMapping) Then
                filesCreated = filesCreated + 1
                reportMsg = reportMsg & sFileName & " (" & sFileType & ")" & vbCrLf
            End If
        End If
    Next i
    
    reportMsg = reportMsg & vbCrLf & "Total files created: " & filesCreated & vbCrLf
    reportMsg = reportMsg & "Location: " & testFolder & vbCrLf & vbCrLf
    reportMsg = reportMsg & "Each file contains 12 records (6 will become duplicates)." & vbCrLf
    reportMsg = reportMsg & "After Part1_Process, should have 6 records remaining."
    
    MsgBox reportMsg, vbInformation, "Test File Generation Complete"
    
    Shell "explorer.exe " & Chr(34) & testFolder & Chr(34), vbNormalFocus
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating test files: " & Err.Description, vbCritical
End Sub

' ==============================================================================
' Generate filename based on pattern
' ==============================================================================
Private Function GenerateTestFileName(wsParsed As Worksheet, rowNum As Long) As String
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
    
    GenerateTestFileName = pattern
End Function

' ==============================================================================
' Create CSV file with CORRECT column mapping per FileType
' ==============================================================================
Private Function CreateTestCSVFile(folderPath As String, fileName As String, _
                                   fileType As String, groupName As String, _
                                   groupID As String, wsMapping As Worksheet) As Boolean
    On Error GoTo FileError
    
    Dim filePath As String
    Dim fileNum As Integer
    Dim oMapping As ValidationEngine.ColumnMapping
    Dim i As Long
    
    ' Get column mapping for this FileType
    oMapping = GetColumnMappingForFileType(fileType, wsMapping)
    
    If oMapping.fileType = "" Then
        Debug.Print "WARNING: No mapping found for FileType: " & fileType
        CreateTestCSVFile = False
        Exit Function
    End If
    
    filePath = folderPath & "\" & fileName
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    
    ' Determine max columns needed
    Dim maxCol As Integer
    maxCol = GetMaxColumnForFileType(oMapping)
    
    ' Write header row
    Print #fileNum, BuildHeaderRow(oMapping, maxCol)
    
    ' Generate test data with duplicate scenarios
    ' Each file will have 12 records that should reduce to 6 after processing
    
    ' Clean records (3 records - no duplicates)
    For i = 1 To 3
        Print #fileNum, GenerateTestDataRow(i, "CLEAN", groupID, groupName, oMapping, maxCol)
    Next i
    
    ' Duplicate Scenario 1: Active + Inactive (should keep active, remove inactive)
    Print #fileNum, GenerateTestDataRow(100, "DUP_ACTIVE", groupID, groupName, oMapping, maxCol)
    Print #fileNum, GenerateTestDataRow(100, "DUP_INACTIVE", groupID, groupName, oMapping, maxCol)
    
    ' Duplicate Scenario 2: Old + New start dates (should keep newer)
    Print #fileNum, GenerateTestDataRow(200, "DUP_OLDER", groupID, groupName, oMapping, maxCol)
    Print #fileNum, GenerateTestDataRow(200, "DUP_NEWER", groupID, groupName, oMapping, maxCol)
    
    ' Duplicate Scenario 3: Three versions (should keep newest)
    Print #fileNum, GenerateTestDataRow(300, "DUP_OLDEST", groupID, groupName, oMapping, maxCol)
    Print #fileNum, GenerateTestDataRow(300, "DUP_MIDDLE", groupID, groupName, oMapping, maxCol)
    Print #fileNum, GenerateTestDataRow(300, "DUP_NEWEST", groupID, groupName, oMapping, maxCol)
    
    Close #fileNum
    CreateTestCSVFile = True
    Exit Function
    
FileError:
    CreateTestCSVFile = False
    If fileNum > 0 Then Close #fileNum
End Function

' ==============================================================================
' Build header row with correct column positions
' ==============================================================================
Private Function BuildHeaderRow(oMapping As ValidationEngine.ColumnMapping, maxCol As Integer) As String
    Dim headers() As String
    ReDim headers(1 To maxCol)
    Dim i As Integer
    
    ' Initialize all columns
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
' Generate test data row with values in CORRECT column positions
' ==============================================================================
Private Function GenerateTestDataRow(recordID As Long, scenario As String, _
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
    
    ' Prepare values based on scenario
    Dim vMemberID As String
    Dim vEffectiveStart As String
    Dim vEffectiveEnd As String
    
    ' Base member ID on record ID
    vMemberID = "TEST" & Format(recordID, "000000")
    
    ' Set dates based on scenario
    Select Case scenario
        Case "CLEAN"
            vEffectiveStart = Format(DateAdd("m", -6, Date), "mm/dd/yyyy")
            vEffectiveEnd = ""
            
        Case "DUP_ACTIVE"
            vEffectiveStart = Format(DateAdd("m", -3, Date), "mm/dd/yyyy")
            vEffectiveEnd = ""
            
        Case "DUP_INACTIVE"
            vEffectiveStart = Format(DateAdd("m", -6, Date), "mm/dd/yyyy")
            vEffectiveEnd = Format(DateAdd("m", -1, Date), "mm/dd/yyyy")
            
        Case "DUP_OLDER"
            vEffectiveStart = Format(DateAdd("m", -12, Date), "mm/dd/yyyy")
            vEffectiveEnd = ""
            
        Case "DUP_NEWER"
            vEffectiveStart = Format(DateAdd("m", -3, Date), "mm/dd/yyyy")
            vEffectiveEnd = ""
            
        Case "DUP_OLDEST"
            vEffectiveStart = Format(DateAdd("m", -24, Date), "mm/dd/yyyy")
            vEffectiveEnd = ""
            
        Case "DUP_MIDDLE"
            vEffectiveStart = Format(DateAdd("m", -12, Date), "mm/dd/yyyy")
            vEffectiveEnd = ""
            
        Case "DUP_NEWEST"
            vEffectiveStart = Format(DateAdd("m", -3, Date), "mm/dd/yyyy")
            vEffectiveEnd = ""
    End Select
    
    ' Place values in CORRECT column positions based on mapping
    If oMapping.FirstName > 0 And oMapping.FirstName <= maxCol Then
        row(oMapping.FirstName) = "Firstname" & recordID
    End If
    
    If oMapping.LastName > 0 And oMapping.LastName <= maxCol Then
        row(oMapping.LastName) = "Lastname" & recordID
    End If
    
    If oMapping.Gender > 0 And oMapping.Gender <= maxCol Then
        row(oMapping.Gender) = IIf(recordID Mod 2 = 0, "M", "F")
    End If
    
    If oMapping.DOB > 0 And oMapping.DOB <= maxCol Then
        row(oMapping.DOB) = Format(DateAdd("yyyy", -30, Date), "mm/dd/yyyy")
    End If
    
    If oMapping.Address1 > 0 And oMapping.Address1 <= maxCol Then
        row(oMapping.Address1) = recordID & " Main Street"
    End If
    
    If oMapping.City > 0 And oMapping.City <= maxCol Then
        row(oMapping.City) = "TestCity"
    End If
    
    If oMapping.State > 0 And oMapping.State <= maxCol Then
        row(oMapping.State) = "TX"
    End If
    
    If oMapping.ZipCode > 0 And oMapping.ZipCode <= maxCol Then
        row(oMapping.ZipCode) = Format(77800 + (recordID Mod 100), "00000")
    End If
    
    If oMapping.EffectiveDate > 0 And oMapping.EffectiveDate <= maxCol Then
        row(oMapping.EffectiveDate) = vEffectiveStart
    End If
    
    If oMapping.effectiveEndDate > 0 And oMapping.effectiveEndDate <= maxCol Then
        row(oMapping.effectiveEndDate) = vEffectiveEnd
    End If
    
    If oMapping.memberID > 0 And oMapping.memberID <= maxCol Then
        row(oMapping.memberID) = vMemberID
    End If
    
    If oMapping.serviceOffering > 0 And oMapping.serviceOffering <= maxCol Then
        row(oMapping.serviceOffering) = "MEDICAL"
    End If
    
    If oMapping.groupID > 0 And oMapping.groupID <= maxCol Then
        row(oMapping.groupID) = groupID  ' ? USE ACTUAL GroupID!
    End If
    
    GenerateTestDataRow = Join(row, ",")
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
' Get column mapping for a FileType
' ==============================================================================
Private Function GetColumnMappingForFileType(fileType As String, _
                                            wsMapping As Worksheet) As ValidationEngine.ColumnMapping
    Dim oMapping As ValidationEngine.ColumnMapping
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = wsMapping.Cells(wsMapping.Rows.count, "A").End(xlUp).row
    
    ' Initialize
    oMapping.fileType = ""
    
    ' Find the FileType
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
    
    GetColumnMappingForFileType = oMapping
End Function

' ==============================================================================
' Folder selection dialog
' ==============================================================================
Private Function SelectOutputFolder() As String
    Dim folderPicker As fileDialog
    
    Set folderPicker = Application.fileDialog(msoFileDialogFolderPicker)
    
    With folderPicker
        .Title = "Select Folder for Test Files"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            SelectOutputFolder = .SelectedItems(1)
        Else
            SelectOutputFolder = ""
        End If
    End With
End Function

' ==============================================================================
' OPTIONAL: Generate test files for specific FileTypes only
' ==============================================================================
Sub GenerateTestFilesForSpecificTypes()
    Dim fileTypes As String
    Dim typeArray() As String
    Dim i As Long
    
    ' Prompt user for FileTypes
    fileTypes = InputBox("Enter FileTypes to generate (comma-separated):" & vbCrLf & _
                         "Example: FileType1,FileType2,FileType3" & vbCrLf & vbCrLf & _
                         "Or leave blank to generate all FileTypes", _
                         "Selective Test File Generation")
    
    If fileTypes = "" Then
        ' Generate all
        GenerateTestFiles
    Else
        ' Parse and generate specific ones
        typeArray = Split(fileTypes, ",")
        ' TODO: Implement selective generation
        MsgBox "Selective generation not yet implemented. Generating all files.", vbInformation
        GenerateTestFiles
    End If
End Sub

