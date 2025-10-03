Attribute VB_Name = "XLSMB_EligRecap_ProcessandSave"
Sub EligRecap_FileSelector()
    Dim fileDialog As fileDialog
    Dim selectedFiles As Collection, appliedFiles As Collection, skippedFiles As Collection
    Dim FileName As String, fileBaseName As String, regex As Object, fileItem As Variant
    Dim report As String, masterWB As Workbook, masterWS As Worksheet
    Dim lastRowDest As Long, isFirstCopy As Boolean
    Dim savePath As String, timeStamp As String
    Dim checkMark As String
    Dim wb As Workbook, ws As Worksheet
    Dim i As Integer
    Dim oneDriveCommercial As String

    Set selectedFiles = New Collection
    Set appliedFiles = New Collection
    Set skippedFiles = New Collection
    Set regex = CreateObject("VBScript.RegExp")
    Set masterWB = Workbooks.Add
    Set masterWS = masterWB.Sheets(1)
    masterWS.Name = "Combined EligRecap"
    isFirstCopy = True

    checkMark = ChrW(&H2713) ' Unicode checkmark ?

    ' Timestamp and path - dynamically get current user's Downloads folder
    timeStamp = Format(Now, "yyyymmdd_HHmm")
    savePath = Environ("USERPROFILE") & "\Downloads\EligibilityRecap_CombinedResults_" & timeStamp & ".xlsx"

    ' Get OneDrive Commercial path for file selection
    oneDriveCommercial = GetOneDriveCommercialPath()
    
    ' Set up file dialog
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    With fileDialog
        .Title = "Select EligibilityRecap CSV Files"
        .AllowMultiSelect = True
        .InitialFileName = oneDriveCommercial & "\Documents - Customer Success\General\GeneratedFiles\EligibilityRecap\"
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        .FilterIndex = 1
    End With

    ' Show dialog and exit if user cancels
    If fileDialog.Show = 0 Then
        masterWB.Close SaveChanges:=False
        MsgBox "File selection cancelled.", vbInformation, "EligRecap Macro"
        Exit Sub
    End If

    ' Match pattern: EligibilityRecapYYYY_MM_DD
    With regex
        .Global = False
        .IgnoreCase = True
        .Pattern = "^EligibilityRecap\d{4}_\d{2}_\d{2}"
    End With

    ' Process selected files and filter by naming pattern
    For i = 1 To fileDialog.SelectedItems.count
        FileName = fileDialog.SelectedItems(i)
        fileBaseName = Mid(FileName, InStrRev(FileName, "\") + 1)
        
        ' Remove file extension to get base name
        If InStrRev(fileBaseName, ".") > 0 Then
            fileBaseName = Left(fileBaseName, InStrRev(fileBaseName, ".") - 1)
        End If

        If regex.Test(fileBaseName) Then
            selectedFiles.Add FileName
        Else
            skippedFiles.Add Mid(FileName, InStrRev(FileName, "\") + 1) ' Just filename for report
        End If
    Next i

    ' Check if any eligible files were found
    If selectedFiles.count = 0 Then
        masterWB.Close SaveChanges:=False
        MsgBox "No files matching the EligibilityRecapYYYY_MM_DD naming pattern were selected.", vbExclamation, "EligRecap Macro"
        Exit Sub
    End If

    ' Process each eligible file
    For Each fileItem In selectedFiles
        ' Open CSV file
        Set wb = Workbooks.Open(fileItem, ReadOnly:=True)
        Set ws = wb.ActiveSheet
        
        ' Apply filtering logic
        Call Run_EligRecap_Filter(ws)
        appliedFiles.Add Mid(fileItem, InStrRev(fileItem, "\") + 1) ' Just filename for report

        ' Copy visible filtered rows
        With ws.AutoFilter.Range
            If isFirstCopy Then
                .SpecialCells(xlCellTypeVisible).Copy ' Include headers
                isFirstCopy = False
            Else
                .Offset(1, 0).Resize(.Rows.count - 1).SpecialCells(xlCellTypeVisible).Copy ' Skip header
            End If
        End With

        ' Paste into master sheet
        With masterWS
            If Application.WorksheetFunction.CountA(.Rows(1)) = 0 Then
                lastRowDest = 1
            Else
                lastRowDest = .Cells(.Rows.count, 1).End(xlUp).Row + 1
            End If
            .Cells(lastRowDest, 1).PasteSpecial Paste:=xlPasteValues
        End With
        
        ' Close the CSV file without saving
        wb.Close SaveChanges:=False
    Next fileItem

    Application.CutCopyMode = False

    ' NEW: Merge duplicate rows (same ClientID + Errors) before sorting
    Call MergeDuplicateRows(masterWS)

    ' Sort column B alphabetically (A-Z) - Column B is "Name"
    With masterWS
        If Application.WorksheetFunction.CountA(.Rows(1)) > 0 Then
            .Range("A1").CurrentRegion.Sort Key1:=.Range("B1"), Order1:=xlAscending, Header:=xlYes
        End If
    End With

    ' Save but keep open
    masterWB.SaveAs FileName:=savePath, FileFormat:=xlOpenXMLWorkbook

    ' Build summary report
    report = "PROCESSED FILES:" & vbCrLf
    For Each fileItem In appliedFiles
        report = report & " - " & fileItem & vbCrLf
    Next fileItem

    If skippedFiles.count > 0 Then
        report = report & vbCrLf & "SKIPPED FILES (wrong naming pattern):" & vbCrLf
        For Each fileItem In skippedFiles
            report = report & " - " & fileItem & vbCrLf
        Next fileItem
    End If

    report = report & vbCrLf & checkMark & " Combined file saved to:" & vbCrLf & savePath & vbCrLf & vbCrLf & "It has been left open for your review."

    MsgBox report, vbInformation, "EligRecap Macro Report"
End Sub

Sub MergeDuplicateRows(ws As Worksheet)
    ' Merges rows with duplicate ClientID (Column A) and Errors (Column H)
    ' Combines FileName values (Column E) with semicolons
    
    Dim lastRow As Long, i As Long, j As Long
    Dim clientID1 As String, clientID2 As String
    Dim errors1 As String, errors2 As String
    Dim fileName1 As String, fileName2 As String
    Dim mergedCount As Long
    
    Application.StatusBar = "Merging duplicate rows..."
    mergedCount = 0
    
    ' Get last row with data
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    ' Start from row 2 (skip header) and work backwards to avoid index issues when deleting
    For i = lastRow To 2 Step -1
        clientID1 = Trim(ws.Cells(i, 1).Value) ' Column A
        errors1 = Trim(ws.Cells(i, 8).Value)   ' Column H (Errors after column filtering)
        
        ' Look for matching rows below current row
        For j = i + 1 To lastRow
            clientID2 = Trim(ws.Cells(j, 1).Value)
            errors2 = Trim(ws.Cells(j, 8).Value)
            
            ' Check if ClientID and Errors match
            If clientID1 = clientID2 And errors1 = errors2 Then
                ' Merge FileName values (Column E)
                fileName1 = Trim(ws.Cells(i, 5).Value)
                fileName2 = Trim(ws.Cells(j, 5).Value)
                
                ' Combine filenames with semicolon separator
                If fileName1 <> "" And fileName2 <> "" Then
                    ws.Cells(i, 5).Value = fileName1 & "; " & fileName2
                ElseIf fileName2 <> "" Then
                    ws.Cells(i, 5).Value = fileName2
                End If
                
                ' Delete the duplicate row
                ws.Rows(j).Delete
                lastRow = lastRow - 1
                j = j - 1 ' Adjust index since we deleted a row
                mergedCount = mergedCount + 1
            End If
        Next j
    Next i
    
    Application.StatusBar = False
    
    If mergedCount > 0 Then
        MsgBox "Merged " & mergedCount & " duplicate rows with matching ClientID and Errors.", _
               vbInformation, "Duplicate Merge Complete"
    End If
End Sub

' === FILTERING LOGIC (Same as original) ===
Sub Run_EligRecap_Filter(ws As Worksheet)
    Dim lastRow As Long, i As Long
    Dim cellValD As String, cellValN As String
    Dim keepRow As Boolean

    ws.AutoFilterMode = False
    ws.Rows.Hidden = False

    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If ws.Cells(ws.Rows.count, 14).End(xlUp).Row > lastRow Then
        lastRow = ws.Cells(ws.Rows.count, 14).End(xlUp).Row
    End If

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add2 key:=ws.Range("B2:B" & lastRow), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range("A1:P" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ' Filter by Column D (FileProcessStatusDesc)
    For i = 2 To lastRow
        cellValD = ws.Cells(i, 4).Value
        If cellValD <> "Completed with Errors" And cellValD <> "Failed to Process File" Then
            ws.Rows(i).Hidden = True
        End If
    Next i

    ' Then filter by Column N (Errors)
    For i = 2 To lastRow
        If ws.Rows(i).Hidden = False Then
            cellValN = ws.Cells(i, 14).Value
            keepRow = False
            If InStr(1, cellValN, "Duplicate MemberID for unique MemberID FileProcess", vbTextCompare) > 0 Then keepRow = True
            If InStr(1, cellValN, "Invalid Product Offering", vbTextCompare) > 0 Then keepRow = True
            If InStr(1, cellValN, "Invalid Group ID", vbTextCompare) > 0 Then keepRow = True
            If Trim(cellValN) = "" Then keepRow = True
            If Not keepRow Then ws.Rows(i).Hidden = True
        End If
    Next i

    ' Final clean-up - Hide columns for cleaner presentation
    ws.Rows("1:1").AutoFilter
    ws.Columns("D:D").EntireColumn.Hidden = True    ' FileProcessStatusDesc
    ws.Columns("F:F").EntireColumn.Hidden = True    ' EndDate
    ws.Columns("J:M").EntireColumn.Hidden = True    ' ErrorCount through TermedMembers
    ws.Columns("O:P").EntireColumn.Hidden = True    ' Groups and Product Offerings
End Sub

Function GetOneDriveCommercialPath() As String
    ' Function to dynamically find OneDrive Commercial path
    Dim fso As Object
    Dim userProfile As String
    Dim oneDriveFolder As Variant
    Dim folder As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    userProfile = Environ("USERPROFILE")
    
    ' Look for OneDrive - [Company Name] folders
    If fso.FolderExists(userProfile) Then
        Set folder = fso.GetFolder(userProfile)
        For Each oneDriveFolder In folder.SubFolders
            If Left(oneDriveFolder.Name, 10) = "OneDrive -" Then
                GetOneDriveCommercialPath = oneDriveFolder.Path
                Exit Function
            End If
        Next oneDriveFolder
    End If
    
    ' Fallback to regular OneDrive if commercial not found
    If fso.FolderExists(userProfile & "\OneDrive") Then
        GetOneDriveCommercialPath = userProfile & "\OneDrive"
    Else
        ' Final fallback to Downloads
        GetOneDriveCommercialPath = userProfile & "\Downloads"
    End If
    
    Set fso = Nothing
End Function

