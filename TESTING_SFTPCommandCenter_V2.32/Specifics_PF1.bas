Attribute VB_Name = "Specifics_PF1"
Option Explicit

'===========================  MAIN UNIFIED PROCEDURE  ===========================
Sub Specifics_PF1()
    'This macro handles all CSV processing in one comprehensive workflow:
    '1. Updates product codes based on Group ID business rules
    '2. Filters out unwanted rows based on Group ID and inactive date criteria
    '3. Applies formatting to ZIP codes
    'Run this macro on the currently active CSV file

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim compCol As Long, prodCol As Long, zipCol As Long, groupCol As Long, inactiveDateCol As Long
    Dim lastRow As Long, r As Long, key As String
    Dim currentDate As Date
    Dim updatedProductCodes As Long, deletedRows As Long
    
    'Performance optimization - turn off screen updating during processing
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    'Get current date for date comparisons (without time component)
    currentDate = Date
    
    'Create mapping dictionary for company name to product code (original functionality)
    Dim companyMap As Object: Set companyMap = CreateObject("Scripting.Dictionary")
    companyMap(CleanStr("SOLIDCORE HOLDINGS LLC")) = "39658"
    companyMap(CleanStr("GEORGETOWN HILL CHILD CARE CENTER INC")) = "33212"
    companyMap(CleanStr("EASY ICE LLC")) = "33212"
    companyMap(CleanStr("BOOMTOWN NETWORK INC")) = "33212"
    
    'Create mapping for Group ID to Product Code replacements (new functionality)
    Dim groupIDMap As Object: Set groupIDMap = CreateObject("Scripting.Dictionary")
    groupIDMap.Add 728072, "39658"   'Group ID 728072 gets Product Code 39658
    groupIDMap.Add 801910, "33212"   'These three Group IDs all get Product Code 33212
    groupIDMap.Add 816941, "33212"
    groupIDMap.Add 816859, "33212"
    
    'Define the Group IDs that need special date filtering
    Dim filterGroupIDs As Variant
    filterGroupIDs = Array(662424, 735759, 601946, 732421, 662427, 600029, 578320, 752737, 603573, 660877, 656835, 623646)
    
    'Work with the currently active workbook
    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(1)           'First (only) sheet in the CSV
    updatedProductCodes = 0
    deletedRows = 0
    
    'STEP 1: IDENTIFY COLUMN POSITIONS
    'Find all the columns we need to work with
    compCol = HeaderCol(ws.Rows(1), "Company Name")
    prodCol = HeaderCol(ws.Rows(1), "Product Code")
    zipCol = HeaderCol(ws.Rows(1), "Zip Code")
    groupCol = HeaderCol(ws.Rows(1), "Group Id")        'Column A
    inactiveDateCol = HeaderCol(ws.Rows(1), "inactive date")  'Column E
    
    'Verify we found the essential columns
    If prodCol = 0 Or groupCol = 0 Then
        MsgBox "Couldn't find required columns (Product Code or Group Id) in " & wb.Name, vbExclamation
        GoTo CleanupAndExit
    End If
    
    'STEP 2: UPDATE PRODUCT CODES BASED ON BUSINESS RULES
    'This happens first, before we delete any rows
    lastRow = ws.Cells(ws.Rows.count, groupCol).End(xlUp).Row
    
    For r = 2 To lastRow                       'Skip header row
        Dim currentGroupID As Variant
        currentGroupID = ws.Cells(r, groupCol).Value
        
        'Handle company name mapping (original functionality)
        If compCol > 0 Then
            key = CleanStr(ws.Cells(r, compCol).Value)
            If companyMap.Exists(key) Then
                ws.Cells(r, prodCol).Value = companyMap(key)
                updatedProductCodes = updatedProductCodes + 1
            End If
        End If
        
        'Handle Group ID to Product Code mapping (new functionality)
        If IsNumeric(currentGroupID) Then
            Dim groupIDLong As Long
            groupIDLong = CLng(currentGroupID)
            
            If groupIDMap.Exists(groupIDLong) Then
                ws.Cells(r, prodCol).Value = groupIDMap(groupIDLong)
                updatedProductCodes = updatedProductCodes + 1
            End If
        End If
    Next r
    
    'STEP 3: FILTER OUT UNWANTED ROWS
    'Work backwards through rows to avoid index shifting issues when deleting
    'This is critical - deleting rows changes the row numbers of all rows below
    For r = lastRow To 2 Step -1              'Skip header row, work backwards
        currentGroupID = ws.Cells(r, groupCol).Value
        
        'Check if this row has one of the Group IDs that need date filtering
        If IsTargetGroupID(currentGroupID, filterGroupIDs) Then
            Dim inactiveDate As Variant
            Dim shouldDelete As Boolean
            shouldDelete = False
            
            'Get the inactive date value
            If inactiveDateCol > 0 Then
                inactiveDate = ws.Cells(r, inactiveDateCol).Value
            Else
                inactiveDate = ""  'Column not found, treat as empty
            End If
            
            'Apply deletion logic: remove rows where inactive date is NOT prior to current date
            If IsEmpty(inactiveDate) Or inactiveDate = "" Then
                'Empty/null dates should be deleted (not prior to current date)
                shouldDelete = True
            ElseIf IsDate(inactiveDate) Then
                'Valid date: delete if NOT prior to current date
                If CDate(inactiveDate) >= currentDate Then
                    shouldDelete = True
                End If
                'If date is prior to current date, keep the row
            Else
                'Invalid date format: delete the row
                shouldDelete = True
            End If
            
            'Delete the entire row if criteria met
            If shouldDelete Then
                ws.Rows(r).EntireRow.Delete
                deletedRows = deletedRows + 1
            End If
        End If
    Next r
    
    'STEP 4: FORMAT ZIP CODE COLUMN
    'Ensure ZIP codes display with leading zeros
    If zipCol > 0 Then ws.Columns(zipCol).NumberFormat = "00000"
    
    'Provide feedback on what was accomplished
    Dim message As String
    message = "Processing complete for " & wb.Name & vbCrLf & _
             "• Product codes updated: " & updatedProductCodes & vbCrLf & _
             "• Rows deleted: " & deletedRows
    MsgBox message, vbInformation
    
CleanupAndExit:
    'Restore Excel performance settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
End Sub

'===========================  HELPER FUNCTIONS  ===========================

'Returns the 1-based column number whose cleaned-up header matches targetHeader
Private Function HeaderCol(headerRow As Range, targetHeader As String) As Long
    Dim c As Range
    For Each c In headerRow.Cells
        If Len(c.Value) = 0 Then Exit For                 'Stop after trailing blanks
        If CleanStr(c.Value) = CleanStr(targetHeader) Then
            HeaderCol = c.Column: Exit Function
        End If
    Next c
    HeaderCol = 0                                         'Not found
End Function

'Normalizes strings: uppercase, strip spaces, commas, tabs, non-breaking spaces & line-breaks
Private Function CleanStr(s As String) As String
    Dim t As String
    t = UCase$(s)
    t = Replace(t, ",", "")
    t = Replace(t, Chr(160), "")      'Non-breaking space
    t = Replace(t, " ", "")
    t = Replace(t, vbTab, "")
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")
    CleanStr = Trim$(t)
End Function

'Checks if a Group ID matches any ID in the target array for filtering
Private Function IsTargetGroupID(GroupID As Variant, targetArray As Variant) As Boolean
    Dim i As Integer
    
    'Handle empty or non-numeric values
    If IsEmpty(GroupID) Or Not IsNumeric(GroupID) Then
        IsTargetGroupID = False
        Exit Function
    End If
    
    'Convert to Long for comparison to handle various numeric formats
    Dim groupIDLong As Long
    groupIDLong = CLng(GroupID)
    
    'Check against each target ID in the array
    For i = LBound(targetArray) To UBound(targetArray)
        If groupIDLong = targetArray(i) Then
            IsTargetGroupID = True
            Exit Function
        End If
    Next i
    
    IsTargetGroupID = False
End Function

