Attribute VB_Name = "Specifics_PrismHR"
Sub Specifics_PrismHR()
    Dim wb As Workbook, ws As Worksheet
    Dim parsedWS As Worksheet
    Dim fileName As String, fmt As String, testPattern As String
    Dim expectedGroupID As String
    Dim regex As Object
    Dim rowCount As Long, lastCol As Long, lastRow As Long, colPEO As Long
    Dim i As Long
    Dim matchFound As Boolean

    Set wb = ActiveWorkbook
    Set ws = wb.Sheets(1)
    fileName = wb.Name
    matchFound = False

    ' Use internal worksheet instead of external file
    On Error GoTo ErrorHandler
    Set parsedWS = ThisWorkbook.Worksheets("Parsed_SFTPFiles")
    rowCount = parsedWS.Cells(parsedWS.Rows.count, 1).End(xlUp).row
    On Error GoTo 0

    Set regex = CreateObject("VBScript.RegExp")
    regex.IgnoreCase = True
    regex.Global = False

    ' Match filename to patterns in the internal worksheet
    For i = 2 To rowCount
        fmt = Trim(parsedWS.Cells(i, 1).Value)  ' Column A: Initial Filename Format
        expectedGroupID = Trim(parsedWS.Cells(i, 11).Value)  ' Column K: Group ID
        
        ' Skip empty rows
        If fmt = "" Or expectedGroupID = "" Then GoTo NextFormat

        ' Build regex pattern from filename format
        testPattern = fmt
        testPattern = Replace(testPattern, ".", "\.")
        testPattern = Replace(testPattern, "-", "\-")
        testPattern = Replace(testPattern, "_", "_")
        testPattern = Replace(testPattern, "mmddyyyy", "\d{8}")
        testPattern = Replace(testPattern, "yyyymmdd", "\d{8}")
        testPattern = Replace(testPattern, "mmddyy", "\d{6}")

        regex.pattern = "^" & testPattern & "$"
        If regex.Test(fileName) Then
            matchFound = True
            Exit For
        End If
NextFormat:
    Next i

    If Not matchFound Or expectedGroupID = "" Then
        MsgBox "? No matching filename pattern found for: " & fileName, vbCritical
        Exit Sub
    End If

    ' Find dimensions of the active worksheet
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row

    ' Locate "PEO ID" column
    colPEO = 0
    For i = 1 To lastCol
        If Trim(ws.Cells(1, i).Value) = "PEO ID" Then
            colPEO = i
            Exit For
        End If
    Next i

    If colPEO = 0 Then
        MsgBox "? 'PEO ID' column not found in the file.", vbCritical
        Exit Sub
    End If

    ' Update all PEO ID values to match the expected Group ID
    Dim updatedCount As Long
    updatedCount = 0
    For i = 2 To lastRow
        If Trim(ws.Cells(i, colPEO).Value) <> expectedGroupID Then
            ws.Cells(i, colPEO).Value = expectedGroupID
            updatedCount = updatedCount + 1
        End If
    Next i

    Exit Sub

ErrorHandler:
    MsgBox "? Error accessing 'Parsed_SFTPFiles' worksheet: " & Err.Description & vbCrLf & _
           "Please ensure the worksheet exists in the current workbook.", vbCritical
End Sub

