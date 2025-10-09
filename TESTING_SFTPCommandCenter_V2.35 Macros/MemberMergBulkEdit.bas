Attribute VB_Name = "MemberMergBulkEdit"
Sub MemberMergBulkEdit()
    Dim wsSrc As Worksheet, wsOut As Worksheet
    Dim lastCol As Long, lastRow As Long
    Dim colIssueKey As Long, colKeepID As Long, colIDs As Long
    Dim i As Long, j As Long, k As Long, maxIDs As Long
    Dim matches As Object, regex As Object, regNum As Object
    Dim cellVal As String
    Dim keepID As String
    Dim csvFilePath As String
    Dim wb As Workbook
    Dim originalWb As Workbook
    
    ' Store reference to original workbook
    Set originalWb = ActiveWorkbook
    
    ' File selection dialog
    csvFilePath = Application.GetOpenFilename( _
        FileFilter:="CSV Files (*.csv), *.csv", _
        Title:="Select CSV File to Process", _
        MultiSelect:=False)
    
    ' Check if user cancelled
    If csvFilePath = "False" Then
        MsgBox "No file selected. Operation cancelled.", vbInformation
        Exit Sub
    End If
    
    ' Open the CSV file
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    On Error GoTo ErrorHandler
    Set wb = Workbooks.Open(csvFilePath)
    Set wsSrc = wb.Sheets(1) ' Use first sheet of opened CSV
    
    lastCol = wsSrc.Cells(1, wsSrc.Columns.count).End(xlToLeft).Column
    lastRow = wsSrc.Cells(wsSrc.Rows.count, 1).End(xlUp).row

    ' Identify required columns
    For i = 1 To lastCol
        If Trim(wsSrc.Cells(1, i).Value) = "Issue key" Then colIssueKey = i
        If wsSrc.Cells(1, i).Value Like "Custom field (*Member ID to Keep Active (If known)*)" Then colKeepID = i
        If wsSrc.Cells(1, i).Value Like "Custom field (*Member ID(s)*)" Then colIDs = i
    Next i

    If colIssueKey = 0 Or colKeepID = 0 Or colIDs = 0 Then
        MsgBox "One or more required columns were not found in the CSV file." & vbNewLine & _
               "Required columns:" & vbNewLine & _
               "- Issue key" & vbNewLine & _
               "- Custom field (*Member ID to Keep Active (If known)*)" & vbNewLine & _
               "- Custom field (*Member ID(s)*)", vbExclamation
        GoTo Cleanup
    End If

    ' Create or clear output sheet in the original workbook
    originalWb.Activate
    On Error Resume Next
    originalWb.Worksheets("Processed Member IDs").Delete
    On Error GoTo ErrorHandler
    Set wsOut = originalWb.Worksheets.Add
    wsOut.Name = "Processed Member IDs"

    ' Set headers
    wsOut.Cells(1, 1).Value = "Issue key"
    wsOut.Cells(1, 2).Value = "Member ID to Keep Active (If known)"
    wsOut.Cells(1, 3).Value = "Member ID(s)"

    ' RegEx setup
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.pattern = "\b\d{9}\b"

    Set regNum = CreateObject("VBScript.RegExp")
    regNum.Global = False
    regNum.pattern = "\d{9}"

    ' Determine max number of member IDs
    maxIDs = 0
    For i = 2 To lastRow
        cellVal = wsSrc.Cells(i, colIDs).Value
        If regex.Test(cellVal) Then
            Set matches = regex.Execute(cellVal)
            If matches.count > maxIDs Then maxIDs = matches.count
        End If
    Next i

    ' Create dynamic headers
    For k = 1 To maxIDs
        wsOut.Cells(1, 3 + k).Value = "Member ID " & k
    Next k

    ' Extract and format data
    For i = 2 To lastRow
        wsOut.Cells(i, 1).Value = wsSrc.Cells(i, colIssueKey).Value

        ' Clean the Keep Active ID
        cellVal = wsSrc.Cells(i, colKeepID).Value
        If regNum.Test(cellVal) Then
            keepID = regNum.Execute(cellVal)(0)
            wsOut.Cells(i, 2).Value = keepID
        Else
            keepID = ""
            wsOut.Cells(i, 2).Value = ""
        End If

        ' Copy raw Member ID(s)
        wsOut.Cells(i, 3).Value = wsSrc.Cells(i, colIDs).Value

        ' Extract individual IDs and compare
        cellVal = wsSrc.Cells(i, colIDs).Value
        If regex.Test(cellVal) Then
            Set matches = regex.Execute(cellVal)
            For j = 0 To matches.count - 1
                wsOut.Cells(i, 4 + j).Value = matches(j)
                If matches(j) = keepID Then
                    wsOut.Cells(i, 4 + j).Interior.Color = RGB(255, 255, 0) ' Highlight yellow
                End If
            Next j
        End If
    Next i

    MsgBox "CSV file processed successfully!" & vbNewLine & _
           "Data cleaned, member IDs extracted, and matching cells highlighted." & vbNewLine & _
           "Results saved in 'Processed Member IDs' worksheet.", vbInformation

Cleanup:
    ' Close the CSV file without saving
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    GoTo Cleanup
End Sub

