Attribute VB_Name = "ValidatorModule"


Sub Validate_WithRequiredColumnsSplitLog(Optional baseFileName As String)
    If baseFileName = "" Then baseFileName = "UnnamedFile"

    Dim ws As Worksheet, resultWB As Workbook, logSheetRequired As Worksheet, logSheetAll As Worksheet
    Dim lastRow As Long, lastCol As Long, colHeader As String
    Dim cellValue As String, col As Long, row As Long
    Dim checkDict As Object, requiredCols As Object
    Dim logRowReq As Long, logRowAll As Long
    Dim headerRow As Range
    Dim logFolderPath As String

    Set ws = ActiveSheet
    Set headerRow = ws.Range("1:1")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Set checkDict = CreateObject("Scripting.Dictionary")
    Set requiredCols = CreateObject("Scripting.Dictionary")

    ' Add required columns
    requiredCols.Add "First Name", True
    requiredCols.Add "Last Name", True
    requiredCols.Add "Address Line 1", True
    requiredCols.Add "City", True
    requiredCols.Add "Zip Code", True
    requiredCols.Add "E-mail Address", True

    ' Add validation rules: column name -> array of (check type, pattern/limit)
    checkDict.Add "First Name", Array("Name Format", 50)
    checkDict.Add "Last Name", Array("Name Format", 50)
    checkDict.Add "Address Line 1", Array("Address Format", 150)
    checkDict.Add "Address Line 2", Array("Max Length Only", 150)
    checkDict.Add "City", Array("Alphanumeric+Space", 150)
    checkDict.Add "Zip Code", Array("Zip Format", 10)
    checkDict.Add "E-mail Address", Array("Email Format", 150)

    ' Create results workbook and two sheets
    Set resultWB = Workbooks.Add
    Set logSheetRequired = resultWB.Sheets(1)
    logSheetRequired.Name = "Required Columns Log"
    Set logSheetAll = resultWB.Sheets.Add(After:=logSheetRequired)
    logSheetAll.Name = "All Columns Log"

    logSheetRequired.Range("A1:D1").Value = Array("Row", "Column", "Value", "Issue")
    logSheetAll.Range("A1:D1").Value = Array("Row", "Column", "Value", "Issue")
    logRowReq = 2
    logRowAll = 2

    For col = 1 To lastCol
        colHeader = Trim(ws.Cells(1, col).Value)
        If checkDict.exists(colHeader) Then
            For row = 2 To lastRow
                cellValue = Trim(ws.Cells(row, col).Value)
                Dim issue As String: issue = ""

                ' Required column blank check
                If requiredCols.exists(colHeader) And cellValue = "" Then
                    issue = "Blank required field"
                    logSheetRequired.Cells(logRowReq, 1).Resize(1, 4).Value = Array(row, colHeader, cellValue, issue)
                    logRowReq = logRowReq + 1
                End If

                ' Check value according to rule
                If cellValue <> "" Then
                    Dim checkType As String
                    Dim checkLimit As Long
                    checkType = checkDict(colHeader)(0)
                    checkLimit = checkDict(colHeader)(1)

                    Select Case checkType
                        Case "Name Format"
                            If Len(cellValue) > checkLimit Then issue = "Exceeds " & checkLimit & " characters"
                            If Not cellValue Like WorksheetFunction.Rept("[A-Za-z0-9 '-]", Len(cellValue)) Then
                                If issue <> "" Then issue = issue & "; " Else issue = ""
                                issue = issue & "Invalid characters"
                            End If

                        Case "Address Format"
                            If Len(cellValue) > checkLimit Then issue = "Exceeds " & checkLimit & " characters"
                            If Not cellValue Like WorksheetFunction.Rept("[A-Za-z0-9 .,-]", Len(cellValue)) Then
                                If issue <> "" Then issue = issue & "; " Else issue = ""
                                issue = issue & "Invalid characters"
                            End If

                        Case "Max Length Only"
                            If Len(cellValue) > checkLimit Then issue = "Exceeds " & checkLimit & " characters"

                        Case "Alphanumeric+Space"
                            If Len(cellValue) > checkLimit Then issue = "Exceeds " & checkLimit & " characters"
                            If Not cellValue Like WorksheetFunction.Rept("[A-Za-z0-9 ]", Len(cellValue)) Then
                                If issue <> "" Then issue = issue & "; " Else issue = ""
                                issue = issue & "Invalid characters"
                            End If

                        Case "Zip Format"
                            If Len(cellValue) > checkLimit Then issue = "Exceeds " & checkLimit & " characters"
                            If Not cellValue Like "#####*" Then
                                If issue <> "" Then issue = issue & "; " Else issue = ""
                                issue = issue & "Invalid zip code"
                            End If

                        Case "Email Format"
                            If Len(cellValue) > checkLimit Then issue = "Exceeds " & checkLimit & " characters"
                            If InStr(1, cellValue, "@") = 0 Or InStr(1, cellValue, ".") = 0 Then
                                If issue <> "" Then issue = issue & "; " Else issue = ""
                                issue = issue & "Invalid email format"
                            End If
                    End Select

                    If issue <> "" Then
                        If requiredCols.exists(colHeader) Then
                            logSheetRequired.Cells(logRowReq, 1).Resize(1, 4).Value = Array(row, colHeader, cellValue, issue)
                            logRowReq = logRowReq + 1
                        End If
                        logSheetAll.Cells(logRowAll, 1).Resize(1, 4).Value = Array(row, colHeader, cellValue, issue)
                        logRowAll = logRowAll + 1
                    End If
                End If
            Next row
        End If
    Next col

    ' Save results
    logFolderPath = ThisWorkbook.Path & "\\Logs\\"
    resultWB.SaveAs Filename:=logFolderPath & baseFileName & "_ValidationLog.xlsx", FileFormat:=xlOpenXMLWorkbook
    resultWB.Activate
    MsgBox "Validation complete. Log saved to: " & logFolderPath, vbInformation
End Sub
