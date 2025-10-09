Attribute VB_Name = "Specifics_Oglethorpe"
Sub Specifics_Oglethorpe()
    Dim ws As Worksheet
    Dim headerRow As Range
    Dim cell As Range
    Dim locationCol As Integer
    Dim lastRow As Long
    Dim i As Long
    Dim originalVal As String
    Dim cleanedVal As String
    Dim numericVal As Variant
    Dim regex As Object

    Set ws = ActiveSheet
    Set headerRow = ws.Rows(1)
    locationCol = 0

    ' Find the "Location" column
    For Each cell In headerRow.Cells
        If Trim(cell.Value) = "Location" Then
            locationCol = cell.Column
            Exit For
        End If
    Next cell

    If locationCol = 0 Then
        MsgBox "Location column not found.", vbExclamation
        Exit Sub
    End If

    ' Create Regex to remove non-alphanumeric characters
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = "[^A-Za-z0-9]"
    regex.Global = True

    ' Find last used row
    lastRow = ws.Cells(ws.Rows.count, locationCol).End(xlUp).row

    ' Process each cell in the Location column
    For i = 2 To lastRow
        originalVal = ws.Cells(i, locationCol).Text
        cleanedVal = regex.Replace(originalVal, "")
        
        If IsNumeric(cleanedVal) Then
            ws.Cells(i, locationCol).Value = Val(cleanedVal)
        Else
            ws.Cells(i, locationCol).Value = ""
        End If
    Next i

    ' Apply custom format #000 to the Location column
    ws.Columns(locationCol).NumberFormat = "#000"

End Sub

