Attribute VB_Name = "Specifics_Focus"
Sub Specifics_Focus()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row

    Dim i As Long
    Dim inactDate As Variant
    Dim ZipCode As String

    For i = 2 To lastRow 'Assuming row 1 is headers
        ' (1) Clear "Inactive Date" (Column E) if year >= 2100
        inactDate = ws.Cells(i, 5).Value
        If IsDate(inactDate) Then
            If Year(inactDate) >= 2100 Then
                ws.Cells(i, 5).ClearContents
            End If
        End If

        ' (2) Replace Group ID 663411 with 665034 (Column A)
        If ws.Cells(i, 1).Value = 663411 Then
            ws.Cells(i, 1).Value = 665034
        End If

        ' (3) Zip Code (Column O) - ensure only 5 digits
        ZipCode = ws.Cells(i, 15).Text 'Column O is 15th
        ZipCode = Replace(ZipCode, " ", "") 'remove spaces
        ZipCode = Replace(ZipCode, "-", "") 'remove dashes
        ZipCode = Replace(ZipCode, ".", "") 'remove periods
        ZipCode = Replace(ZipCode, ",", "") 'remove commas

        ' Remove all non-numeric characters
        Dim j As Integer, ch As String, result As String
        result = ""
        For j = 1 To Len(ZipCode)
            ch = Mid(ZipCode, j, 1)
            If ch Like "#" Then result = result & ch
        Next j

        ' If more than 5 digits, take the first 5; if less, pad with leading zeroes
        If Len(result) > 5 Then
            result = Left(result, 5)
        ElseIf Len(result) < 5 Then
            result = Right("00000" & result, 5)
        End If

        ws.Cells(i, 15).Value = result
    Next i

End Sub
