Attribute VB_Name = "Specifics_Mightywell"
Sub Specifics_Mightywell()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim zipCol As Long
    Dim effectiveEndCol As Long
    Dim i As Long
    Dim j As Long
    Dim zipValue As String
    Dim dateValue As String
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Find the last row and column with data
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    
    ' Initialize variables
    zipCol = 0
    effectiveEndCol = 0
    
    ' Find ZipCode column (check for variations)
    For j = 1 To lastCol
        Dim headerText As String
        headerText = UCase(Replace(Replace(ws.Cells(1, j).Value, " ", ""), "_", ""))
        
        If headerText = "ZIPCODE" Or headerText = "ZIP" Or headerText = "POSTALCODE" Or _
           headerText = "ZIPCD" Or headerText = "ZIPCDE" Then
            zipCol = j
            Exit For
        End If
    Next j
    
    ' Find EffectiveEnd column (check for variations)
    For j = 1 To lastCol
        headerText = UCase(Replace(Replace(ws.Cells(1, j).Value, " ", ""), "_", ""))
        
        If InStr(headerText, "EFFECTIVEEND") > 0 Or InStr(headerText, "ENDDATE") > 0 Or _
           InStr(headerText, "EXPIREDATE") > 0 Or InStr(headerText, "EXPIRATIONDATE") > 0 Or _
           headerText = "EFFEND" Or headerText = "ENDDT" Then
            effectiveEndCol = j
            Exit For
        End If
    Next j
    

    
    ' Process ZipCode column if found
    If zipCol > 0 Then
        For i = 2 To lastRow ' Start from row 2 to skip header
            zipValue = Trim(CStr(ws.Cells(i, zipCol).Value))
            
            ' Remove any non-numeric characters
            Dim cleanZip As String
            cleanZip = ""
            Dim k As Integer
            For k = 1 To Len(zipValue)
                If IsNumeric(Mid(zipValue, k, 1)) Then
                    cleanZip = cleanZip & Mid(zipValue, k, 1)
                End If
            Next k
            
            ' Format as 5-digit zip code
            If Len(cleanZip) > 0 Then
                If Len(cleanZip) < 5 Then
                    ' Pad with leading zeros
                    cleanZip = Right("00000" & cleanZip, 5)
                ElseIf Len(cleanZip) > 5 Then
                    ' Take first 5 digits
                    cleanZip = Left(cleanZip, 5)
                End If
                
                ' Update the cell with formatted zip code
                ws.Cells(i, zipCol).Value = cleanZip
                ws.Cells(i, zipCol).NumberFormat = "00000"
            End If
        Next i
    End If
    
    ' Process EffectiveEnd column if found - clear cells with 2999 dates
    Dim clearedCells As Integer
    clearedCells = 0
    
    If effectiveEndCol > 0 Then
        For i = 2 To lastRow ' Start from row 2 to skip header
            dateValue = CStr(ws.Cells(i, effectiveEndCol).Value)
            
            ' Check if the date contains year 2999
            If InStr(dateValue, "2999") > 0 Then
                ' Clear the cell content
                ws.Cells(i, effectiveEndCol).ClearContents
                clearedCells = clearedCells + 1
            End If
        Next i
    End If
    

    
    
End Sub

