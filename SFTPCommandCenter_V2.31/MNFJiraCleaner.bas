Attribute VB_Name = "MNFJiraCleaner"
Sub MNFJiraCleaner()
    Dim wsSrc As Worksheet, wsOut As Worksheet
    Dim lastCol As Long, lastRow As Long
    Dim headers() As String, cleanedHeaders() As String
    Dim i As Long, j As Integer
    Dim colMap As Object, colCounts As Object, colUsed As Object
    Set colMap = CreateObject("Scripting.Dictionary")
    Set colCounts = CreateObject("Scripting.Dictionary")
    Set colUsed = CreateObject("Scripting.Dictionary")

    Dim desired As Variant
    desired = Array("Issue key", "Status", "Created", "First Name", "Last Name", _
                    "Date of Birth", "Member ID", "Group Name", "Client Name", _
                    "Issue", "Service Status")

    Set wsSrc = ActiveSheet
    lastCol = wsSrc.Cells(1, wsSrc.Columns.count).End(xlToLeft).Column
    lastRow = wsSrc.Cells(wsSrc.Rows.count, 1).End(xlUp).Row

    ' Clean and map headers with duplicates
    ReDim headers(1 To lastCol)
    ReDim cleanedHeaders(1 To lastCol)

    For i = 1 To lastCol
        headers(i) = wsSrc.Cells(1, i).Value
        If headers(i) Like "Custom field (*)" Then
            cleanedHeaders(i) = Mid(headers(i), InStr(headers(i), "(") + 1, InStr(headers(i), ")") - InStr(headers(i), "(") - 1)
        Else
            cleanedHeaders(i) = headers(i)
        End If

        If Not colCounts.Exists(cleanedHeaders(i)) Then
            Set colCounts(cleanedHeaders(i)) = New Collection
        End If
        colCounts(cleanedHeaders(i)).Add i
    Next i

    ' Choose column with most non-empty cells among duplicates
    For Each key In colCounts
        Dim maxCount As Long: maxCount = 0
        Dim bestCol As Long: bestCol = 0

        For Each idx In colCounts(key)
            Dim count As Long
            count = Application.WorksheetFunction.CountA(wsSrc.Range(wsSrc.Cells(2, idx), wsSrc.Cells(lastRow, idx)))
            If count > maxCount Then
                maxCount = count
                bestCol = idx
            End If
        Next idx

        If bestCol > 0 Then colMap(key) = bestCol
    Next key

    ' Create new worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("Final Cleaned Jira").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsOut = Worksheets.Add
    wsOut.Name = "Final Cleaned Jira"

    ' Write headers and values in order
    Dim colIndex As Integer: colIndex = 1
    For i = 0 To UBound(desired)
        If colMap.Exists(desired(i)) Then
            wsOut.Cells(1, colIndex).Value = desired(i)
            wsSrc.Range(wsSrc.Cells(2, colMap(desired(i))), wsSrc.Cells(lastRow, colMap(desired(i)))).Copy
            wsOut.Cells(2, colIndex).PasteSpecial Paste:=xlPasteValues
            colIndex = colIndex + 1
        End If
    Next i

    ' Format Dates
    On Error Resume Next
    wsOut.Range("C2:C" & lastRow).NumberFormat = "mm/dd/yyyy"
    wsOut.Range("F2:F" & lastRow).NumberFormat = "mm/dd/yyyy"
    On Error GoTo 0

    ' Clean Group Name
    With wsOut
        For j = 1 To .Cells(1, .Columns.count).End(xlToLeft).Column
            If .Cells(1, j).Value = "Group Name" Then
                For i = 2 To lastRow
                    If InStr(.Cells(i, j).Value, "Group: ") > 0 Then
                        .Cells(i, j).Value = Mid(.Cells(i, j).Value, InStr(.Cells(i, j).Value, "Group: ") + 7)
                        If InStr(.Cells(i, j).Value, " ") > 0 Then
                            .Cells(i, j).Value = Mid(.Cells(i, j).Value, InStr(.Cells(i, j).Value, " ") + 1)
                        End If
                    End If
                Next i
                Exit For
            End If
        Next j
    End With

    ' Clean Client Name
    With wsOut
        For j = 1 To .Cells(1, .Columns.count).End(xlToLeft).Column
            If .Cells(1, j).Value = "Client Name" Then
                For i = 2 To lastRow
                    If InStr(.Cells(i, j).Value, "Client: ") > 0 Then
                        .Cells(i, j).Value = Mid(.Cells(i, j).Value, InStr(.Cells(i, j).Value, "Client: ") + 8)
                    End If
                    If InStr(LCase(.Cells(i, j).Value), "group:") > 0 Then
                        .Cells(i, j).Value = Left(.Cells(i, j).Value, InStr(LCase(.Cells(i, j).Value), "group:") - 1)
                    End If
                Next i
                Exit For
            End If
        Next j
    End With

    MsgBox "Jira export cleaned. Most complete versions of duplicate columns retained!", vbInformation
End Sub


