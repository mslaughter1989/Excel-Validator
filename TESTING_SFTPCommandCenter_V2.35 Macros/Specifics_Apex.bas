Attribute VB_Name = "Specifics_Apex"
Sub Specifics_Apex()
    Dim ws As Worksheet
    Dim rngP As Range, cell As Range
    Dim dict As Object
    Dim i As Long
    Dim smallestRow As Long
    Dim minValue As Double
    
    ' Set the active worksheet dynamically
    Set ws = ActiveSheet
    
    ' Define the range dynamically
    Set rngP = ws.Range("P2:P" & ws.Cells(ws.Rows.count, "P").End(xlUp).row)
    
    ' Step 1: Identify duplicates in Column P
    Set dict = CreateObject("Scripting.Dictionary")
    
    For Each cell In rngP
        If Not dict.Exists(cell.Value) Then
            dict.Add cell.Value, 1
        Else
            dict(cell.Value) = dict(cell.Value) + 1
        End If
    Next cell
    
    ' Step 2: Delete rows where Column P is a duplicate AND Column N has data
    For i = ws.Cells(ws.Rows.count, "P").End(xlUp).row To 2 Step -1
        If dict.Exists(ws.Cells(i, "P").Value) And dict(ws.Cells(i, "P").Value) > 1 And ws.Cells(i, "N").Value <> "" Then
            ws.Rows(i).Delete
        End If
    Next i
    
    ' Step 3: Identify remaining duplicates and delete the row with the smallest value in Column M
    Set dict = CreateObject("Scripting.Dictionary") ' Reinitialize dictionary for remaining duplicates
    
    For Each cell In ws.Range("P2:P" & ws.Cells(ws.Rows.count, "P").End(xlUp).row)
        If Not dict.Exists(cell.Value) Then
            dict.Add cell.Value, cell.row
        Else
            ' Compare Column M values for duplicates
            If ws.Cells(dict(cell.Value), "M").Value < ws.Cells(cell.row, "M").Value Then
                ws.Rows(dict(cell.Value)).Delete
            Else
                ws.Rows(cell.row).Delete
            End If
        End If
    Next cell
    
    ' Step 4: Format Column I as a ZIP code
    ws.Columns("I").NumberFormat = "00000" ' Ensures 5-digit ZIP code format
    
End Sub


