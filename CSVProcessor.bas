Attribute VB_Name = "CSVProcessor"
Option Explicit

Public Function ReadCSVToArray(sFilePath As String) As Variant
    Const sPROC_NAME As String = "ReadCSVToArray"
    
    On Error GoTo ErrorHandler
    
    ' ADD THESE DEBUG LINES HERE
    Debug.Print "=== ReadCSVToArray Debug ==="
    Debug.Print "File path: " & sFilePath
    Debug.Print "File exists: " & (Dir(sFilePath) <> "")
    Dim lFileNum As Long
    Dim sFileContent As String
    Dim vLines As Variant
    Dim vData As Variant
    Dim lRow As Long
    Dim vFields As Variant
    
    ' Check if file exists
    If Dir(sFilePath) = "" Then
        ReadCSVToArray = Empty
        Exit Function
    End If
    
    ' Read entire file at once for better performance
    lFileNum = FreeFile
    Open sFilePath For Input As #lFileNum
    sFileContent = Input$(LOF(lFileNum), lFileNum)
    Close #lFileNum
    
    ' Handle different line endings
    sFileContent = Replace(sFileContent, vbCrLf, vbLf)
    sFileContent = Replace(sFileContent, vbCr, vbLf)
    
    ' Split into lines
    vLines = Split(sFileContent, vbLf)
    
    ' Count non-empty lines
    Dim lValidRows As Long
    For lRow = 0 To UBound(vLines)
        If Trim(vLines(lRow)) <> "" Then
            lValidRows = lValidRows + 1
        End If
    Next lRow
    
    If lValidRows = 0 Then
        ReadCSVToArray = Empty
        Exit Function
    End If
    
    ' Process first line to determine column count
    Dim lColCount As Long
    vFields = ParseCSVLine(CStr(vLines(0)))
    lColCount = UBound(vFields) + 1
    
    ' Create 2D array
    ReDim vData(1 To lValidRows, 1 To lColCount)
    
    Dim lDataRow As Long
    lDataRow = 1
    
    ' Parse all lines
    For lRow = 0 To UBound(vLines)
        If Trim(vLines(lRow)) <> "" Then
            vFields = ParseCSVLine(CStr(vLines(lRow)))
            
            Dim lCol As Long
            For lCol = 0 To UBound(vFields)
                If lCol + 1 <= lColCount Then
                    vData(lDataRow, lCol + 1) = vFields(lCol)
                End If
            Next lCol
            
            lDataRow = lDataRow + 1
        End If
    Next lRow
    
    ReadCSVToArray = vData
    Exit Function
    
ErrorHandler:
    If lFileNum > 0 Then Close #lFileNum
    Call ErrorHandler_Central(sPROC_NAME, Err.Number, Err.Description, sFilePath)
    ReadCSVToArray = Empty
End Function

Public Function ParseCSVLine(sLine As String) As Variant
    Dim vFields() As String
    Dim lFieldCount As Long
    Dim lPos As Long
    Dim sCurrentField As String
    Dim bInQuotes As Boolean
    Dim sChar As String
    Dim lIndex As Long
    
    ReDim vFields(0 To 100) ' Initial size
    lFieldCount = 0
    lPos = 1
    bInQuotes = False
    sCurrentField = ""
    
    Do While lPos <= Len(sLine)
        sChar = Mid(sLine, lPos, 1)
        
        Select Case sChar
            Case """"
                If bInQuotes Then
                    ' Check for escaped quote
                    If lPos < Len(sLine) And Mid(sLine, lPos + 1, 1) = """" Then
                        sCurrentField = sCurrentField & """"
                        lPos = lPos + 1 ' Skip next quote
                    Else
                        bInQuotes = False
                    End If
                Else
                    bInQuotes = True
                End If
                
            Case ","
                If Not bInQuotes Then
                    ' End of field
                    vFields(lFieldCount) = sCurrentField
                    lFieldCount = lFieldCount + 1
                    sCurrentField = ""
                    
                    ' Expand array if needed
                    If lFieldCount > UBound(vFields) Then
                        ReDim Preserve vFields(0 To UBound(vFields) + 50)
                    End If
                Else
                    sCurrentField = sCurrentField & sChar
                End If
                
            Case Else
                sCurrentField = sCurrentField & sChar
        End Select
        
        lPos = lPos + 1
    Loop
    
    ' Add final field
    vFields(lFieldCount) = sCurrentField
    
    ' Resize array to exact size
    ReDim Preserve vFields(0 To lFieldCount)
    
    ParseCSVLine = vFields
End Function
