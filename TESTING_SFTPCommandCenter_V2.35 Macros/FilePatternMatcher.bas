Attribute VB_Name = "FilePatternMatcher"
Option Explicit

Public Type FileInfo
    fileName As String
    fileType As String
    groupID As String
    fileDate As Date
    isValid As Boolean
End Type

Public Function MatchFilenamePattern(sFileName As String) As FileInfo
    Const sPROC_NAME As String = "MatchFilenamePattern"
    
    Dim oFileInfo As FileInfo
    Dim wsPatterns As Worksheet
    Dim lLastRow As Long
    Dim lRow As Long
    
    On Error GoTo ErrorHandler
    
    Set wsPatterns = ThisWorkbook.Worksheets("Parsed_SFTPfiles")
    lLastRow = wsPatterns.Cells(wsPatterns.Rows.count, "M").End(xlUp).row
    
    ' Initialize result as invalid
    oFileInfo.isValid = False
    oFileInfo.fileName = sFileName
    oFileInfo.fileType = ""
    oFileInfo.groupID = ""
    oFileInfo.fileDate = 0
    
    ' Loop through all patterns - SINGLE LOOP ONLY
    For lRow = 2 To lLastRow
        Dim sPattern As String
        Dim sGroupID As String
        Dim sFileType As String
        
        sPattern = wsPatterns.Cells(lRow, "M").Value  ' Column M
        sGroupID = wsPatterns.Cells(lRow, "K").Value  ' Column K
        sFileType = wsPatterns.Cells(lRow, "O").Value ' Column O
        
        If TestFilenameAgainstPattern(sFileName, sPattern, sGroupID) Then
            ' Add debug output to see what matched
            Debug.Print "MATCH FOUND for: " & sFileName
            Debug.Print "  Matched Pattern (Row " & lRow & "): " & sPattern
            Debug.Print "  FileType: " & sFileType
            Debug.Print "  GroupID: " & sGroupID
            
            oFileInfo.fileName = sFileName
            oFileInfo.fileType = sFileType
            oFileInfo.groupID = sGroupID
            oFileInfo.fileDate = ExtractDateFromFilename(sFileName)
            oFileInfo.isValid = True
            
            MatchFilenamePattern = oFileInfo
            Exit Function
        End If
    Next lRow
    
    ' No match found - return invalid FileInfo
    MatchFilenamePattern = oFileInfo
    Exit Function
    
ErrorHandler:
    Call ErrorHandler_Central(sPROC_NAME, Err.Number, Err.Description)
    oFileInfo.isValid = False
    MatchFilenamePattern = oFileInfo
End Function

Private Function TestFilenameAgainstPattern(sFileName As String, sPattern As String, sGroupID As String) As Boolean
    ' Convert pattern to regex
    Dim sRegexPattern As String
    sRegexPattern = ConvertPatternToRegex(sPattern, sGroupID)
    
    Dim oRegex As Object
    Set oRegex = CreateObject("VBScript.RegExp")
    
    With oRegex
        .pattern = sRegexPattern
        .IgnoreCase = True
        .Global = False
    End With
    
    Dim bResult As Boolean
    bResult = oRegex.Test(sFileName)
    
    TestFilenameAgainstPattern = bResult
End Function
Private Function ConvertPatternToRegex(sPattern As String, sGroupID As String) As String
    Dim sResult As String
    sResult = sPattern
    
    ' First escape special regex characters (but NOT backslash)
    sResult = Replace(sResult, ".", "\.")
    sResult = Replace(sResult, "+", "\+")
    sResult = Replace(sResult, "[", "\[")
    sResult = Replace(sResult, "]", "\]")
    sResult = Replace(sResult, "(", "\(")
    sResult = Replace(sResult, ")", "\)")
    sResult = Replace(sResult, "^", "\^")
    sResult = Replace(sResult, "$", "\$")
    
    ' Replace date patterns with simple digit matching
    ' Using \d{8} is simpler and avoids the complex date validation
    sResult = Replace(sResult, "mmddyyyy", "\d{8}")
    sResult = Replace(sResult, "ddmmyyyy", "\d{8}")
    sResult = Replace(sResult, "yyyymmdd", "\d{8}")
    sResult = Replace(sResult, "mmddyy", "\d{6}")
    
    ' Replace GroupID placeholder
    If sGroupID <> "" Then
        sResult = Replace(sResult, "{GroupID}", sGroupID)
        sResult = Replace(sResult, "[Adjusted groupID]", sGroupID)
    End If
    
    ' NOW replace wildcards (AFTER everything else)
    sResult = Replace(sResult, "*", ".*")
    sResult = Replace(sResult, "?", ".")
    
    ' Add anchors
    ConvertPatternToRegex = "^" & sResult & "$"
End Function

Public Function ExtractDateFromFilename(sFileName As String) As Date
    Dim oRegex As Object
    Dim oMatches As Object
    
    Set oRegex = CreateObject("VBScript.RegExp")
    With oRegex
        .Global = False
        .IgnoreCase = True
    End With
    
    ' Pattern 1: MMDDYYYY
    oRegex.pattern = "(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])(\d{4})"
    Set oMatches = oRegex.Execute(sFileName)
    If oMatches.count > 0 Then
        Dim sMonth As String, sDay As String, sYear As String
        sMonth = oMatches(0).SubMatches(0)
        sDay = oMatches(0).SubMatches(1)
        sYear = oMatches(0).SubMatches(2)
        
        On Error Resume Next
        ExtractDateFromFilename = CDate(sMonth & "/" & sDay & "/" & sYear)
        If Err.Number <> 0 Then ExtractDateFromFilename = 0
        On Error GoTo 0
        Exit Function
    End If
    
    ' Pattern 2: YYYY-MM-DD
    oRegex.pattern = "(\d{4})[-.]([01]\d)[-.]([0-3]\d)"
    Set oMatches = oRegex.Execute(sFileName)
    If oMatches.count > 0 Then
        On Error Resume Next
        ExtractDateFromFilename = CDate(oMatches(0).SubMatches(1) & "/" & _
                                      oMatches(0).SubMatches(2) & "/" & _
                                      oMatches(0).SubMatches(0))
        If Err.Number <> 0 Then ExtractDateFromFilename = 0
        On Error GoTo 0
        Exit Function
    End If
    
    ' No date found
    ExtractDateFromFilename = 0
End Function

