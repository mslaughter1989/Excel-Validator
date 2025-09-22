Attribute VB_Name = "FilePatternMatcher"
Option Explicit

Public Type FileInfo
    FileName As String
    FileType As String
    GroupID As String
    FileDate As Date
    IsValid As Boolean
End Type

Public Function MatchFilenamePattern(sFileName As String) As FileInfo
    Const sPROC_NAME As String = "MatchFilenamePattern"
    
    Dim oFileInfo As FileInfo
    Dim wsPatterns As Worksheet
    Dim lLastRow As Long
    Dim lRow As Long
    
    On Error GoTo ErrorHandler
    
    Set wsPatterns = ThisWorkbook.Worksheets("Parsed_SFTPfiles")
    lLastRow = wsPatterns.Cells(wsPatterns.Rows.Count, "M").End(xlUp).row
    
    ' Initialize result as invalid
    oFileInfo.IsValid = False
    oFileInfo.FileName = sFileName
    oFileInfo.FileType = ""
    oFileInfo.GroupID = ""
    oFileInfo.FileDate = 0
    
    ' Loop through all patterns
    For lRow = 2 To lLastRow
        Dim sPattern As String
        Dim sGroupID As String
        Dim sFileType As String
        
        sPattern = wsPatterns.Cells(lRow, "M").Value  ' Column M
        sGroupID = wsPatterns.Cells(lRow, "K").Value  ' Column K
        sFileType = wsPatterns.Cells(lRow, "O").Value ' Column O
        
        If TestFilenameAgainstPattern(sFileName, sPattern, sGroupID) Then
            oFileInfo.FileName = sFileName
            oFileInfo.FileType = sFileType
            oFileInfo.GroupID = sGroupID
            oFileInfo.FileDate = ExtractDateFromFilename(sFileName)
            oFileInfo.IsValid = True
            
            ' Return the Type directly (no Set keyword needed)
            MatchFilenamePattern = oFileInfo
            Exit Function
        End If
    Next lRow
    
    ' No match found - return invalid FileInfo
    MatchFilenamePattern = oFileInfo
    Exit Function
    
ErrorHandler:
    Call ErrorHandler_Central(sPROC_NAME, err.Number, err.description)
    oFileInfo.IsValid = False
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
    
    TestFilenameAgainstPattern = oRegex.Test(sFileName)
End Function

Private Function ConvertPatternToRegex(sPattern As String, sGroupID As String) As String
    Dim sResult As String
    sResult = sPattern
    
    ' Replace wildcards with regex equivalents
    sResult = Replace(sResult, "*", ".*")
    sResult = Replace(sResult, "?", ".")
    
    ' Replace date patterns
    sResult = Replace(sResult, "mmddyyyy", "(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])(\d{4})")
    sResult = Replace(sResult, "ddmmyyyy", "(0[1-9]|[12]\d|3[01])(0[1-9]|1[0-2])(\d{4})")
    sResult = Replace(sResult, "yyyymmdd", "(\d{4})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])")
    
    ' Replace GroupID placeholder
    sResult = Replace(sResult, "{GroupID}", sGroupID)
    
    ' Escape special regex characters
    sResult = Replace(sResult, ".", "\.")
    sResult = Replace(sResult, "(", "\(")
    sResult = Replace(sResult, ")", "\)")
    
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
    If oMatches.Count > 0 Then
        Dim sMonth As String, sDay As String, sYear As String
        sMonth = oMatches(0).SubMatches(0)
        sDay = oMatches(0).SubMatches(1)
        sYear = oMatches(0).SubMatches(2)
        
        On Error Resume Next
        ExtractDateFromFilename = CDate(sMonth & "/" & sDay & "/" & sYear)
        If err.Number <> 0 Then ExtractDateFromFilename = 0
        On Error GoTo 0
        Exit Function
    End If
    
    ' Pattern 2: YYYY-MM-DD
    oRegex.pattern = "(\d{4})[-.]([01]\d)[-.]([0-3]\d)"
    Set oMatches = oRegex.Execute(sFileName)
    If oMatches.Count > 0 Then
        On Error Resume Next
        ExtractDateFromFilename = CDate(oMatches(0).SubMatches(1) & "/" & _
                                      oMatches(0).SubMatches(2) & "/" & _
                                      oMatches(0).SubMatches(0))
        If err.Number <> 0 Then ExtractDateFromFilename = 0
        On Error GoTo 0
        Exit Function
    End If
    
    ' No date found
    ExtractDateFromFilename = 0
End Function

