VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Public properties to store error information
Public RowNumber As Long
Public FieldName As String
Public ErrorMessage As String
Public ErrorType As String          ' "Error" or "Warning"
Public Severity As String           ' "High", "Medium", "Low"
Public ErrorCode As String          ' Optional error code
Public ActualValue As String        ' The value that caused the error
Public ExpectedFormat As String     ' What was expected
Public ValidationRule As String     ' Which rule failed

' Initialize with default values
Private Sub Class_Initialize()
    RowNumber = 0
    FieldName = ""
    ErrorMessage = ""
    ErrorType = "Error"
    Severity = "Medium"
    ErrorCode = ""
    ActualValue = ""
    ExpectedFormat = ""
    ValidationRule = ""
End Sub

' Method: GetFormattedError - Return formatted error message
Public Function GetFormattedError() As String
    Dim sFormatted As String
    
    sFormatted = "[" & ErrorType & "] "
    
    If RowNumber > 0 Then
        sFormatted = sFormatted & "Row " & RowNumber & " - "
    End If
    
    If FieldName <> "" Then
        sFormatted = sFormatted & "Field '" & FieldName & "': "
    End If
    
    sFormatted = sFormatted & ErrorMessage
    
    If ActualValue <> "" Then
        sFormatted = sFormatted & " (Value: '" & ActualValue & "')"
    End If
    
    If ExpectedFormat <> "" Then
        sFormatted = sFormatted & " (Expected: " & ExpectedFormat & ")"
    End If
    
    GetFormattedError = sFormatted
End Function

' Method: GetCSVFormat - Return error in CSV format for reporting
Public Function GetCSVFormat() As String
    Dim sCSV As String
    
    ' Format: RowNumber,FieldName,ErrorType,Severity,ErrorMessage,ActualValue,ExpectedFormat
    sCSV = RowNumber & "," & _
           """" & FieldName & """," & _
           """" & ErrorType & """," & _
           """" & Severity & """," & _
           """" & Replace(ErrorMessage, """", """""") & """," & _
           """" & Replace(ActualValue, """", """""") & """," & _
           """" & Replace(ExpectedFormat, """", """""") & """"
    
    GetCSVFormat = sCSV
End Function

' Method: IsHighPriority - Check if this is a high priority error
Public Function IsHighPriority() As Boolean
    IsHighPriority = (UCase(Severity) = "HIGH" Or UCase(ErrorType) = "ERROR")
End Function

' Method: GetSeverityLevel - Return numeric severity (for sorting)
Public Function GetSeverityLevel() As Long
    Select Case UCase(Severity)
        Case "HIGH"
            GetSeverityLevel = 3
        Case "MEDIUM"
            GetSeverityLevel = 2
        Case "LOW"
            GetSeverityLevel = 1
        Case Else
            GetSeverityLevel = 2 ' Default to medium
    End Select
    
    ' Errors are always higher priority than warnings
    If UCase(ErrorType) = "ERROR" Then
        GetSeverityLevel = GetSeverityLevel + 10
    End If
End Function

' Method: SetErrorDetails - Convenient way to set all error details at once
Public Sub SetErrorDetails(lRow As Long, sField As String, sMessage As String, _
                           Optional sType As String = "Error", _
                           Optional sSeverity As String = "Medium", _
                           Optional sActual As String = "", _
                           Optional sExpected As String = "")
    RowNumber = lRow
    FieldName = sField
    ErrorMessage = sMessage
    ErrorType = sType
    Severity = sSeverity
    ActualValue = sActual
    ExpectedFormat = sExpected
End Sub

