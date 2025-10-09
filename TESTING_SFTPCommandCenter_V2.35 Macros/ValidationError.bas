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

Public RowNumber As Long
Public fieldName As String
Public ErrorMessage As String
Public errorType As String          ' "Error" or "Warning"
Public Severity As String           ' "High", "Medium", "Low"
Public ErrorCode As String          ' Optional error code
Public ActualValue As String        ' The value that caused the error
Public ExpectedFormat As String     ' What was expected
Public ValidationRule As String     ' Which rule failed

Private Sub Class_Initialize()
    RowNumber = 0
    fieldName = ""
    ErrorMessage = ""
    errorType = "Error"
    Severity = "Medium"
    ErrorCode = ""
    ActualValue = ""
    ExpectedFormat = ""
    ValidationRule = ""
End Sub

Public Function GetFormattedError() As String
    Dim sFormatted As String
    
    sFormatted = "[" & errorType & "] "
    
    If RowNumber > 0 Then
        sFormatted = sFormatted & "Row " & RowNumber & " - "
    End If
    
    If fieldName <> "" Then
        sFormatted = sFormatted & "Field '" & fieldName & "': "
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

Public Function GetCSVFormat() As String
    Dim sCSV As String
    
    sCSV = RowNumber & "," & _
           """" & fieldName & """," & _
           """" & errorType & """," & _
           """" & Severity & """," & _
           """" & Replace(ErrorMessage, """", """""") & """," & _
           """" & Replace(ActualValue, """", """""") & """," & _
           """" & Replace(ExpectedFormat, """", """""") & """"
    
    GetCSVFormat = sCSV
End Function

Public Function IsHighPriority() As Boolean
    IsHighPriority = (UCase(Severity) = "HIGH" Or UCase(errorType) = "ERROR")
End Function

Public Function GetSeverityLevel() As Long
    Select Case UCase(Severity)
        Case "HIGH"
            GetSeverityLevel = 3
        Case "MEDIUM"
            GetSeverityLevel = 2
        Case "LOW"
            GetSeverityLevel = 1
        Case Else
            GetSeverityLevel = 2
    End Select
    
    If UCase(errorType) = "ERROR" Then
        GetSeverityLevel = GetSeverityLevel + 10
    End If
End Function

Public Sub SetErrorDetails(lRow As Long, sField As String, sMessage As String, _
                           Optional sType As String = "Error", _
                           Optional sSeverity As String = "Medium", _
                           Optional sActual As String = "", _
                           Optional sExpected As String = "")
    RowNumber = lRow
    fieldName = sField
    ErrorMessage = sMessage
    errorType = sType
    Severity = sSeverity
    ActualValue = sActual
    ExpectedFormat = sExpected
End Sub

