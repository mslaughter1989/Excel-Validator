VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Private variables to store validation results
Private m_ErrorCount As Long
Private m_WarningCount As Long
Private m_TotalRecords As Long
Private m_IsValid As Boolean
Private m_ValidationComplete As Boolean
Private m_ReportPath As String
Private m_Errors As Collection
Private m_Warnings As Collection
Private m_FileName As String
Private m_FilePath As String
Private m_ProcessedDate As Date
Private m_FileType As String
Private m_GroupID As String
Private m_GroupName As String
Private m_ValidationChecks As Collection

' Add these property procedures with the other properties:

' Property: FileName - Get/Set the name of the file being validated
Public Property Get FileName() As String
    FileName = m_FileName
End Property

Public Property Let FileName(Value As String)
    m_FileName = Value
End Property

' Property: FilePath - Get/Set the full path of the file
Public Property Get FilePath() As String
    FilePath = m_FilePath
End Property

Public Property Let FilePath(Value As String)
    m_FilePath = Value
End Property

' Property: ProcessedDate - Get/Set when the file was processed
Public Property Get ProcessedDate() As Date
    ProcessedDate = m_ProcessedDate
End Property

Public Property Let ProcessedDate(Value As Date)
    m_ProcessedDate = Value
End Property
Public Property Get FileType() As String
    FileType = m_FileType
End Property

Public Property Let FileType(Value As String)
    m_FileType = Value
End Property

Public Property Get GroupID() As String
    GroupID = m_GroupID
End Property

Public Property Let GroupID(Value As String)
    m_GroupID = Value
End Property

Public Property Get GroupName() As String
    GroupName = m_GroupName
End Property

Public Property Let GroupName(Value As String)
    m_GroupName = Value
End Property

Public Property Get ValidationChecks() As Collection
    Set ValidationChecks = m_ValidationChecks
End Property

' Initialize the class when it's created
Private Sub Class_Initialize()
    Set m_Errors = New Collection
    Set m_Warnings = New Collection
    m_IsValid = True
    m_ErrorCount = 0
    m_WarningCount = 0
    m_TotalRecords = 0
    m_ValidationComplete = False
    m_ReportPath = ""
    m_FileName = ""
    m_FilePath = ""
    m_ProcessedDate = Now
    m_FileType = ""
    m_GroupID = ""
    m_GroupName = ""
    Set m_ValidationChecks = New Collection
End Sub
' Add this new method to record validation checks
Public Sub AddValidationCheck(sCheckName As String, sResult As String)
    Dim sCheck As String
    sCheck = sCheckName & ": " & sResult
    m_ValidationChecks.Add sCheck
End Sub

' Clean up when class is destroyed
Private Sub Class_Terminate()
    Set m_Errors = Nothing
    Set m_Warnings = Nothing
End Sub

' Property: ErrorCount - Get/Set number of errors
Public Property Get ErrorCount() As Long
    ErrorCount = m_ErrorCount
End Property

Public Property Let ErrorCount(Value As Long)
    m_ErrorCount = Value
    If Value > 0 Then m_IsValid = False
End Property

' Property: WarningCount - Get/Set number of warnings
Public Property Get WarningCount() As Long
    WarningCount = m_WarningCount
End Property

Public Property Let WarningCount(Value As Long)
    m_WarningCount = Value
End Property

' Property: TotalRecords - Get/Set total number of records processed
Public Property Get TotalRecords() As Long
    TotalRecords = m_TotalRecords
End Property

Public Property Let TotalRecords(Value As Long)
    m_TotalRecords = Value
End Property

' Property: IsValid - Get/Set whether validation passed
Public Property Get IsValid() As Boolean
    IsValid = m_IsValid And (m_ErrorCount = 0)
End Property

Public Property Let IsValid(Value As Boolean)
    m_IsValid = Value
End Property

' Property: ValidationComplete - Get/Set completion status
Public Property Get ValidationComplete() As Boolean
    ValidationComplete = m_ValidationComplete
End Property

Public Property Let ValidationComplete(Value As Boolean)
    m_ValidationComplete = Value
End Property

' Property: ReportPath - Get/Set path to validation report
Public Property Get ReportPath() As String
    ReportPath = m_ReportPath
End Property

Public Property Let ReportPath(Value As String)
    m_ReportPath = Value
End Property

' Property: Errors - Get collection of error objects
Public Property Get Errors() As Collection
    Set Errors = m_Errors
End Property

' Property: Warnings - Get collection of warning objects
Public Property Get Warnings() As Collection
    Set Warnings = m_Warnings
End Property

' Method: AddError - Add a new error to the collection
Public Sub AddError(lRowNumber As Long, sFieldName As String, sErrorMessage As String)
    Dim oError As ValidationError
    Set oError = New ValidationError
    
    With oError
        .RowNumber = lRowNumber
        .fieldName = sFieldName
        .ErrorMessage = sErrorMessage
        .ErrorType = "Error"
        .Severity = "High"
    End With
    
    m_Errors.Add oError
    m_ErrorCount = m_ErrorCount + 1
    m_IsValid = False
End Sub

' Method: AddWarning - Add a new warning to the collection
Public Sub AddWarning(lRowNumber As Long, sFieldName As String, sWarningMessage As String)
    Dim oWarning As ValidationError
    Set oWarning = New ValidationError
    
    With oWarning
        .RowNumber = lRowNumber
        .fieldName = sFieldName
        .ErrorMessage = sWarningMessage
        .ErrorType = "Warning"
        .Severity = "Low"
    End With
    
    m_Warnings.Add oWarning
    m_WarningCount = m_WarningCount + 1
End Sub

' Method: GetSummary - Return a summary string of validation results
Public Function GetSummary() As String
    Dim sSummary As String
    
    sSummary = "Validation Summary:" & vbCrLf
    sSummary = sSummary & "Total Records: " & m_TotalRecords & vbCrLf
    sSummary = sSummary & "Errors: " & m_ErrorCount & vbCrLf
    sSummary = sSummary & "Warnings: " & m_WarningCount & vbCrLf
    sSummary = sSummary & "Status: " & IIf(m_IsValid, "PASSED", "FAILED")
    
    GetSummary = sSummary
End Function

' Method: ClearResults - Reset all validation results
Public Sub ClearResults()
    Set m_Errors = New Collection
    Set m_Warnings = New Collection
    m_ErrorCount = 0
    m_WarningCount = 0
    m_IsValid = True
    m_ValidationComplete = False
    m_ReportPath = ""
End Sub
