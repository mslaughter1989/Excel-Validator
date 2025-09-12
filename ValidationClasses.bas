Attribute VB_Name = "ValidationClasses"
' Class Module: ValidationResult
Option Explicit

Private m_ErrorCount As Long
Private m_WarningCount As Long
Private m_TotalRecords As Long
Private m_IsValid As Boolean
Private m_ValidationComplete As Boolean
Private m_ReportPath As String
Private m_Errors As Collection
Private m_Warnings As Collection

Private Sub Class_Initialize()
    Set m_Errors = New Collection
    Set m_Warnings = New Collection
    m_IsValid = True
End Sub

Public Property Get ErrorCount() As Long
    ErrorCount = m_ErrorCount
End Property

Public Property Let ErrorCount(Value As Long)
    m_ErrorCount = Value
End Property

Public Property Get WarningCount() As Long
    WarningCount = m_WarningCount
End Property

Public Property Let WarningCount(Value As Long)
    m_WarningCount = Value
End Property

Public Property Get TotalRecords() As Long
    TotalRecords = m_TotalRecords
End Property

Public Property Let TotalRecords(Value As Long)
    m_TotalRecords = Value
End Property

Public Property Get IsValid() As Boolean
    IsValid = m_IsValid
End Property

Public Property Let IsValid(Value As Boolean)
    m_IsValid = Value
End Property

Public Property Get ValidationComplete() As Boolean
    ValidationComplete = m_ValidationComplete
End Property

Public Property Let ValidationComplete(Value As Boolean)
    m_ValidationComplete = Value
End Property

Public Property Get ReportPath() As String
    ReportPath = m_ReportPath
End Property

Public Property Let ReportPath(Value As String)
    m_ReportPath = Value
End Property

Public Property Get Errors() As Collection
    Set Errors = m_Errors
End Property

Public Property Get Warnings() As Collection
    Set Warnings = m_Warnings
End Property

Public Sub AddError(lRowNumber As Long, sFieldName As String, sErrorMessage As String)
    Dim oError As ValidationError
    Set oError = New ValidationError
    
    With oError
        .RowNumber = lRowNumber
        .FieldName = sFieldName
        .ErrorMessage = sErrorMessage
        .ErrorType = "Error"
    End With
    
    m_Errors.Add oError
    m_ErrorCount = m_ErrorCount + 1
    m_IsValid = False
End Sub

Public Sub AddWarning(lRowNumber As Long, sFieldName As String, sWarningMessage As String)
    Dim oWarning As ValidationError
    Set oWarning = New ValidationError
    
    With oWarning
        .RowNumber = lRowNumber
        .FieldName = sFieldName
        .ErrorMessage = sWarningMessage
        .ErrorType = "Warning"
    End With
    
    m_Warnings.Add oWarning
    m_WarningCount = m_WarningCount + 1
End Sub

' Class Module: ValidationError
Option Explicit

Public RowNumber As Long
Public FieldName As String
Public ErrorMessage As String
Public ErrorType As String
