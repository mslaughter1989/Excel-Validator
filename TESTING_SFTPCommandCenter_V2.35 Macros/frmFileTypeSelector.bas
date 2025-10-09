VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFileTypeSelector 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmFileTypeSelector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFileTypeSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' This code goes in the frmFileTypeSelector UserForm module
Option Explicit

Private m_SelectedFileType As String
Private m_Cancelled As Boolean

' Property to get the selected file type
Public Property Get SelectedFileType() As String
    SelectedFileType = m_SelectedFileType
End Property

' Property to check if user cancelled
Public Property Get Cancelled() As Boolean
    Cancelled = m_Cancelled
End Property

' Initialize the form with file types
Public Sub InitializeForm(colFileTypes As Collection)
    Dim vFileType As Variant
    
    ' Clear any existing items
    cboFileType.Clear
    
    ' Add all file types to the dropdown
    For Each vFileType In colFileTypes
        cboFileType.AddItem CStr(vFileType)
    Next vFileType
    
    ' Select the first item by default
    If cboFileType.ListCount > 0 Then
        cboFileType.ListIndex = 0
    End If
    
    ' Set form properties
    Me.Caption = "Select File Type"
    lblPrompt.Caption = "Please select the file type for validation:"
    btnOK.Caption = "OK"
    btnCancel.Caption = "Cancel"
    
    ' Initialize variables
    m_SelectedFileType = ""
    m_Cancelled = True
End Sub

Private Sub UserForm_Initialize()
    ' Set form size and position
    Me.Width = 350
    Me.Height = 150
    
    ' Position controls
    With lblPrompt
        .Top = 12
        .Left = 12
        .Width = 300
        .Height = 20
        .Font.Size = 10
    End With
    
    With cboFileType
        .Top = 40
        .Left = 12
        .Width = 300
        .Height = 24
        .Style = fmStyleDropDownList ' Forces dropdown list (no typing)
        .Font.Size = 10
    End With
    
    With btnOK
        .Top = 80
        .Left = 150
        .Width = 75
        .Height = 25
        .Default = True ' Make it the default button (Enter key)
        .Font.Size = 9
    End With
    
    With btnCancel
        .Top = 80
        .Left = 235
        .Width = 75
        .Height = 25
        .Cancel = True ' Make it respond to Escape key
        .Font.Size = 9
    End With
End Sub

Private Sub btnOK_Click()
    ' Validate selection
    If cboFileType.ListIndex = -1 Then
        MsgBox "Please select a file type from the dropdown.", vbExclamation, "No Selection"
        Exit Sub
    End If
    
    ' Store the selection
    m_SelectedFileType = cboFileType.Value
    m_Cancelled = False
    
    ' Hide the form (don't unload yet)
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    ' Set cancelled flag
    m_Cancelled = True
    m_SelectedFileType = ""
    
    ' Hide the form
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' If user clicks the X button, treat it as Cancel
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        btnCancel_Click
    End If
End Sub
