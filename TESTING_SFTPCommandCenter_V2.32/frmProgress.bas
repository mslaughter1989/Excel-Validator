VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgress 
   Caption         =   "File Validation Progress"
   ClientHeight    =   2415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "frmProgress.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()
    ' Modern flat design
    Me.BackColor = RGB(248, 248, 248)
    Me.BorderStyle = fmBorderStyleSingle
    
    ' Center on screen
    Me.StartUpPosition = 1
    
    ' Initialize controls
    lblMessage.Caption = "Initializing validation system..."
    lblPercent.Caption = "0%"
End Sub

Public Sub UpdateProgress(sMessage As String, lPercent As Long)
    lblMessage.Caption = sMessage
    lblPercent.Caption = lPercent & "%"
    DoEvents ' Allow screen to refresh
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Prevent user from closing during validation
    If CloseMode = 0 Then ' User clicked X button
        Cancel = True
        MsgBox "Please wait for validation to complete.", vbInformation
    End If
End Sub
