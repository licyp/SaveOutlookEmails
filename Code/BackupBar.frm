VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BackupBar 
   Caption         =   "Back Up Outlook"
   ClientHeight    =   1425
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4755
   OleObjectBlob   =   "BackupBar.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BackupBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Cancel_Click()
    EndCode = 1
    Call CloseBackupProgressBar
End Sub
Private Sub Userform_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        EndCode = 2
        Call CloseBackupProgressBar
    End If
End Sub

