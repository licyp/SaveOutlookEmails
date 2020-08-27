VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BackupBar 
   Caption         =   "Back Up Outlook"
   ClientHeight    =   1548
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   4752
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
    End_Code = 1
    Call Close_Backup_Progress_Bar
End Sub

Private Sub Userform_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        End_Code = 2
        Call Close_Backup_Progress_Bar
    End If
End Sub

Private Sub LinkLink_Click()
    Call Open_Url(Link_To_Git_Hub)
End Sub


