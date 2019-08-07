Attribute VB_Name = "Function_HDD"
Option Explicit
'These are to create folders on local drive

Public fso As Scripting.FileSystemObject ' add MS scripting Runtime

Sub CreateHDDFolderForOutlookFolder(CreateHDDFolderForOutlookFolderInput As Outlook.MAPIFolder)
'Checks existence of folder on HDD of the selected Outlook folder; if doesn't exist call sub to create one
Dim FolderLoop As Outlook.MAPIFolder
Dim SubFolderLoop As Outlook.MAPIFolder
    Set FolderLoop = CreateHDDFolderForOutlookFolderInput
    If ValidOutlookFolder(FolderLoop) = True Then
        Call CreateHDDFolder(DefultBackupLocation & "\" & CleanOutlookFullPathName(FolderLoop))
    Else
    End If
'Process all folders and subfolders recursively
    If FolderLoop.Folders.Count Then
       For Each SubFolderLoop In FolderLoop.Folders
           Call CreateHDDFolderForOutlookFolder(SubFolderLoop)
       Next
    End If
'Debug.Print "############# " & "CreateHDDFolderForOutlookFolder"
'Debug.Print "CreateHDDFolderForOutlookFolderInput: " & CreateHDDFolderForOutlookFolderInput
'Debug.Print "############# " & "CreateHDDFolderForOutlookFolder"
End Sub

Sub CreateHDDFolder(CreateHDDFolderInput As String)
'Creates folder on HDD if doesn't exist
Dim HDDFolderLoop As Scripting.Folder
Dim FolderNameTest As String
Dim FolderNameTestRoot As String
Dim i As Double
    Set fso = New Scripting.FileSystemObject
    FolderNameTest = CreateHDDFolderInput
    If fso.FolderExists(FolderNameTest) = False Then
        i = InStr(10, FolderNameTest, "\", 1)
        FolderNameTestRoot = Left(FolderNameTest, i)
        While fso.FolderExists(FolderNameTest) = False
            i = InStr(i, FolderNameTest, "\", 1)
            If i = 0 Or i = 1 Then
                FolderNameTestRoot = FolderNameTest
            Else
                FolderNameTestRoot = Left(FolderNameTest, i)
            End If
            If fso.FolderExists(FolderNameTestRoot) = False Then
                MkDir FolderNameTestRoot
            End If
            i = i + 1
        Wend
    End If
    Set fso = Nothing
    Set HDDFolderLoop = Nothing
'Debug.Print "############# " & "CreateHDDFolder"
'Debug.Print "CreateHDDFolderInput: " & CreateHDDFolderInput
'Debug.Print "############# " & "CreateHDDFolder"
End Sub

