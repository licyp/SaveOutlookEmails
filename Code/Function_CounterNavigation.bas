Attribute VB_Name = "Function_CounterNavigation"
Option Explicit
'These are basic counters helping the main process loops and navigating in folders

Public fso As Scripting.FileSystemObject ' add MS scripting Runtime

Function ActiveOutlookAccount() As Outlook.MAPIFolder
'Set active Outlook account as main account
Dim i As Double
    For i = 1 To Outlook.Application.Session.Folders.Count
        If Outlook.Application.Session.Folders(i) = Application.Session.Accounts(1) Then
            Set ActiveOutlookAccount = Outlook.Application.Session.Folders(i)
        End If
    Next
'Debug.Print "############# " & "ActiveOutlookAccount"
'Debug.Print "ActiveOutlookAccount: " & ActiveOutlookAccount
'Debug.Print "############# " & "ActiveOutlookAccount"
End Function

Function TopOutlookFolder(TopOutlookFolderInput As Outlook.MAPIFolder) As Outlook.MAPIFolder
'Finds Outlook account of the selected Outlook folder
Dim TopFolderLoop As Outlook.MAPIFolder
    Set TopFolderLoop = TopOutlookFolderInput
    Do While TopFolderLoop.Parent <> "Mapi"
        Set TopFolderLoop = TopFolderLoop.Parent
    Loop
    Set TopOutlookFolder = TopFolderLoop
    Set TopFolderLoop = Nothing
'Debug.Print "############# " & "TopOutlookFolder"
'Debug.Print "TopOutlookFolderInput: " & TopOutlookFolderInput
'Debug.Print "TopOutlookFolder: " & TopOutlookFolder
'Debug.Print "############# " & "TopOutlookFolder"
End Function

Function FullPathOutlookFolder(FullPathOutlookFolderInput As Outlook.MAPIFolder) As String
'Creates full path name for the selected Outlook folder
Dim FullPathFolderLoop As Outlook.MAPIFolder
    Set FullPathFolderLoop = FullPathOutlookFolderInput
    Do While FullPathFolderLoop.Parent <> "Mapi"
        FullPathOutlookFolder = FullPathFolderLoop.Parent & "\" & FullPathOutlookFolder
        Set FullPathFolderLoop = FullPathFolderLoop.Parent
    Loop
    FullPathOutlookFolder = FullPathOutlookFolder & FullPathOutlookFolderInput
    Set FullPathFolderLoop = Nothing
'Debug.Print "############# " & "FullPathOutlookFolder"
'Debug.Print "FullPathOutlookFolderInput: " & FullPathOutlookFolderInput
'Debug.Print "FullPathOutlookFolder: " & FullPathOutlookFolder
'Debug.Print "############# " & "FullPathOutlookFolder"
End Function

Sub OutlookFolderItemCount(OutlookFolderItemCountInput As Outlook.MAPIFolder)
'Counts valid Outlook folders (and subfolders) and all Outlooks items within (including invalid ones) from selected Outlook folder
Dim FolderLoop As Outlook.MAPIFolder
Dim SubFolderLoop As Outlook.MAPIFolder
    Set FolderLoop = OutlookFolderItemCountInput
    If ValidOutlookFolder(FolderLoop) = True Then
        OutlookItemCount = OutlookItemCount + FolderLoop.Items.Count
        OutlookFolderCount = OutlookFolderCount + 1
    Else
    End If
'Process all folders and subfolders recursively
    If FolderLoop.Folders.Count Then
       For Each SubFolderLoop In FolderLoop.Folders
           Call OutlookFolderItemCount(SubFolderLoop)
       Next
    End If
'Debug.Print "############# " & "OutlookFolderItemCount"
'Debug.Print "OutlookFolderItemCountInput: " & OutlookFolderItemCountInput
'Debug.Print "OutlookFolderCount: " & OutlookFolderCount
'Debug.Print "OutlookItemCount: " & OutlookItemCount
'Debug.Print "############# " & "OutlookFolderItemCount"
End Sub

Sub HDDFolderItemCount(HDDFolderItemCountInput) ' As Scripting.Folder)
'Counts folders (and subfolders) and files from selected folder
Dim FolderLoop As Scripting.Folder
Dim SubFolderLoop As Scripting.Folder
    Set fso = New Scripting.FileSystemObject
    Set FolderLoop = fso.GetFolder(HDDFolderItemCountInput)
    If fso.FolderExists(FolderLoop) = True Then
        HDDFileCount = HDDFileCount + FolderLoop.Files.Count
        HDDFolderCount = HDDFolderCount + 1
    Else
    End If
'Process all folders and subfolders recursively
    If FolderLoop.SubFolders.Count Then
       For Each SubFolderLoop In FolderLoop.SubFolders
           Call HDDFolderItemCount(SubFolderLoop)
       Next
    End If
'Debug.Print "############# " & "HDDFolderItemCount"
'Debug.Print "HDDFolderItemCountInput: " & HDDFolderItemCountInput
'Debug.Print "HDDFolderCount: " & HDDFolderCount
'Debug.Print "HDDFileCount: " & HDDFileCount
'Debug.Print "############# " & "HDDFolderItemCount"
End Sub
