Option Explicit
'These are basic counters helping the main process loops and navigating in folders

Public fso As Scripting.FileSystemObject ' add MS scripting Runtime

Function Active_Outlook_Account() As Outlook.MAPIFolder
'Set active Outlook account as main account
Dim i As Double
    For i = 1 To Outlook.Application.Session.Folders.Count
        If Outlook.Application.Session.Folders(i) = Application.Session.Accounts(1) Then
            Set Active_Outlook_Account = Outlook.Application.Session.Folders(i)
        End If
    Next
'Debug.Print "############# " & "Active_Outlook_Account"
'Debug.Print "Active_Outlook_Account: " & Active_Outlook_Account
'Debug.Print "############# " & "Active_Outlook_Account"
End Function

Function Top_Outlook_Folder(Top_Outlook_Folder_Input As Outlook.MAPIFolder) As Outlook.MAPIFolder
'Finds Outlook account of the selected Outlook folder
Dim Top_Folder_Loop As Outlook.MAPIFolder
    Set Top_Folder_Loop = Top_Outlook_Folder_Input
    Do While Top_Folder_Loop.Parent <> "Mapi"
        Set Top_Folder_Loop = Top_Folder_Loop.Parent
    Loop
    Set Top_Outlook_Folder = Top_Folder_Loop
    Set Top_Folder_Loop = Nothing
'Debug.Print "############# " & "Top_Outlook_Folder"
'Debug.Print "Top_Outlook_Folder_Input: " & Top_Outlook_Folder_Input
'Debug.Print "Top_Outlook_Folder: " & Top_Outlook_Folder
'Debug.Print "############# " & "Top_Outlook_Folder"
End Function

Function Full_Path_Outlook_Folder(Full_Path_Outlook_Folder_Input As Outlook.MAPIFolder) As String
'Creates full path name for the selected Outlook folder
Dim Full_Path_Folder_Loop As Outlook.MAPIFolder
    Set Full_Path_Folder_Loop = Full_Path_Outlook_Folder_Input
    Do While Full_Path_Folder_Loop.Parent <> "Mapi"
        Full_Path_Outlook_Folder = Full_Path_Folder_Loop.Parent & "\" & Full_Path_Outlook_Folder
        Set Full_Path_Folder_Loop = Full_Path_Folder_Loop.Parent
    Loop
    Full_Path_Outlook_Folder = Full_Path_Outlook_Folder & Full_Path_Outlook_Folder_Input
    Set Full_Path_Folder_Loop = Nothing
'Debug.Print "############# " & "Full_Path_Outlook_Folder"
'Debug.Print "Full_Path_Outlook_Folder_Input: " & Full_Path_Outlook_Folder_Input
'Debug.Print "Full_Path_Outlook_Folder: " & Full_Path_Outlook_Folder
'Debug.Print "############# " & "Full_Path_Outlook_Folder"
End Function

Sub Outlook_Folder_Item_Count(Outlook_Folder_Item_Count_Input As Outlook.MAPIFolder)
'Counts valid Outlook folders (and subfolders) and all Outlooks items within (including invalid ones) from selected Outlook folder
Dim Folder_Loop As Outlook.MAPIFolder
Dim Sub_Folder_Loop As Outlook.MAPIFolder
    Set Folder_Loop = Outlook_Folder_Item_Count_Input
    If Valid_Outlook_Folder(Folder_Loop) = True Then
        Outlook_Item_Count = Outlook_Item_Count + Folder_Loop.Items.Count
        Outlook_Folder_Count = Outlook_Folder_Count + 1
    Else
    End If
'Process all folders and subfolders recursively
'    If Folder_Loop.Folders.Count And Valid_Outlook_Folder(Folder_Loop) = True Then
    If Folder_Loop.Folders.Count Then
       For Each Sub_Folder_Loop In Folder_Loop.Folders
           Call Outlook_Folder_Item_Count(Sub_Folder_Loop)
       Next
    End If
'Debug.Print "############# " & "Outlook_Folder_Item_Count"
'Debug.Print "Outlook_Folder_Item_Count_Input: " & Outlook_Folder_Item_Count_Input
'Debug.Print "Outlook_Folder_Count: " & Outlook_Folder_Count
'Debug.Print "Outlook_Item_Count: " & Outlook_Item_Count
'Debug.Print "############# " & "Outlook_Folder_Item_Count"
End Sub

Sub HDD_Folder_Item_Count(HDD_Folder_Item_Count_Input) ' As Scripting.Folder)
'Counts folders (and subfolders) and files from selected folder
Dim Folder_Loop As Scripting.Folder
Dim Sub_Folder_Loop As Scripting.Folder
    Set fso = New Scripting.FileSystemObject
    Set Folder_Loop = fso.GetFolder(HDD_Folder_Item_Count_Input)
    If fso.FolderExists(Folder_Loop) = True Then
        HDD_File_Count = HDD_File_Count + Folder_Loop.Files.Count
        HDD_Folder_Count = HDD_Folder_Count + 1
    Else
    End If
'Process all folders and subfolders recursively
    If Folder_Loop.SubFolders.Count Then
       For Each Sub_Folder_Loop In Folder_Loop.SubFolders
           Call HDD_Folder_Item_Count(Sub_Folder_Loop)
       Next
    End If
'Debug.Print "############# " & "HDD_Folder_Item_Count"
'Debug.Print "HDD_Folder_Item_Count_Input: " & HDD_Folder_Item_Count_Input
'Debug.Print "HDD_Folder_Count: " & HDD_Folder_Count
'Debug.Print "HDD_File_Count: " & HDD_File_Count
'Debug.Print "############# " & "HDD_Folder_Item_Count"
End Sub
