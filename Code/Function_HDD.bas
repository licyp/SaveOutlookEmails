Option Explicit
'These are to create folders on local drive

Public fso As Scripting.FileSystemObject ' add MS scripting Runtime

Sub Create_HDD_Folder_For_Outlook_Folder(Create_HDD_Folder_For_Outlook_Folder_Input As Outlook.MAPIFolder)
'Checks existence of folder on HDD of the selected Outlook folder; if doesn't exist call sub to create one
Dim Folder_Loop As Outlook.MAPIFolder
Dim Sub_Folder_Loop As Outlook.MAPIFolder
    Set Folder_Loop = Create_HDD_Folder_For_Outlook_Folder_Input
    If Valid_Outlook_Folder(Folder_Loop) = True Then
        Call Create_HDD_Folder(Default_Backup_Location & "\" & Clean_Outlook_Full_Path_Name(Folder_Loop))
    Else
    End If
'Process all folders and subfolders recursively
    If Folder_Loop.Folders.Count And Valid_Outlook_Folder(Folder_Loop) = True Then
       For Each Sub_Folder_Loop In Folder_Loop.Folders
           Call Create_HDD_Folder_For_Outlook_Folder(Sub_Folder_Loop)
       Next
    End If
'Debug.Print "############# " & "Create_HDD_Folder_For_Outlook_Folder"
'Debug.Print "Create_HDD_Folder_For_Outlook_Folder_Input: " & Create_HDD_Folder_For_Outlook_Folder_Input
'Debug.Print "############# " & "Create_HDD_Folder_For_Outlook_Folder"
End Sub

Sub Create_HDD_Folder(Create_HDD_Folder_Input As String)
'Creates folder on HDD if doesn't exist
Dim HDD_Folder_Loop As Scripting.Folder
Dim Folder_Name_Test As String
Dim Folder_Name_Test_Root As String
Dim i As Double
    Set fso = New Scripting.FileSystemObject
    Folder_Name_Test = Create_HDD_Folder_Input
    If fso.FolderExists(Folder_Name_Test) = False Then
        i = InStr(10, Folder_Name_Test, "\", 1)
        Folder_Name_Test_Root = Left(Folder_Name_Test, i)
        While fso.FolderExists(Folder_Name_Test) = False
            i = InStr(i, Folder_Name_Test, "\", 1)
            If i = 0 Or i = 1 Then
                Folder_Name_Test_Root = Folder_Name_Test
            Else
                Folder_Name_Test_Root = Left(Folder_Name_Test, i)
            End If
            If fso.FolderExists(Folder_Name_Test_Root) = False Then
                MkDir Folder_Name_Test_Root
            End If
            i = i + 1
        Wend
    End If
    Set fso = Nothing
    Set HDD_Folder_Loop = Nothing
'Debug.Print "############# " & "Create_HDD_Folder"
'Debug.Print "Create_HDD_Folder_Input: " & Create_HDD_Folder_Input
'Debug.Print "############# " & "Create_HDD_Folder"
End Sub
