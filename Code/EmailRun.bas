'https://docs.microsoft.com/en-us/office/vba/api/overview/outlook
Option Explicit

Public fso As Scripting.FileSystemObject ' add MS scripting Runtime
Public HDD_Main_Folder As Scripting.Folder
Public HDD_Sub_Folder As Scripting.Folder

Public Outlook_Account_Folder As Outlook.MAPIFolder
Public Outlook_Main_Folder As Outlook.MAPIFolder
Public Outlook_Current_Folder As Outlook.MAPIFolder
Public Outlook_Sub_Folder As Outlook.MAPIFolder

Public Last_Found_Was_At As Double

Public Outlook_Folder_Count As Double
Public Outlook_Item_Count As Double
Public HDD_Folder_Count As Double
Public HDD_File_Count As Double
Public HDD_File_Count_Today As Double
Public HDD_Folder_Count_Today As Double
Public Saved_Already_Counter As Double
Public Saved_Already_First As Double
Public Saved_Already_Current As Double
Public Overlap_Days As Double
Public Overlap_Resaved As Double
Public Overlap_Subject As Double

Public Default_Backup_Location As String
Public Default_Backup_Location_Log As String
Public Last_Saved_Item_Date_Log As String
Public Invalid_Folders As Variant
Public Valid_Items As Variant
Public Archived_Array As Variant
Public Default_Folder As String
Public Logged_In_User_File_Location As String
Public Max_Folder_Name_Length As Double
Public Max_File_Name_Length As Double
Public Min_File_Name_Length As Double
Public Max_Path_Length As Double
Public Max_Item_To As Double
Public Max_Item_Size As Double
Public Replace_Char_By As String
Public Suffix_Text As String
Public Force_Resave As Boolean
Public Save_Result As String
Public File_Number As Double
Public Save_Items_To_HDD As Boolean
Public Log_File_Flow As String
Public Log_File_Sum As String
Public Error_Skip_Code As String
Public Default_Config_File As String

Public Outlook_Folder_Current_Count  As Double
Public Outlook_Item_Current_Count  As Double
Public Outlook_Item_Current_Count_Today  As Double
Public Outlook_Item_Current_Count_Today_In_Folder  As Double
Public Outlook_Folder_Current_Count_Today As Double
Public Progress_Start_Time  As Double
Public Progress_Now_Time  As Double
Public End_Code As Double
Public RGB_Step_Count As Double
Public No_Of_Lines_In_File As Double
Public Separator_In_File As String
Public Outlook_Item_Saved_Already As Boolean

Public Log_Array As Variant
Public Item_Array As Variant
Public Item_Short_Array As Variant
Public File_Array As Variant
Public Archived_Log_Array As Variant
Public Archived_File_Array As Variant
Public End_Message As Boolean
Public From_New_To_Old As Boolean
Public Auto_Run As Boolean
Public Item_Array_Heading As Variant
Public File_Array_Heading As Variant
Public Link_To_Git_Hub As String
Public Undeliverable_Error As String
Public Date_Error As Boolean
Public Last_Item_Checked_Date As Double
Public Item_Date_Only As Double

Sub Wipe_Me_Clean()
'Cleans variables (in case of previous unfinished runs)
    Set HDD_Main_Folder = Nothing
    Set HDD_Sub_Folder = Nothing
    Set Outlook_Account_Folder = Active_Outlook_Account
    Set Outlook_Main_Folder = Nothing
    Set Outlook_Current_Folder = Nothing
    Set Outlook_Sub_Folder = Nothing

    Outlook_Folder_Count = 0
    Outlook_Item_Count = 0
    HDD_Folder_Count = 0
    HDD_File_Count = 0
    End_Code = 0
    Saved_Already_Counter = 0
    Last_Found_Was_At = 0
    Log_Array = Empty
    Item_Array = Empty
    Archived_Log_Array = Empty
    Archived_File_Array = Empty
'Debug.Print "############# " & "Wipe_Me_Clean"
End Sub

Sub Set_Config()
'Sets basic boundaries
    Set fso = New Scripting.FileSystemObject

    Suffix_Text = "..."
    Replace_Char_By = "_"
    Separator_In_File = Chr(9)
    Max_Folder_Name_Length = 100
    Max_File_Name_Length = 200
    Min_File_Name_Length = 40
    Max_Path_Length = 240
    Max_Item_To = 250
    Max_Item_Size = 25000000 '25MB
    Overlap_Days = 7
    Overlap_Resaved = 100
    Overlap_Subject = 20

    Default_Config_File = "SaveOutlookEmails.txt"
    Logged_In_User_File_Location = CStr(Environ("USERPROFILE"))
'    Default_Folder = "Desktop\eMails"
'    Default_Backup_Location = Logged_In_User_File_Location & "\" & Default_Folder

Dim Where As String
Dim Whole_Line As String
Dim i As Integer

    Where = Logged_In_User_File_Location & "\" & Default_Config_File

    File_Number = FreeFile
    If fso.FileExists(Where) Then
        Open Where For Input As #File_Number
        Do Until EOF(1)
            Line Input #File_Number, Whole_Line
            i = i + 1
            If i = 1 Then
                Default_Backup_Location = Whole_Line
            Else
            End If
        Loop
        Close #File_Number
    Else
        Default_Folder = "Desktop\eMails"
        Default_Backup_Location = Logged_In_User_File_Location & "\" & Default_Folder

        Open Where For Output Access Write As #File_Number
            Print #File_Number, Default_Backup_Location
            Print #File_Number, ""
            Print #File_Number, "The first lice is used for SaveOutlookEmails bakcup location."
            Print #File_Number, "It should look like:"
            Print #File_Number, "C:\Users\[Your-Name]\Desktop\eMails"
            Print #File_Number, "C:\Users[Your-Name]\OneDrive - [Company-Name]\eMails"
        Close #File_Number
    End If

    Default_Backup_Location_Log = Default_Backup_Location & "\" & "Logs"
    Save_Items_To_HDD = True
    Force_Resave = False
    Log_File_Sum = "Log of Saved Outlook Items"
    Last_Saved_Item_Date_Log = "Last_Checked_Item_Date"
    Link_To_Git_Hub = "https://github.com/licyp/SaveOutlookEmails"

'https://docs.microsoft.com/en-us/office/vba/api/outlook.olDefault_Folders
'Name    Value   Folder Name Description
'OL_FolderConflicts   19  Conflicts   The Conflicts folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
'OL_FolderContacts    10  Contacts    The Contacts folder.
'OL_FolderDeletedItems    3   Deleted Items   The Deleted Items folder.
'OL_FolderJournal 11  Journal The Journal folder.
'OL_FolderJunk    23  Junk E-Mail The Junk E-Mail folder.
'OL_FolderLocalFailures   21  Local Failures  The Local Failures folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
'OL_FolderRssFeeds    25  RSS Feeds   The RSS Feeds folder.
'OL_FolderServerFailures  22  Server Failures The Server Failures folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
'OL_FolderSuggestedContacts   30  Suggested Contacts  The Suggested Contacts folder.
'OL_FolderSyncIssues  20  Sync Issues The Sync Issues folder. Only available for an Exchange account.
'OL_FolderManagedEmail    29      The top-level folder in the Managed Folders group. For more information on Managed Folders, see the Help in Microsoft Outlook. Only available for an Exchange account.
'olPublicFoldersAllPublicFolders 18      The All Public Folders folder in the Exchange Public Folders store. Only available for an Exchange account.
'OL_FolderCalendar    9   Calendar    The Calendar folder.
'OL_FolderDrafts  16  Drafts  The Drafts folder.
'OL_FolderInbox   6   Inbox   The Inbox folder.
'OL_FolderNotes   12  Notes   The Notes folder.
'OL_FolderOutbox  4   Outbox  The Outbox folder.
'OL_FolderSentMail    5   Sent Mail   The Sent Mail folder.
'OL_FolderTasks   13  Tasks   The Tasks folder.
'OL_FolderToDo    28  To Do   The To Do folder.

    Invalid_Folders = Array("Conflicts", "Contacts", "Journal", _
        "Junk E-Mail", "Local Failures", "RSS Feeds", "Server Failures", _
        "Suggested Contacts", "Sync Issues", "Recipient Cache") '"Deleted Items",
    Valid_Items = Array("IPM.Appointment", "IPM.Schedule", "IPM.Note", "IPM.Task", "IPM.StickyNote") ' Start with
    Archived_Array = Array("EAS") ' End with
    Item_Array_Heading = Array("Backup Status", "Error", _
        "Folder", "Folder Validity", "Item Count", _
        "Title", "Date", "Unread", "From", "To", "Shortened Title", _
        "Type", "Type Validity", "Size", "Size Validity", _
        "Recipients Count", "Recipients Validity", _
        "Path on Drive", "Path Validity")
    File_Array_Heading = Array("Date", "Subject", "Path")
    Undeliverable_Error = "Undeliverable_"

    Call Create_HDD_Folder(Default_Backup_Location_Log)
    Call Read_Last_Item_Date_Log(Default_Backup_Location_Log & "\" & Last_Saved_Item_Date_Log & ".txt")
'    If fso.FileExists(Default_Backup_Location_Log & "\" & Log_File_Sum & ".txt") = False Then
'        Call Build_HDD_Array(Default_Backup_Location, Default_Backup_Location_Log & "\" & Log_File_Sum & ".txt")
'    End If

'Debug.Print "############# " & "Set_Config"
End Sub

Sub Quick_Access_Save_Emails()
'Sub to be used on QuickLunch
Dim Msg_Box_Title As String
Dim Msg_Box_Buttons As String
Dim Msg_Box_Text As String
Dim Msg_Box_Response As Double
    Call Wipe_Me_Clean
    Call Set_Config
    End_Message = True
    From_New_To_Old = False
    Auto_Run = False
'Reset log file?
    Msg_Box_Buttons = vbYesNo + vbExclamation + vbDefaultButton1
    Msg_Box_Title = "Reset Log File"
    Msg_Box_Text = "Do you want to reset log file?" & vbNewLine & vbNewLine & _
        "It will speed up scanning already saved files;" & vbNewLine & _
        "only needs doing when scan is slow."

    Msg_Box_Response = MsgBox(Msg_Box_Text, Msg_Box_Buttons, Msg_Box_Title)
    If Msg_Box_Response = 6 Then
        Call Build_HDD_Array
    Else
    End If

    Call Back_Up_Outlook_Folder(Auto_Run)
End Sub

Sub Back_Up_Outlook_Folder(Optional Auto_Run As Boolean)
    Call Wipe_Me_Clean
    Call Set_Config
Dim Msg_Box_Title As String
Dim Msg_Box_Buttons As String
Dim Msg_Box_Text As String
Dim Msg_Box_Response As Double

If Auto_Run = True Then
    Msg_Box_Response = 6
    GoTo Back_Up_Main_Account
End If
'Set folder to back up
StartAgain:
    Msg_Box_Buttons = vbYesNoCancel + vbQuestion + vbDefaultButton1
    Msg_Box_Title = "Backup Outlook Folder"
    Msg_Box_Text = "Back up '" & Outlook_Account_Folder & "' folder instead?"

    Set Outlook_Current_Folder = Outlook.Application.Session.PickFolder
    If Outlook_Current_Folder Is Nothing Then
        Msg_Box_Response = MsgBox(Msg_Box_Text, Msg_Box_Buttons, Msg_Box_Title)

Back_Up_Main_Account:
        Select Case Msg_Box_Response
        Case 6 'Yes
            Set Outlook_Current_Folder = Outlook_Account_Folder
        Case 7 'No
            GoTo StartAgain
        Case 2 'Cancel
            Call Wipe_Me_Clean
            Exit Sub
        End Select
    Else
    End If

'Is chosen folder valid folder for backup?
    Msg_Box_Buttons = vbOKOnly + vbExclamation + vbDefaultButton1
    Msg_Box_Title = "Backup Outlook Folder"
    Msg_Box_Text = "Selected '" & Outlook_Current_Folder & "' folder is not valid for backup."
    If Valid_Outlook_Folder(Outlook_Current_Folder) = False Then
        MsgBox Msg_Box_Text, Msg_Box_Buttons, Msg_Box_Title
        GoTo StartAgain
    Else
    End If

    Call Set_Backup_Progress_Bar_Data
    Set Outlook_Main_Folder = Top_Outlook_Folder(Outlook_Current_Folder)
    Call Outlook_Folder_Item_Count(Outlook_Current_Folder)
    Call Create_HDD_Folder_For_Outlook_Folder(Outlook_Current_Folder)
    Call HDD_Folder_Item_Count(Default_Backup_Location & "\" & Clean_Outlook_Full_Path_Name(Outlook_Current_Folder))
    Call Log_File_Create_With_Heading(Default_Backup_Location_Log & "\" & Log_File_Sum & ".txt", File_Array_Heading)
    Call Read_HDD_In_As_Array(Default_Backup_Location_Log & "\" & Log_File_Sum & ".txt")
    Call Log_File_Open(Default_Backup_Location_Log & "\" & Log_File_Sum & ".txt")
    Call Loop_Outlook_Folders(Outlook_Current_Folder, Save_Items_To_HDD)
    Call Log_File_Close

    Select Case End_Code
        Case 1
            Unload BackupBar
            MsgBox "Cancelled"
        Case 2
            Unload BackupBar
            MsgBox "Red cross"
        Case Else
            If End_Message = False Then
                Unload BackupBar
            Else
                Unload BackupBar
                If Auto_Run = False Then
                    Call Wipe_Me_Clean
                    Call Set_Config
                    Call Set_Backup_Progress_Bar_Data
                    Call Read_HDD_In_As_Array(Default_Backup_Location_Log & "\" & Log_File_Sum & ".txt")
                    Call Rebuild_Log_File
                End If
                Call MsgBox_All_Done
'                MsgBox "All done"
            End If
    End Select

    If Auto_Run = False Then
        Call Log_Last_Item_Checked(Default_Backup_Location_Log & "\" & Last_Saved_Item_Date_Log & ".txt")
    End If
    Call Wipe_Me_Clean

'Debug.Print "############# " & "Back_Up_Outlook_Folder"
'Debug.Print "Outlook_Account_Folder: " & Outlook_Account_Folder
'Debug.Print "Msg_Box_Response: " & Msg_Box_Response
'Debug.Print "Outlook_Current_Folder: " & Outlook_Current_Folder
'Debug.Print "Outlook_Main_Folder: " & Outlook_Main_Folder
'Debug.Print "Full_Path_Outlook_Folder: " & Full_Path_Outlook_Folder(Outlook_Current_Folder)
'Debug.Print "Clean_Outlook_Full_Path_Name: " & Clean_Outlook_Full_Path_Name(Outlook_Current_Folder)
'Debug.Print "Outlook_Folder_Count: " & Outlook_Folder_Count
'Debug.Print "Outlook_Item_Count: " & Outlook_Item_Count
'Debug.Print "Default_Backup_Location: " & Default_Backup_Location
'Debug.Print "HDD_Folder_Count: " & HDD_Folder_Count
'Debug.Print "HDD_File_Count: " & HDD_File_Count
'Debug.Print "############# " & "Back_Up_Outlook_Folder"
End Sub

Sub MsgBox_All_Done()
    Dim Ask_Time As Integer
    Dim Info_Box As Object
    Set Info_Box = CreateObject("WScript.Shell")

    'Set the message box to close after 1 seconds
    Ask_Time = 1
    Select Case Info_Box.Popup("All done." & vbNewLine & "This window will automatically close", Ask_Time, "All done", 0)
        Case 1, -1
            Exit Sub
    End Select
End Sub

Sub MsgBox_Delay_Start()
    Dim Ask_Time As Integer
    Dim Info_Box As Object
    Set Info_Box = CreateObject("WScript.Shell")

    'Set the message box to close after 1 seconds
    Ask_Time = 1
    Select Case Info_Box.Popup("Wait for Outlook to connect..." & vbNewLine & "This window will automatically close.", Ask_Time, "Delayed Start", 0)
        Case 1, -1
            Exit Sub
    End Select
End Sub

Sub Loop_Outlook_Folders(Loop_Outlook_Folders_Input As Outlook.MAPIFolder, Save_Item As Boolean)
Dim Folder_Loop As Outlook.MAPIFolder
Dim Sub_Folder_Loop As Outlook.MAPIFolder
Dim i As Double
Dim k As Double
Dim i_Short_Cut As Double
Dim i_Short_Cut_Step As Double

    Set Folder_Loop = Loop_Outlook_Folders_Input
    If Valid_Outlook_Folder(Folder_Loop) = True Then
        Outlook_Folder_Current_Count_Today = Outlook_Folder_Current_Count_Today + 1
        If From_New_To_Old = False Then
'Add shortcut########################################################################################
            If Folder_Loop.Items.Count > 100 And Force_Resave = False And Auto_Run = False Then
                i_Short_Cut_Step = 1
                HDD_File_Count = 100
                For k = 1 To 100 - i_Short_Cut_Step Step i_Short_Cut_Step
                    HDD_File_Count_Today = k
                    Call Add_To_Short_Item_Date(Folder_Loop, Folder_Loop.Items(Round(Folder_Loop.Items.Count / 100 * k)))
'Debug.Print Round(Folder_Loop.Items.Count / 100 * k) & " at : " & k
                    For i = 1 To UBound(Archived_File_Array, 2)
                        If Archived_File_Array(0, i) = Format(Item_Date_Only, "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
                                Format(Item_Date_Only, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem) Then
                            i_Short_Cut = k
                            Progress_Now_Time = Now()
                            Call Update_HDD_Progress_Bar(Folder_Loop, " ", "Check already saved items")
                            DoEvents
                        End If
                    Next
                    If k > i_Short_Cut Then
                        i_Short_Cut = Round(Folder_Loop.Items.Count / 100 * (i_Short_Cut - i_Short_Cut_Step))
                        GoTo ShortCut
                    End If
                Next
        End If
ShortCut:
            Unload BackupBar
            Call Set_Backup_Progress_Bar_Data
            If i_Short_Cut <= 0 Then
                i = 1
            Else
                i = i_Short_Cut
                Outlook_Item_Current_Count_Today = i - 1
            End If
'Add shortcut########################################################################################
            For i = i To Folder_Loop.Items.Count 'from old to new
                Date_Error = False
'Checks Auto_Run overlap and stop scanning Outlook folder
                If Force_Resave = False And Auto_Run = True And _
                    Saved_Already_Counter > Overlap_Resaved And Abs(Saved_Already_Current - Saved_Already_First) > Overlap_Days Then
                    Exit Sub
                End If
'If cancelled or red cross exist then stop
                If End_Code = 1 Or End_Code = 2 Then
                    Exit Sub
                End If
'Check file existence
                Call Add_To_Short_Item_Array(Folder_Loop, Folder_Loop.Items(i))
                If Left(Item_Short_Array(1), Len(Undeliverable_Error)) = Undeliverable_Error Then
                    GoTo NextItemON
                End If
                Call File_Exists_In_Log_Or_HDD
                If Outlook_Item_Saved_Already = True And Force_Resave = False Then
                    Save_Result = "Saved Already"
                    Saved_Already_Counter = Saved_Already_Counter + 1
                    If Date_Error = False Then
                        If Saved_Already_Counter = 1 Then
                            Saved_Already_First = Text_To_Date_Time(Item_Short_Array(0) & " ")
                        Else
                            Saved_Already_Current = Text_To_Date_Time(Item_Short_Array(0) & " ")
                        End If
                    End If
                Else
                    Saved_Already_Counter = 0
'Read full Outlook item data
                    Call Add_To_Item_Array(Folder_Loop, Folder_Loop.Items(i))
                    If Save_Item = True And Item_Array(1) = "OK" Then
'Save item
                        Call Save_Outlook_Item(Folder_Loop.Items(i), Force_Resave, Item_Array)
                        Item_Array(0) = Save_Result
                        If Error_Skip_Code <> "" Then
                            Item_Array(1) = Error_Skip_Code
                        Else
'Add successful file save to log file
                            Call Log_File_Add_Line(Item_Short_Array)
                        End If
                    End If
                End If
                Outlook_Item_Current_Count_Today_In_Folder = i
                Outlook_Item_Current_Count_Today = Outlook_Item_Current_Count_Today + 1
'Update progress bar
NextItemON:
                Progress_Now_Time = Now()
                Call Update_Backup_Progress_Bar(Folder_Loop, Folder_Loop.Items(i))
                DoEvents
'Debug.Print "Saved_Already_Counter: " & Saved_Already_Counter
            Next
        Else
            For i = Folder_Loop.Items.Count To 1 Step -1 'new old to old
                Date_Error = False
'Checks Auto_Run overlap and stop scanning Outlook folder
                If Force_Resave = False And Auto_Run = True And _
                    Saved_Already_Counter > Overlap_Resaved And Abs(Saved_Already_Current - Saved_Already_First) > Overlap_Days Then
                    Exit Sub
                End If
'If cancelled or red cross exist then stop
                If End_Code = 1 Or End_Code = 2 Then
                    Exit Sub
                End If
'Check file existence
                Call Add_To_Short_Item_Array(Folder_Loop, Folder_Loop.Items(i))
                If Left(Item_Short_Array(1), Len(Undeliverable_Error)) = Undeliverable_Error Then
                    GoTo NextItemNO
                End If
                Call File_Exists_In_Log_Or_HDD
                If Outlook_Item_Saved_Already = True And Force_Resave = False Then
                    Save_Result = "Saved Already"
                    Saved_Already_Counter = Saved_Already_Counter + 1
                    If Date_Error = False Then
                        If Saved_Already_Counter = 1 Then
                            Saved_Already_First = Text_To_Date_Time(Item_Short_Array(0) & " ")
                        Else
                            Saved_Already_Current = Text_To_Date_Time(Item_Short_Array(0) & " ")
                        End If
                    End If
                Else
                    Saved_Already_Counter = 0
'Read full Outlook item data
                    Call Add_To_Item_Array(Folder_Loop, Folder_Loop.Items(i))
                    If Save_Item = True And Item_Array(1) = "OK" Then
'Save item
                        Call Save_Outlook_Item(Folder_Loop.Items(i), Force_Resave, Item_Array)
                        Item_Array(0) = Save_Result
                        If Error_Skip_Code <> "" Then
                            Item_Array(1) = Error_Skip_Code
                        Else
'Add successful file save to log file
                            Call Log_File_Add_Line(Item_Short_Array)
                        End If
                    End If
                End If
                Outlook_Item_Current_Count_Today_In_Folder = i
                Outlook_Item_Current_Count_Today = Outlook_Item_Current_Count_Today + 1
'Update progress bar
NextItemNO:
                Progress_Now_Time = Now()
                Call Update_Backup_Progress_Bar(Folder_Loop, Folder_Loop.Items(i))
                DoEvents
'Debug.Print "Saved_Already_Counter: " & Saved_Already_Counter
            Next
        End If
    Else
    End If
'Process all folders and subfolders recursively
    If Folder_Loop.Folders.Count And Valid_Outlook_Folder(Folder_Loop) = True Then
       For Each Sub_Folder_Loop In Folder_Loop.Folders
           Call Loop_Outlook_Folders(Sub_Folder_Loop, Save_Item)
       Next
    End If
'Debug.Print "############# " & "Loop_Outlook_Folders"
'Debug.Print "Loop_Outlook_Folders_Input: " & Loop_Outlook_Folders_Input
'Debug.Print "Folder_Loop.Items(i): " & Folder_Loop.Items(i).Subject
'Debug.Print "UBound(Log_Array, 1): " & UBound(Log_Array, 1)
'Debug.Print "Saved_Already_Counter: " & Saved_Already_Counter
'Debug.Print "############# " & "Loop_Outlook_Folders"
End Sub

Sub Save_Outlook_Item(Outlook_Item_Input, Resave As Boolean, Item_Data)
'Saves Outlook item if not exists or force Resave=true; other attributes are used from Item_Array related to selected Outlook item
Dim Outlook_App As Outlook.Application
Dim Object_Inspector As Outlook.Inspector
Dim Item_To_Be_Saved As Object
Dim Item_To_Be_Saved_Open As Object
Dim Unread As Boolean
Dim Item_Status As String
Dim Save_File_Name As String
Dim Save_Path_Name As String
Dim File_Exists As Boolean
Dim Archived_Item As Boolean

    Error_Skip_Code = ""
    Set Outlook_App = Outlook.Application
    Set Item_To_Be_Saved = Outlook_Item_Input
    Set fso = New Scripting.FileSystemObject
    Unread = Item_Data(7)
    Item_Status = Item_Data(0)
    Save_Path_Name = Item_Data(17)
    Save_File_Name = Item_Data(10)
    Archived_Item = Archived_Outlook_Item(Item_To_Be_Saved)

    If Item_Status = "Error" Then
        Save_Result = Item_Status
    Else
        If Resave = True Then
            If Archived_Item = True Then
                Set Object_Inspector = Nothing
                Item_To_Be_Saved.Display
                Do While Object_Inspector Is Nothing
                    Set Object_Inspector = Outlook_App.ActiveInspector
                Loop
                Set Item_To_Be_Saved_Open = Object_Inspector.CurrentItem
            Else
                Set Item_To_Be_Saved_Open = Item_To_Be_Saved
            End If
            Item_To_Be_Saved_Open.SaveAs Save_Path_Name & Save_File_Name, olMSG
            If Unread = True Then
                Item_To_Be_Saved_Open.Unread = True
            End If
            If Archived_Item = True Then
                Item_To_Be_Saved_Open.Close olDiscard
            Else
            End If
            Save_Result = "Resaved"
        Else
            If Archived_Item = True Then
                Set Object_Inspector = Nothing
                Item_To_Be_Saved.Display
                Do While Object_Inspector Is Nothing
                    Set Object_Inspector = Outlook_App.ActiveInspector
                Loop
                Set Item_To_Be_Saved_Open = Object_Inspector.CurrentItem
            Else
                Set Item_To_Be_Saved_Open = Item_To_Be_Saved
            End If
On Error GoTo SkipError
            Item_To_Be_Saved_Open.SaveAs Save_Path_Name & Save_File_Name, olMSG
            If Unread = True Then
                Item_To_Be_Saved_Open.Unread = True
            End If
            If Archived_Item = True Then
                Item_To_Be_Saved_Open.Close olDiscard
            Else
            End If
            Save_Result = "Saved"
        End If
    End If

SkipError:
If Err.Number <> 0 Then
    Error_Skip_Code = Err.Number & " " & Err.Description
    Save_Result = "Error"
End If
    Set Outlook_App = Nothing
    Set Object_Inspector = Nothing
    Set fso = Nothing
    Set Item_To_Be_Saved = Nothing
    Set Item_To_Be_Saved_Open = Nothing
'Debug.Print "############# " & "Save_Outlook_Item"
'Debug.Print "Outlook_Item_Input: " & Outlook_Item_Input
'Debug.Print "Resave: " & Resave
'Debug.Print "Item_Status: " & Item_Status
'Debug.Print "Unread: " & Unread
'Debug.Print "Save_Result: " & Save_Result
'Debug.Print Save_Path_Name & Save_File_Name
'Debug.Print "############# " & "Save_Outlook_Item"
End Sub
