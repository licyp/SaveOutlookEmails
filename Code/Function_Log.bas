Option Explicit
'These are creating and handling arrays and log files

Public fso As Scripting.FileSystemObject ' add MS scripting Runtime

Sub Add_To_Item_Array(Outlook_Folder_Input As Outlook.MAPIFolder, Outlook_Item_Input) 'As Outlook.MailItem)
'Creates Array_Line for Outlook item with attributes to be validated and with statuses
Dim Status As String
Dim Error_Type As String
Dim OL_Folder As Outlook.MAPIFolder
Dim OL_Folder_Type_Valid As Boolean
Dim OL_Item_In_Folder_Count As Double
Dim OL_Item As Object 'Outlook.MailItem
Dim OL_Item_Title As String
Dim OL_Item_Date As String
Dim OL_Item_Unread As Boolean
Dim OL_Item_Title_Short As String
Dim OL_Item_Type As String
Dim OL_Item_Type_Valid As Boolean
Dim OL_Item_Size As Double
Dim OL_Item_Size_Valid As Boolean
Dim OL_Item_From As String
Dim OL_Item_To As String
Dim OL_Item_To_Count As Double
Dim OL_Item_To_Count_Valid As Boolean
Dim HDD_Folder_Path As String
Dim HDD_Item_Valid_Length As Boolean
Dim Full_Item_Path As String
Dim Clean_Item_Name As String
Dim i As Double
Dim Class_Check_Char As Double

    On Error GoTo NetworkError

    Set OL_Folder = Outlook_Folder_Input
    Set OL_Item = Outlook_Item_Input
    Class_Check_Char = 8

    OL_Folder_Type_Valid = Valid_Outlook_Folder(OL_Folder)
'Check validity
    If OL_Folder_Type_Valid = False Then
        If Error_Type = "" Then
            Error_Type = "Invalid Folder"
        Else
            Error_Type = Error_Type & "; " & "Invalid Folder"
        End If
    End If

    OL_Item_In_Folder_Count = OL_Folder.Items.Count
    HDD_Folder_Path = Default_Backup_Location & "\" & Clean_Outlook_Full_Path_Name(OL_Folder) & "\"
    OL_Item_Type_Valid = Valid_Outlook_Item(OL_Item)
    OL_Item_Title = OL_Item.Subject
    OL_Item_Size = OL_Item.Size
    OL_Item_Type = OL_Item.MessageClass 'Class

    Select Case Left(OL_Item.MessageClass, Class_Check_Char)
        Case Left("IPM.Appointment", Class_Check_Char) 'Appointment
'Debug.Print "RecurrenceState: " & OL_Item.RecurrenceState
            OL_Item_Type = "Appointment"
            OL_Item_Date = Format(OL_Item.Start, "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
                Format(OL_Item.Start, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem)
            OL_Item_Unread = OL_Item.Unread
            OL_Item_From = OL_Item.Organizer
            OL_Item_To_Count = OL_Item.Recipients.Count
            If OL_Item_To_Count = 0 Then
                OL_Item_To = "-"
            Else
                If OL_Item_To_Count > Max_Item_To Then
                    For i = 1 To Max_Item_To
                        If i = 1 Then
                            OL_Item_To = OL_Item.Recipients(i)
                        Else
                            OL_Item_To = OL_Item_To & ", " & OL_Item.Recipients(i)
                        End If
                    Next
                Else
                    For i = 1 To OL_Item_To_Count
                        If i = 1 Then
                            OL_Item_To = OL_Item.Recipients(i)
                        Else
                            OL_Item_To = OL_Item_To & ", " & OL_Item.Recipients(i)
                        End If
                    Next
                End If
            End If
        Case Left("IPM.Note", Class_Check_Char) 'Mail
            OL_Item_Type = "Mail"
            OL_Item_Date = Format(OL_Item.ReceivedTime, "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
                Format(OL_Item.ReceivedTime, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem)
            OL_Item_Unread = OL_Item.Unread
            OL_Item_From = OL_Item.SenderName
            OL_Item_To_Count = OL_Item.Recipients.Count
            If OL_Item_To_Count = 0 Then
                OL_Item_To = "-"
            Else
                If OL_Item_To_Count > Max_Item_To Then
                    For i = 1 To Max_Item_To
                        If i = 1 Then
                            OL_Item_To = OL_Item.Recipients(i)
                        Else
                            OL_Item_To = OL_Item_To & ", " & OL_Item.Recipients(i)
                        End If
                    Next
                Else
                    For i = 1 To OL_Item_To_Count
                        If i = 1 Then
                            OL_Item_To = OL_Item.Recipients(i)
                        Else
                            OL_Item_To = OL_Item_To & ", " & OL_Item.Recipients(i)
                        End If
                    Next
                End If
            End If
        Case Left("IPM.Schedule.Meeting.Resp.Tent", Class_Check_Char), _
            Left("IPM.Schedule.Meeting.Resp.Pos", Class_Check_Char), _
            Left("IPM.Schedule.Meeting.Resp.Neg", Class_Check_Char), _
            Left("IPM.Schedule.Meeting.Request", Class_Check_Char), _
            Left("IPM.Schedule.Meeting.Canceled", Class_Check_Char) 'Meeting
            OL_Item_Type = "Meeting"
            OL_Item_Date = Format(OL_Item.ReceivedTime, "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
                Format(OL_Item.ReceivedTime, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem)
            OL_Item_Unread = OL_Item.Unread
            OL_Item_From = OL_Item.SenderName
            OL_Item_To_Count = OL_Item.Recipients.Count
            If OL_Item_To_Count = 0 Then
                OL_Item_To = "-"
            Else
                If OL_Item_To_Count > Max_Item_To Then
                    For i = 1 To Max_Item_To
                        If i = 1 Then
                            OL_Item_To = OL_Item.Recipients(i)
                        Else
                            OL_Item_To = OL_Item_To & ", " & OL_Item.Recipients(i)
                        End If
                    Next
                Else
                    For i = 1 To OL_Item_To_Count
                        If i = 1 Then
                            OL_Item_To = OL_Item.Recipients(i)
                        Else
                            OL_Item_To = OL_Item_To & ", " & OL_Item.Recipients(i)
                        End If
                    Next
                End If
            End If
        Case Left("IPM.StickyNote", Class_Check_Char) 'Note
            OL_Item_Type = "Note"
            OL_Item_Date = Format(OL_Item.LastModificationTime, "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
                Format(OL_Item.LastModificationTime, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem)
        Case Left("IPM.Task", Class_Check_Char) 'Task
            OL_Item_Type = "Task"
            OL_Item_Date = Format(OL_Item.StartDate, "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
                Format(OL_Item.StartDate, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem)
            OL_Item_Unread = OL_Item.Unread
            OL_Item_From = OL_Item.Owner
            OL_Item_To_Count = OL_Item.Recipients.Count
            If OL_Item_To_Count = 0 Then
                OL_Item_To = "-"
            Else
                If OL_Item_To_Count > Max_Item_To Then
                    For i = 1 To Max_Item_To
                        If i = 1 Then
                            OL_Item_To = OL_Item.Recipients(i)
                        Else
                            OL_Item_To = OL_Item_To & ", " & OL_Item.Recipients(i)
                        End If
                    Next
                Else
                    For i = 1 To OL_Item_To_Count
                        If i = 1 Then
                            OL_Item_To = OL_Item.Recipients(i)
                        Else
                            OL_Item_To = OL_Item_To & ", " & OL_Item.Recipients(i)
                        End If
                    Next
                End If
            End If
        Case Else
            OL_Item_Type_Valid = False
    End Select
'Check validity
    If OL_Item_Type_Valid = False Then
        If Error_Type = "" Then
            Error_Type = "Invalid Item"
        Else
            Error_Type = Error_Type & "; " & "Invalid Item"
        End If
        GoTo NotValid:
    Else
    End If

    Clean_Item_Name = Replace_Illegal_Chars_File_Folder_Name(OL_Item_Title, Replace_Char_By, Max_File_Name_Length, False)
    Full_Item_Path = HDD_Folder_Path & OL_Item_Date & " - " & Clean_Item_Name & ".msg"
    If Len(Full_Item_Path) > Max_Path_Length Then
        If Max_Path_Length - Len(HDD_Folder_Path) - Len(OL_Item_Date & " - " & ".msg") < Min_File_Name_Length Then
            OL_Item_Title_Short = "-"
            HDD_Item_Valid_Length = False
        Else
            Clean_Item_Name = Replace_Illegal_Chars_File_Folder_Name(OL_Item_Title, Replace_Char_By, _
                Max_Path_Length - Len(HDD_Folder_Path) - Len(OL_Item_Date & " - " & ".msg"), False)
            OL_Item_Title_Short = OL_Item_Date & " - " & Clean_Item_Name & ".msg"
            HDD_Item_Valid_Length = True
        End If
    Else
        OL_Item_Title_Short = OL_Item_Date & " - " & Clean_Item_Name & ".msg"
        HDD_Item_Valid_Length = True
    End If
'Check validity
    If HDD_Item_Valid_Length = False Then
        If Error_Type = "" Then
            Error_Type = "File Name Length"
        Else
            Error_Type = Error_Type & "; " & "File Name Length"
        End If
    Else
    End If

    If OL_Item_Size > Max_Item_Size Then
        OL_Item_Size_Valid = False
    Else
        OL_Item_Size_Valid = True
    End If
'Check validity
    If OL_Item_Size_Valid = False Then
        If Error_Type = "" Then
            Error_Type = "Email size"
        Else
            Error_Type = Error_Type & "; " & "Email size"
        End If
    Else
    End If

    If OL_Item_To_Count > Max_Item_To Then
        OL_Item_To_Count_Valid = False
    Else
        OL_Item_To_Count_Valid = True
    End If
'Check validity
    If OL_Item_To_Count_Valid = False Then
        If Error_Type = "" Then
            Error_Type = "Number of recipients"
        Else
            Error_Type = Error_Type & "; " & "Number of recipients"
        End If
    Else
    End If

NotValid:
    If Error_Type = "" Then
        Status = "Listed"
        Error_Type = "OK"
    Else
        Status = "Error"
    End If

    Item_Array = Array(Status, Error_Type, _
        OL_Folder, OL_Folder_Type_Valid, OL_Item_In_Folder_Count, _
        OL_Item_Title, OL_Item_Date, OL_Item_Unread, OL_Item_From, OL_Item_To, OL_Item_Title_Short, _
        OL_Item_Type, OL_Item_Type_Valid, OL_Item_Size, OL_Item_Size_Valid, _
        OL_Item_To_Count, OL_Item_To_Count_Valid, _
        HDD_Folder_Path, HDD_Item_Valid_Length)

NetworkError:
If Err.Number <> 0 And Left(Err.Description, Len("Network")) = "Network" Then
    MsgBox "Hmmm..., Add_To_Item_Array"
End If
'Debug.Print "############# " & "Add_To_Item_Array"
'Debug.Print "Outlook_Folder_Input: " & Outlook_Folder_Input
'Debug.Print "Outlook_Item_Input: " & Outlook_Item_Input
'Debug.Print "Status: " & Status
'Debug.Print "Error_Type: " & Error_Type
'Debug.Print "OL_Folder: " & OL_Folder
'Debug.Print "OL_Folder_Type_Valid: " & OL_Folder_Type_Valid
'Debug.Print "OL_Item_In_Folder_Count: " & OL_Item_In_Folder_Count
'Debug.Print "OL_Item_Title: " & OL_Item_Title
'Debug.Print "OL_Item_Date: " & OL_Item_Date
'Debug.Print "OL_Item_From: " & OL_Item_From
'Debug.Print "OL_Item_To: " & OL_Item_To
'Debug.Print "OL_Item_Title_Short: " & OL_Item_Title_Short
'Debug.Print "OL_Item_Type: " & OL_Item_Type
'Debug.Print "OL_Item_Type_Valid: " & OL_Item_Type_Valid
'Debug.Print "OL_Item_Size: " & OL_Item_Size
'Debug.Print "OL_Item_Size_Valid: " & OL_Item_Size_Valid
'Debug.Print "OL_Item_To_Count: " & OL_Item_To_Count
'Debug.Print "OL_Item_To_Count_Valid: " & OL_Item_To_Count_Valid
'Debug.Print "HDD_Folder_Path: " & HDD_Folder_Path
'Debug.Print "HDD_Item_Valid_Length: " & HDD_Item_Valid_Length
'Debug.Print "############# " & "Add_To_Item_Array"
'    For i = LBound(Item_Array) To UBound(Item_Array)
'    Debug.Print Item_Array(i)
'    Next
'Debug.Print "############# " & "Add_To_Item_Array"
End Sub

Sub Add_To_Short_Item_Array(Outlook_Folder_Input As Outlook.MAPIFolder, Outlook_Item_Input) 'As Outlook.MailItem)
'Creates Short Array_Line for Outlook item with attributes to be checked against saved files
Dim OL_Folder As Outlook.MAPIFolder
Dim OL_Item As Object 'Outlook.MailItem
Dim OL_Item_Title As String
Dim OL_Item_Date As String
Dim OL_Item_Title_Short As String
Dim OL_Item_Type As String
Dim HDD_Folder_Path As String
Dim Clean_Item_Name As String
Dim i As Double
Dim Class_Check_Char As Double

    On Error GoTo NetworkError

    Set OL_Folder = Outlook_Folder_Input
    Set OL_Item = Outlook_Item_Input
    Class_Check_Char = 8

    OL_Item_Title = OL_Item.Subject
    OL_Item_Type = OL_Item.MessageClass 'Class
    HDD_Folder_Path = Default_Backup_Location & "\" & Clean_Outlook_Full_Path_Name(OL_Folder) & "\"

    Select Case Left(OL_Item.MessageClass, Class_Check_Char)
        Case Left("IPM.Appointment", Class_Check_Char) 'Appointment
'Debug.Print "RecurrenceState: " & OL_Item.RecurrenceState
            OL_Item_Date = Format(OL_Item.Start, "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
                Format(OL_Item.Start, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem)
        Case Left("IPM.Note", Class_Check_Char) 'Mail
            OL_Item_Date = Format(OL_Item.ReceivedTime, "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
                Format(OL_Item.ReceivedTime, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem)
        Case Left("IPM.Schedule.Meeting.Resp.Tent", Class_Check_Char), _
            Left("IPM.Schedule.Meeting.Resp.Pos", Class_Check_Char), _
            Left("IPM.Schedule.Meeting.Resp.Neg", Class_Check_Char), _
            Left("IPM.Schedule.Meeting.Request", Class_Check_Char), _
            Left("IPM.Schedule.Meeting.Canceled", Class_Check_Char) 'Meeting
            OL_Item_Date = Format(OL_Item.ReceivedTime, "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
                Format(OL_Item.ReceivedTime, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem)
        Case Left("IPM.StickyNote", Class_Check_Char) 'Note
            OL_Item_Date = Format(OL_Item.LastModificationTime, "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
                Format(OL_Item.LastModificationTime, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem)
        Case Left("IPM.Task", Class_Check_Char) 'Task
            OL_Item_Date = Format(OL_Item.StartDate, "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
                Format(OL_Item.StartDate, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem)
        Case Else
    End Select

    Clean_Item_Name = Replace_Illegal_Chars_File_Folder_Name(OL_Item_Title, Replace_Char_By, Max_File_Name_Length, False) & ".msg"

    Item_Short_Array = Array(OL_Item_Date, Clean_Item_Name, HDD_Folder_Path)

NetworkError:
If Err.Number <> 0 And Left(Err.Description, Len("Network")) = "Network" Then
    MsgBox "Hmmm..., Add_To_Short_Item_Array"
End If
'Debug.Print "############# " & "Add_To_Short_Item_Array"
'Debug.Print "Outlook_Folder_Input: " & Outlook_Folder_Input
'Debug.Print "Outlook_Item_Input: " & Outlook_Item_Input
'Debug.Print "OL_Item_Type: " & OL_Item_Type
'Debug.Print "OL_Item_Date: " & OL_Item_Date
'Debug.Print "OL_Item_Title_Short: " & OL_Item_Title_Short
'Debug.Print "HDD_Folder_Path: " & HDD_Folder_Path
'Debug.Print "############# " & "Add_To_Short_Item_Array"
'    For i = LBound(Item_Short_Array) To UBound(Item_Short_Array)
'       Debug.Print Item_Short_Array(i)
'    Next
'Debug.Print "############# " & "Add_To_Short_Item_Array"
End Sub

Sub Log_File_Open(File_Name As String)
'Open the selected file
Dim Where As String
    File_Number = FreeFile
    Where = File_Name
    Open Where For Append Access Write As #File_Number
'Debug.Print "############# " & "Log_File_Open"
'Debug.Print "File_Name: " & File_Name
'Debug.Print "############# " & "Log_File_Open"
End Sub

Sub Log_File_Create_With_Heading(File_Name As String, Heading As Variant)
'Open the selected file or creates it if doesn't exit with header
Dim Rows As Double
Dim Columns As Double
Dim Array_Value As String
Dim Whole_Line As String
Dim Where As String
Dim Array_Heading As Variant
    Set fso = New Scripting.FileSystemObject
    File_Number = FreeFile

    Where = File_Name
    ReDim Array_Heading(UBound(Heading))
    Array_Heading = Heading

    If fso.FileExists(File_Name) = False Then
        File_Number = FreeFile
        Where = File_Name
        Open Where For Output Access Write As #File_Number
        Whole_Line = ""
        For Rows = LBound(Array_Heading) To UBound(Array_Heading)
            Array_Value = Array_Heading(Rows)
            If Whole_Line = "" Then
                Whole_Line = Array_Value
            Else
                Whole_Line = Whole_Line & Separator_In_File & Array_Value
            End If
        Next Rows
        Print #File_Number, Whole_Line
        Close #File_Number
    End If
'Debug.Print "############# " & "Log_File_Create_With_Heading"
'Debug.Print "File_Name: " & File_Name
'Dim i As Double
'    For i = LBound(Array_Heading) To UBound(Array_Heading)
'        Debug.Print "Heading" & i & ": " & Heading(i)
'    Next
'Debug.Print "############# " & "Log_File_Create_With_Heading"
End Sub

Sub Log_File_Add_Line(Array_To_Add As Variant)
'Add an array as line to opened text file
Dim Rows As Double
Dim Columns As Double
Dim Array_Value As String
Dim Whole_Line As String
Dim Where As String
Dim Array_Line As Variant

    Whole_Line = ""
    ReDim Array_Line(UBound(Array_To_Add))
    Array_Line = Array_To_Add

    For Columns = LBound(Array_Line) To UBound(Array_Line)
        Array_Value = Array_Line(Columns)
        If Whole_Line = "" Then
            Whole_Line = Array_Value
        Else
            Whole_Line = Whole_Line & Separator_In_File & Array_Value
        End If
    Next Columns
    Print #File_Number, Whole_Line
'Debug.Print "############# " & "Log_File_Add_Line"
'Debug.Print "Whole_Line: " & Whole_Line
'Debug.Print "############# " & "Log_File_Add_Line"
End Sub

Sub Log_File_Close()
'Close an open txt file
    Close #File_Number
'Debug.Print "############# " & "Log_File_Close"
End Sub

Sub Build_HDD_Array(Optional Folder_To_Check As String, Optional File_Name As String)
'Build initial log file of already saved Outlook items
Dim What As String
Dim Where As String

    Call Wipe_Me_Clean
    Call Set_Config

    If File_Name = Empty Then
        Where = Default_Backup_Location_Log & "\" & Log_File_Sum & ".txt"
    Else
        Where = File_Name
    End If

    If Folder_To_Check = Empty Then
        What = Default_Backup_Location '& "\"
    Else
        What = Folder_To_Check
    End If

    Call Set_Backup_Progress_Bar_Data
    Call HDD_Folder_Item_Count(What)
    Call HDD_Item_To_Array(What)
    Call Rebuild_Log_File
    Unload BackupBar

'Debug.Print "############# " & "Build_HDD_Array"
'Debug.Print "Folder_To_Check: " & Folder_To_Check
'Debug.Print "File_Name: " & File_Name
'Debug.Print "UBound(Archived_File_Array, 2): " & UBound(Archived_File_Array, 2)
'Debug.Print "HDD_Folder_Count: " & HDD_Folder_Count
'Debug.Print "HDD_File_Count: " & HDD_File_Count
'Debug.Print "############# " & "Build_HDD_Array"
End Sub

Sub HDD_Item_To_Array(HDD_Folder_Input As Variant)
'Loop through folders (and subfolders) and files from selected folder and append to HDD array
Dim Folder_Loop As Scripting.Folder
Dim Sub_Folder_Loop As Scripting.Folder
Dim HDD_Folder As Variant
Dim HDD_File As Variant
Dim HDD_File_Name As String
Dim HDD_Date As String
Dim HDD_Subject As String

Dim i As Double
    Set fso = New Scripting.FileSystemObject
    Set Folder_Loop = fso.GetFolder(HDD_Folder_Input)
    If fso.FolderExists(Folder_Loop) = True Then
        HDD_Folder_Count_Today = HDD_Folder_Count_Today + 1
        Set HDD_Folder = Folder_Loop.Files
        For Each HDD_File In HDD_Folder
            HDD_File_Count_Today = HDD_File_Count_Today + 1
            HDD_File_Name = HDD_File.Name
            If Len(HDD_File_Name) > 21 And IsNumeric(Left(HDD_File_Name, 2)) = True Then
                HDD_Date = Left(HDD_File_Name, 17) 'Text_To_Date_Time(HDD_File_Name)
                HDD_Subject = Mid(HDD_File_Name, 21, 9999)
                File_Array = Array(HDD_Date, HDD_Subject, Folder_Loop)
                Call Add_To_HDD_Array(File_Array)
                Progress_Now_Time = Now()
                Call Update_HDD_Progress_Bar(Folder_Loop, HDD_File_Name, "Create array of existing files")
                DoEvents
            End If
        Next
    Else
    End If
'Process all folders and subfolders recursively
    If Folder_Loop.SubFolders.Count Then
       For Each Sub_Folder_Loop In Folder_Loop.SubFolders
           Call HDD_Item_To_Array(Sub_Folder_Loop)
       Next
    End If
'Debug.Print "############# " & "HDD_Item_To_Array"
'Debug.Print "HDD_Folder_Input: " & HDD_Folder_Input
'Debug.Print "UBound(Archived_File_Array, 1): " & UBound(Archived_File_Array, 1)
'Debug.Print "############# " & "HDD_Item_To_Array"
End Sub

Sub Add_To_HDD_Array(NewArray_Line As Variant)
'Add 'Array_Line' to HDD array (horizontal)
Dim New_Line As Variant
Dim Row As Double
Dim Col As Double
Dim Temp_Array As Variant
Dim Array_Row_Size As Double
Dim Array_Column_Size As Double

    New_Line = NewArray_Line
    Array_Row_Size = UBound(New_Line)
    If IsEmpty(Archived_File_Array) Then
        ReDim Archived_File_Array(UBound(File_Array_Heading), 1)
        For Row = 0 To Array_Row_Size
            Archived_File_Array(Row, Col) = File_Array_Heading(Row)
        Next
    Else
        Array_Column_Size = UBound(Archived_File_Array, 2)
        ReDim Preserve Archived_File_Array(Array_Row_Size, Array_Column_Size + 1)
    End If
    Array_Column_Size = UBound(Archived_File_Array, 2)
    For Row = 0 To Array_Row_Size
        Archived_File_Array(Row, Array_Column_Size) = New_Line(Row)
    Next
'Debug.Print "############# " & "Add_To_HDD_Array"
'Debug.Print "UBound(New_Line, 2): " & UBound(New_Line, 2)
'Debug.Print "UBound(Archived_File_Array, 2): " & UBound(Archived_File_Array, 2)
'Debug.Print "############# " & "AddToLog_Array"
'Dim i As Double
'Dim j As Double
'    For i = LBound(Archived_File_Array, 2) To UBound(Archived_File_Array, 2)
'        Debug.Print "                                             New Line"
'        For j = LBound(Archived_File_Array, 1) To UBound(Archived_File_Array, 1)
'            Debug.Print Archived_File_Array(j, i)
'        Next
'    Next
'Debug.Print "############# " & "Add_To_HDD_Array"
End Sub

Sub Log_HDD_File_In_One_Vertical(File_Name As String)
'Creates a text log file with date and headings and a list of files saved (vertical log file from horizontal array)
'Overwrites existing file
Dim Rows As Double
Dim Columns As Double
Dim Array_Value As String
Dim Whole_Line As String
Dim Where As String
Dim File_Number As Double

    File_Number = FreeFile
    Where = File_Name
    Open Where For Output Access Write As #File_Number

    HDD_File_Count = UBound(Archived_File_Array, 2) + 1
    For Columns = LBound(Archived_File_Array, 2) To UBound(Archived_File_Array, 2)
        Whole_Line = ""
        For Rows = LBound(Archived_File_Array, 1) To UBound(Archived_File_Array, 1)
            Array_Value = Archived_File_Array(Rows, Columns)
            If Whole_Line = "" Then
                Whole_Line = Array_Value
            Else
                Whole_Line = Whole_Line & Separator_In_File & Array_Value
            End If
        Next Rows
        Print #File_Number, Whole_Line
        HDD_File_Count_Today = Columns + 1
        Progress_Now_Time = Now()
        Call Update_HDD_Progress_Bar(Where, Whole_Line, "Save log file")
        DoEvents
    Next Columns
    Close #File_Number
'Debug.Print "############# " & "Log_HDD_File_In_One_Vertical"
'Debug.Print "Where: " & Where
'Debug.Print "############# " & "Log_HDD_File_In_One_Vertical"
End Sub

Sub Read_HDD_In_As_Array(File_Name As String) 'File_Name As String)
'Reads log file into array (vertical file into horizontal array)
Dim Where As String
Dim Whole_Line As String
Dim New_Line As Variant
Dim Row As Double
Dim Col As Double
Dim Row_Start As Double
Dim Row_End As Double
Dim Col_Start As Double
Dim Col_End As Double
Dim Temp_Array As Variant
Dim Array_Row_Size As Double
Dim Array_Column_Size As Double

StartAgain:
    File_Number = FreeFile
    Where = File_Name

'On Error GoTo ResetLog

    Array_Row_Size = UBound(File_Array_Heading)
    ReDim New_Line(Array_Row_Size)

    Open Where For Input As #File_Number
    Do Until EOF(1)
        Line Input #File_Number, Whole_Line
On Error GoTo ResetLog
        For Col = 0 To Array_Row_Size
            If Col = 0 Then
                Col_Start = 1
            Else
                Col_Start = Col_End + 1
            End If
            If Col = Array_Row_Size Then
                Col_End = Len(Whole_Line) + 1
            Else
                Col_End = InStr(Col_Start, Whole_Line, Separator_In_File)
            End If
            New_Line(Col) = Mid(Whole_Line, Col_Start, Col_End - Col_Start)
        Next

        If IsEmpty(Archived_File_Array) Then
            ReDim Archived_File_Array(UBound(File_Array_Heading), 1)
            For Row = 0 To Array_Row_Size
                Archived_File_Array(Row, 0) = File_Array_Heading(Row)
            Next
        Else
            Array_Column_Size = UBound(Archived_File_Array, 2)
            ReDim Preserve Archived_File_Array(Array_Row_Size, Array_Column_Size + 1)
        End If
        Array_Column_Size = UBound(Archived_File_Array, 2)
        For Row = 0 To Array_Row_Size
            Archived_File_Array(Row, Array_Column_Size) = New_Line(Row)
        Next
'Dim i As Double
'Dim j As Double
'        i = 1 + i
'        Debug.Print "                                                              New Line"
'        For j = LBound(Archived_File_Array, 1) To UBound(Archived_File_Array, 1)
'            Debug.Print Archived_File_Array(j, i)
'        Next
    Loop
ResetLog:
    Close #File_Number

    If Err.Number = 0 Then
    Else
'Debug.Print Err.Number
'Debug.Print Err.Description
        Err.Clear
        Call Rebuild_Log_File
        GoTo StartAgain
    End If
    
'Debug.Print "ReadIn: end: " & Now()
'Debug.Print "############# " & "Read_HDD_In_As_Array"
'Debug.Print "UBound(Archived_File_Array, 1): " & UBound(Archived_File_Array, 1)
'Debug.Print "UBound(Archived_File_Array, 2): " & UBound(Archived_File_Array, 2)
'Debug.Print "############# " & "Read_HDD_In_As_Array"
'Dim i As Double
'Dim j As Double
'    For i = LBound(Archived_File_Array, 2) To UBound(Archived_File_Array, 2)
'        Debug.Print "                                                              New Line"
'        For j = LBound(Archived_File_Array, 1) To UBound(Archived_File_Array, 1)
'            Debug.Print Archived_File_Array(j, i)
'        Next
'    Next
'Debug.Print "############# " & "Read_HDD_In_As_Array"
End Sub

Sub Rebuild_Log_File(Optional File_Name As String)
'Reads log file into array sort it by date and Resave
Dim i As Long
Dim Where As String
Dim Temp_Array As Variant
Dim New_Line As Variant
Dim Row As Double
Dim Col As Double

    If IsEmpty(Archived_File_Array) Then
'        Call Wipe_Me_Clean
'        Call Set_Config
        If File_Name = Empty Then
            Where = Default_Backup_Location_Log & "\" & Log_File_Sum & ".txt"
        Else
            Where = File_Name
        End If
        Call Set_Backup_Progress_Bar_Data
        Call Read_HDD_In_As_Array(Where)
    Else
        If File_Name = Empty Then
            Where = Default_Backup_Location_Log & "\" & Log_File_Sum & ".txt"
        Else
            Where = File_Name
        End If
    End If

    Call QuickSort2(Archived_File_Array, 1, 0)

    For i = UBound(Archived_File_Array, 2) To 1 Step -1
        If IsNumeric(Left(Archived_File_Array(0, i), 1)) = False Then
            ReDim Preserve Archived_File_Array(UBound(Archived_File_Array, 1), UBound(Archived_File_Array, 2) - 1)
        Else
            GoTo ResaveFile
        End If
    Next

ResaveFile:
    Temp_Array = Archived_File_Array
    Archived_File_Array = Empty

    ReDim New_Line(UBound(Temp_Array, 1))
'    Debug.Print UBound(New_Line)

    For Col = LBound(Temp_Array, 2) To UBound(Temp_Array, 2)
        For Row = LBound(Temp_Array, 1) To UBound(Temp_Array, 1)
             New_Line(Row) = Temp_Array(Row, Col)
        Next
        Call Add_To_HDD_Array(New_Line)
    Next

    Call Log_HDD_File_In_One_Vertical(Where)
    Unload BackupBar
'    Call Wipe_Me_Clean
'Debug.Print "############# " & "Rebuild_Log_File"
'Debug.Print "UBound(Archived_File_Array, 1): " & UBound(Archived_File_Array, 1)
'Debug.Print "UBound(Archived_File_Array, 2): " & UBound(Archived_File_Array, 2)
'Debug.Print "############# " & "Rebuild_Log_File"
'Dim i As Double
'Dim j As Double
'    For i = LBound(Archived_File_Array, 2) To UBound(Archived_File_Array, 2)
'        Debug.Print "                                                              New Line"
'        For j = LBound(Archived_File_Array, 1) To UBound(Archived_File_Array, 1)
'            Debug.Print Archived_File_Array(j, i)
'        Next
'    Next
'Debug.Print "############# " & "Rebuild_Log_File"
End Sub

Sub Log_Last_Item_Checked(File_Name As String)
'Create file with a date of item last checked
Dim Where As String
Dim Whole_Line As String
    File_Number = FreeFile
    Where = File_Name
    Whole_Line = Last_Item_Checked_Date
    Open Where For Output Access Write As #File_Number
    Print #File_Number, Whole_Line
    Close #File_Number
'Debug.Print "############# " & Log_Last_Item_Checked
'Debug.Print "File_Name: " & File_Name
'Debug.Print "Last_Item_Checked_Date: " & Last_Item_Checked_Date
'Debug.Print "############# " & Log_Last_Item_Checked
End Sub

Sub Read_Last_Item_Date_Log(File_Name As String)
'Reads last item date log file and sets Last_Item_Checked_Date
Dim Where As String
Dim Whole_Line As String
    Set fso = New Scripting.FileSystemObject
    File_Number = FreeFile
    Where = File_Name
    If fso.FileExists(Where) Then
        Open Where For Input As #File_Number
        Do Until EOF(1)
            Line Input #File_Number, Whole_Line
            Last_Item_Checked_Date = Whole_Line
        Loop
        Close #File_Number
    End If
'Debug.Print "############# " & "Read_Last_Item_Date_Log"
'Debug.Print "File_Name: " & File_Name
'Debug.Print "Last_Item_Checked_Date: " & Last_Item_Checked_Date
'Debug.Print "############# " & "Read_Last_Item_Date_Log"
End Sub

Sub Add_To_Short_Item_Date(Outlook_Folder_Input As Outlook.MAPIFolder, Outlook_Item_Input) 'As Outlook.MailItem)
'Creates ShortDateArray for Outlook item
Dim OL_Folder As Outlook.MAPIFolder
Dim OL_Item As Object 'Outlook.MailItem
Dim OL_Item_Date As String
Dim OL_Item_Type As String
Dim Class_Check_Char As Double
Dim i As Double

    On Error GoTo NetworkError

    Set OL_Folder = Outlook_Folder_Input
    Set OL_Item = Outlook_Item_Input
    Class_Check_Char = 8

    OL_Item_Type = OL_Item.MessageClass 'Class

    Select Case Left(OL_Item.MessageClass, Class_Check_Char)
        Case Left("IPM.Appointment", Class_Check_Char) 'Appointment
'Debug.Print "RecurrenceState: " & OL_Item.RecurrenceState
            OL_Item_Date = OL_Item.Start
        Case Left("IPM.Note", Class_Check_Char) 'Mail
            OL_Item_Date = Format(OL_Item.ReceivedTime, "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
                Format(OL_Item.ReceivedTime, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem)
        Case Left("IPM.Schedule.Meeting.Resp.Tent", Class_Check_Char), _
            Left("IPM.Schedule.Meeting.Resp.Pos", Class_Check_Char), _
            Left("IPM.Schedule.Meeting.Resp.Neg", Class_Check_Char), _
            Left("IPM.Schedule.Meeting.Request", Class_Check_Char), _
            Left("IPM.Schedule.Meeting.Canceled", Class_Check_Char) 'Meeting
            OL_Item_Date = Format(OL_Item.ReceivedTime, "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
                Format(OL_Item.ReceivedTime, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem)
        Case Left("IPM.StickyNote", Class_Check_Char) 'Note
            OL_Item_Date = Format(OL_Item.LastModificationTime, "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
                Format(OL_Item.LastModificationTime, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem)
        Case Left("IPM.Task", Class_Check_Char) 'Task
            OL_Item_Date = Format(OL_Item.StartDate, "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
                Format(OL_Item.StartDate, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem)
        Case Else
    End Select

    Item_Date_Only = Text_To_Date_Time(OL_Item_Date)

NetworkError:
If Err.Number <> 0 And Left(Err.Description, Len("Network")) = "Network" Then
    MsgBox "Hmmm..., Add_To_Short_Item_DateArray"
End If
'Debug.Print "############# " & "Add_To_Short_Item_DateArray"
'Debug.Print "UBound(Item_Short_Array): " & UBound(Item_Short_Array)
'Debug.Print "############# " & "Add_To_Short_Item_DateArray"
End Sub
