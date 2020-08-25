Option Explicit
'These are the used validations

Public fso As Scripting.FileSystemObject ' add MS scripting Runtime

Function Valid_Outlook_Folder(Valid_Outlook_Folder_Input As Outlook.MAPIFolder) As Boolean
'Checks Outlook folder validity based on invalid folders defined in Config sub
Dim Folder_Name As String
Dim i As Double
Dim Number_Of_Invalid_Folders As Double
    Number_Of_Invalid_Folders = UBound(Invalid_Folders) - LBound(Invalid_Folders)
    Folder_Name = Valid_Outlook_Folder_Input.Name
    Valid_Outlook_Folder = True
    For i = 0 To Number_Of_Invalid_Folders
        If UCase(Folder_Name) = UCase(Invalid_Folders(i)) Then
            Valid_Outlook_Folder = False
        End If
    Next
'Debug.Print "############# " & "Valid_Outlook_Folder"
'Debug.Print "Valid_Outlook_Folder_Input: " & Valid_Outlook_Folder_Input
'Debug.Print "Valid_Outlook_Folder: " & Valid_Outlook_Folder
'Debug.Print "############# " & "Valid_Outlook_Folder"
End Function

Function Valid_Outlook_Item(Valid_Outlook_Item_Input) As Boolean 'As Outlook.MailItem)
'Checks Outlook item validity based on valid items defined in Config sub
Dim Item_Name As String
Dim i As Double
Dim Number_Of_Valid_Items As Double
    Number_Of_Valid_Items = UBound(Valid_Items) - LBound(Valid_Items)
    Item_Name = Valid_Outlook_Item_Input.MessageClass
    Valid_Outlook_Item = False
    For i = 0 To Number_Of_Valid_Items
        If UCase(Left(Item_Name, Len(Valid_Items(i)))) = UCase(Valid_Items(i)) Then
            Valid_Outlook_Item = True
        End If
    Next
'Debug.Print "############# " & "Valid_Outlook_Item"
'Debug.Print "Valid_Outlook_Item_Input: " & Valid_Outlook_Item_Input
'Debug.Print "Valid_Outlook_Item: " & Valid_Outlook_Item
'Debug.Print "############# " & "Valid_Outlook_Item"
End Function

Function Archived_Outlook_Item(Archived_Outlook_Item_Input) As Boolean 'As Outlook.MailItem)
'Checks Archived Outlook item based on archived items defined in Config sub
Dim Item_Name As String
Dim i As Double
Dim Number_Of_Archived_Items As Double
    Number_Of_Archived_Items = UBound(Archived_Array) - LBound(Archived_Array)
    Item_Name = Archived_Outlook_Item_Input.MessageClass
    Archived_Outlook_Item = False
    For i = 0 To Number_Of_Archived_Items
        If UCase(Right(Item_Name, Len(Archived_Array(i)))) = UCase(Archived_Array(i)) Then
            Archived_Outlook_Item = True
        End If
    Next
'Debug.Print "############# " & "Valid_Outlook_Item"
'Debug.Print "Archived_Outlook_Item_Input: " & Archived_Outlook_Item_Input
'Debug.Print "Valid_Outlook_Item: " & Valid_Outlook_Item
'Debug.Print "############# " & "Valid_Outlook_Item"
End Function

Sub File_Exists_In_Log_Or_HDD()
'Checks HDD for selected Outlook item (exact match) or log file (date time and partial subject match)
Dim i As Double
Dim Item_Date As String
Dim Log_Date As String
Dim Overlap_Subject_Real As Double
    
    Set fso = New Scripting.FileSystemObject
    Outlook_Item_Saved_Already = False
    If Auto_Run = False Then
        If Last_Found_Was_At = 0 Then
            If From_New_To_Old = False Then
                Last_Found_Was_At = LBound(Archived_File_Array, 2)
            Else
                Last_Found_Was_At = UBound(Archived_File_Array, 2)
            End If
        End If
    Else
        Outlook_Item_Saved_Already = fso.FileExists(Item_Short_Array(2) & Item_Short_Array(0) & " - " & Item_Short_Array(1))
        Exit Sub
    End If
    
    If Item_Short_Array(0) = "" Then
            Outlook_Item_Saved_Already = True
            Date_Error = True
            Exit Sub
    End If
    
    Item_Date = Text_To_Date_Time(Item_Short_Array(0) & " ")
    
    If Overlap_Subject > Len(Item_Short_Array(1)) Then
        Overlap_Subject_Real = Len(Item_Short_Array(1)) - 1
    Else
        Overlap_Subject_Real = Overlap_Subject
    End If
    
'Debug.Print "Looking for: " & Item_Date
    If From_New_To_Old = False Then
        For i = Last_Found_Was_At + 1 To UBound(Archived_File_Array, 2)
            If Archived_File_Array(0, i) <> File_Array_Heading(0) Then
                Log_Date = Text_To_Date_Time(Archived_File_Array(0, i) & " ")
            Else
                Log_Date = Archived_File_Array(0, i)
            End If
'Debug.Print "Is it this one? " & Archived_File_Array(0, i)
            If IsDate(Log_Date) Then
                If DateValue(Log_Date) = DateValue(Item_Date) And _
                    TimeValue(Log_Date) = TimeValue(Item_Date) Then
                    If Left(Replace_Illegal_Chars_File_Folder_Name(Item_Short_Array(1) & " ", Replace_Char_By, Max_File_Name_Length, False), Overlap_Subject_Real) _
                        = Left(Archived_File_Array(1, i), Overlap_Subject_Real) Then
                        Outlook_Item_Saved_Already = True
                        Last_Found_Was_At = i
                        Exit Sub
                    End If
                End If
            End If
        Next
    Else
        For i = Last_Found_Was_At - 1 To LBound(Archived_File_Array, 2) Step -1
'Debug.Print "Is it this one? " & Archived_File_Array(1, i)
            If IsDate(Archived_File_Array(1, i)) Then
                If DateValue(Archived_File_Array(1, i)) = DateValue(Text_To_Date_Time(Item_Short_Array(0) & " ")) And _
                    TimeValue(Archived_File_Array(1, i)) = TimeValue(Text_To_Date_Time(Item_Short_Array(0) & " ")) Then
                    If Left(Replace_Illegal_Chars_File_Folder_Name(Item_Short_Array(1) & " ", Replace_Char_By, Max_File_Name_Length, False), Overlap_Subject_Real) _
                        = Left(Archived_File_Array(1, i), Overlap_Subject_Real) Then
                        Outlook_Item_Saved_Already = True
                        Last_Found_Was_At = i
                        Exit Sub
                    End If
                End If
            End If
        Next
    End If
    
    If fso.FileExists(Item_Short_Array(2) & Item_Short_Array(0) & " - " & Item_Short_Array(1)) = False Then
        If From_New_To_Old = False Then
            Last_Found_Was_At = LBound(Archived_File_Array, 2)
        Else
            Last_Found_Was_At = UBound(Archived_File_Array, 2)
        End If
    Else
        Outlook_Item_Saved_Already = fso.FileExists(Item_Short_Array(2) & Item_Short_Array(0) & " - " & Item_Short_Array(1))
        Call Log_File_Add_Line(Item_Short_Array)
        Exit Sub
    End If

'Debug.Print "############# " & "File_Exists_In_Log_Or_HDD"
'Debug.Print "OutlookItem: " & Item_Short_Array(2) & Item_Short_Array(0) & " - " & Item_Short_Array(1)
'Debug.Print "Outlook_Item_Saved_Already: " & Outlook_Item_Saved_Already
'Debug.Print "############# " & "File_Exists_In_Log_Or_HDD"
End Sub
