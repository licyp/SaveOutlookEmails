Option Explicit
'These are cleaning up folder and file names; converts text to datetime

Function Clean_Outlook_Full_Path_Name(Clean_Outlook_Folder_Name_Input As Outlook.MAPIFolder) As String
'Removed invalid characters (using Replace_Illegal_Chars_File_Folder_Name sub) from the full path name of the selected Outlook folder
Dim Full_Path_Folder_Loop As Outlook.MAPIFolder
    Set Full_Path_Folder_Loop = Clean_Outlook_Folder_Name_Input
    Do While Full_Path_Folder_Loop.Parent <> "Mapi"
        If Clean_Outlook_Full_Path_Name = "" Then
            Clean_Outlook_Full_Path_Name = _
                Replace_Illegal_Chars_File_Folder_Name(Full_Path_Folder_Loop.Name, Replace_Char_By, Max_Folder_Name_Length, False)
        Else
            Clean_Outlook_Full_Path_Name = _
                Replace_Illegal_Chars_File_Folder_Name(Full_Path_Folder_Loop.Name, Replace_Char_By, Max_Folder_Name_Length, False) & "\" & _
                Clean_Outlook_Full_Path_Name
        End If
        Set Full_Path_Folder_Loop = Full_Path_Folder_Loop.Parent
    Loop
    Clean_Outlook_Full_Path_Name = _
            Replace_Illegal_Chars_File_Folder_Name(Full_Path_Folder_Loop.Name, Replace_Char_By, Max_Folder_Name_Length, False) & "\" & _
            Clean_Outlook_Full_Path_Name
    Set Full_Path_Folder_Loop = Nothing
'Debug.Print "############# " & "Clean_Outlook_Full_Path_Name"
'Debug.Print "Clean_Outlook_Folder_Name_Input: " & Clean_Outlook_Folder_Name_Input
'Debug.Print "Clean_Outlook_Full_Path_Name: " & Clean_Outlook_Full_Path_Name
'Debug.Print "############# " & "Clean_Outlook_Full_Path_Name"
End Function

Function Replace_Illegal_Chars_File_Folder_Name(To_Check_Name As String, Replace_By As String, Max_Length As Double, Folder_Check As Boolean) As String
'Replace illegal characters by 'Replace_By' of the selected string and trims it to 'Max_Length'; if 'Folder_Check' is true then '\' and ':' are accepted characters
Dim Suffix_Length As Double
Dim Working_String As String
    Working_String = To_Check_Name
    
'Replace illegal characters
    Working_String = Replace(Working_String, "*", Replace_By) '(asterisk)
    Working_String = Replace(Working_String, "/", Replace_By) '(forward slash)
    If Folder_Check = False Then
        Working_String = Replace(Working_String, "\", Replace_By) '(backslash)
        Working_String = Replace(Working_String, ":", Replace_By) '(colon)
    Else
    End If
    Working_String = Replace(Working_String, "?", Replace_By) '(question mark)
    Working_String = Replace(Working_String, Chr(34), Replace_By) '(double quote)
    Working_String = Replace(Working_String, "%", Replace_By) '(percent)
    Working_String = Replace(Working_String, "<", Replace_By) '(less than)
    Working_String = Replace(Working_String, ">", Replace_By) '(greater than)
    Working_String = Replace(Working_String, "|", Replace_By) '(vertical bar)
'    Working_String = Replace(Working_String, "#", Replace_By)'(hash)
'    Working_String = Replace(Working_String, "@", Replace_By) '(at)
'    Working_String = Replace(Working_String, "'", Replace_By)'(apostrophe)
    Working_String = Replace(Working_String, Chr(10), Replace_By) '(line feed)
    Working_String = Replace(Working_String, Chr(13), Replace_By) '(carriage return)
    Working_String = Replace(Working_String, Chr(9), Replace_By) '(horizontal tabulation)
    
'Clean up Replace_By mess and double spaces
    Do While InStr(1, Working_String, Replace_By & " " & Replace_By) <> 0 Or _
            InStr(1, Working_String, Replace_By & Replace_By) <> 0 Or _
            InStr(1, Working_String, "  ") <> 0
        Working_String = Replace(Working_String, Replace_By & " " & Replace_By, Replace_By)
        Working_String = Replace(Working_String, Replace_By & Replace_By, Replace_By)
        Working_String = Replace(Working_String, "  ", " ")
    Loop
    
'Shorten name and add suffix
    Suffix_Length = Len(Suffix_Text)
    If Len(Working_String) >= Max_Length + 1 Then
        If InStr(Max_Length - (Suffix_Length + 1), Working_String, " ") = Max_Length - Suffix_Length Then
            Working_String = Left(Working_String, Max_Length - (Suffix_Length + 1)) & Suffix_Text
        Else
            Working_String = Left(Working_String, Max_Length - Suffix_Length) & Suffix_Text
        End If
    End If
    Replace_Illegal_Chars_File_Folder_Name = Working_String
'Debug.Print "############# " & "Replace_Illegal_Chars_File_Folder_Name"
'Debug.Print "To_Check_Name: " & To_Check_Name
'Debug.Print "Replace_Illegal_Chars_File_Folder_Name: " & Replace_Illegal_Chars_File_Folder_Name
'Debug.Print "Replace_By: " & Replace_By
'Debug.Print "Max_Length: " & Max_Length
'Debug.Print "Folder_Check: " & Folder_Check
'Debug.Print "############# " & "Replace_Illegal_Chars_File_Folder_Name"
End Function

Function Text_To_Date_Time(Text_To_Date_Time_Input As String)
'Converts 'YYYY*MM*DD*hhmmss' (from file name or log file) to date and time values
Text_To_Date_Time = DateSerial(Mid(Text_To_Date_Time_Input, 1, 4) * 1, Mid(Text_To_Date_Time_Input, 6, 2) * 1, Mid(Text_To_Date_Time_Input, 9, 2) * 1) + _
    TimeSerial(Mid(Text_To_Date_Time_Input, 12, 2) * 1, Mid(Text_To_Date_Time_Input, 14, 2) * 1, Mid(Text_To_Date_Time_Input, 16, 2) * 1)
'Debug.Print "############# " & "Text_To_Date_Time"
'Debug.Print "Text_To_Date_Time_Input: " & Text_To_Date_Time_Input
'Debug.Print "Text_To_Date_Time: " & Text_To_Date_Time
'Debug.Print "############# " & "Text_To_Date_Time"
End Function
