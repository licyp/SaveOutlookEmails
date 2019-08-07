Attribute VB_Name = "Function_CleanUpString"
Option Explicit
'These are cleaning up folder and file names; converts text to datetime

Function CleanOutlookFullPathName(CleanOutlookFolderNameInput As Outlook.MAPIFolder) As String
'Removed invalid characters (using ReplaceIllegalCharsFileFolderName sub) from the full path name of the selected Outlook folder
Dim FullPathFolderLoop As Outlook.MAPIFolder
    Set FullPathFolderLoop = CleanOutlookFolderNameInput
    Do While FullPathFolderLoop.Parent <> "Mapi"
        If CleanOutlookFullPathName = "" Then
            CleanOutlookFullPathName = _
                ReplaceIllegalCharsFileFolderName(FullPathFolderLoop.Name, ReplaceCharBy, MaxFolderNameLenght, False)
        Else
            CleanOutlookFullPathName = _
                ReplaceIllegalCharsFileFolderName(FullPathFolderLoop.Name, ReplaceCharBy, MaxFolderNameLenght, False) & "\" & _
                CleanOutlookFullPathName
        End If
        Set FullPathFolderLoop = FullPathFolderLoop.Parent
    Loop
    CleanOutlookFullPathName = _
            ReplaceIllegalCharsFileFolderName(FullPathFolderLoop.Name, ReplaceCharBy, MaxFolderNameLenght, False) & "\" & _
            CleanOutlookFullPathName
    Set FullPathFolderLoop = Nothing
'Debug.Print "############# " & "CleanOutlookFullPathName"
'Debug.Print "CleanOutlookFolderNameInput: " & CleanOutlookFolderNameInput
'Debug.Print "CleanOutlookFullPathName: " & CleanOutlookFullPathName
'Debug.Print "############# " & "CleanOutlookFullPathName"
End Function

Function ReplaceIllegalCharsFileFolderName(ToCheckName As String, ReplaceBy As String, MaxLength As Double, FolderCheck As Boolean) As String
'Replace illegal characters by 'ReplaceBy' of the selected string and trims it to 'MaxLength'; if 'FolderCheck' is true then '\' and ':' are accepted characters
Dim SuffixLength As Double
Dim WorkingString As String
    WorkingString = ToCheckName
    
'Replace illegal characters
    WorkingString = Replace(WorkingString, "*", ReplaceBy) '(asterisk)
    WorkingString = Replace(WorkingString, "/", ReplaceBy) '(forward slash)
    If FolderCheck = False Then
        WorkingString = Replace(WorkingString, "\", ReplaceBy) '(backslash)
        WorkingString = Replace(WorkingString, ":", ReplaceBy) '(colon)
    Else
    End If
    WorkingString = Replace(WorkingString, "?", ReplaceBy) '(question mark)
    WorkingString = Replace(WorkingString, Chr(34), ReplaceBy) '(double quote)
    WorkingString = Replace(WorkingString, "%", ReplaceBy) '(percent)
    WorkingString = Replace(WorkingString, "<", ReplaceBy) '(less than)
    WorkingString = Replace(WorkingString, ">", ReplaceBy) '(greater than)
    WorkingString = Replace(WorkingString, "|", ReplaceBy) '(vertical bar)
'    WorkingString = Replace(WorkingString, "#", ReplaceBy)'(hash)
'    WorkingString = Replace(WorkingString, "@", ReplaceBy) '(at)
'    WorkingString = Replace(WorkingString, "'", ReplaceBy)'(apostrophe)
    WorkingString = Replace(WorkingString, Chr(10), ReplaceBy) '(line feed)
    WorkingString = Replace(WorkingString, Chr(13), ReplaceBy) '(carriage return)
    WorkingString = Replace(WorkingString, Chr(9), ReplaceBy) '(horizontal tabulation)
    
'Clean up ReplaceBy mess and double spaces
    Do While InStr(1, WorkingString, ReplaceBy & " " & ReplaceBy) <> 0 Or _
            InStr(1, WorkingString, ReplaceBy & ReplaceBy) <> 0 Or _
            InStr(1, WorkingString, "  ") <> 0
        WorkingString = Replace(WorkingString, ReplaceBy & " " & ReplaceBy, ReplaceBy)
        WorkingString = Replace(WorkingString, ReplaceBy & ReplaceBy, ReplaceBy)
        WorkingString = Replace(WorkingString, "  ", " ")
    Loop
    
'Shorten name and add suffix
    SuffixLength = Len(SuffixText)
    If Len(WorkingString) >= MaxLength + 1 Then
        If InStr(MaxLength - (SuffixLength + 1), WorkingString, " ") = MaxLength - SuffixLength Then
            WorkingString = Left(WorkingString, MaxLength - (SuffixLength + 1)) & SuffixText
        Else
            WorkingString = Left(WorkingString, MaxLength - SuffixLength) & SuffixText
        End If
    End If
    ReplaceIllegalCharsFileFolderName = WorkingString
'Debug.Print "############# " & "ReplaceIllegalCharsFileFolderName"
'Debug.Print "ToCheckName: " & ToCheckName
'Debug.Print "ReplaceIllegalCharsFileFolderName: " & ReplaceIllegalCharsFileFolderName
'Debug.Print "ReplaceBy: " & ReplaceBy
'Debug.Print "MaxLength: " & MaxLength
'Debug.Print "FolderCheck: " & FolderCheck
'Debug.Print "############# " & "ReplaceIllegalCharsFileFolderName"
End Function

Function TextToDateTime(TextToDateTimeInput As String)
'Converts 'YYYY*MM*DD*hhmmss' (from file name or log file) to date and time values
TextToDateTime = DateSerial(Mid(TextToDateTimeInput, 1, 4) * 1, Mid(TextToDateTimeInput, 6, 2) * 1, Mid(TextToDateTimeInput, 9, 2) * 1) + _
    TimeSerial(Mid(TextToDateTimeInput, 12, 2) * 1, Mid(TextToDateTimeInput, 14, 2) * 1, Mid(TextToDateTimeInput, 16, 2) * 1)
'Debug.Print "############# " & "TextToDateTime"
'Debug.Print "TextToDateTimeInput: " & TextToDateTimeInput
'Debug.Print "TextToDateTime: " & TextToDateTime
'Debug.Print "############# " & "TextToDateTime"
End Function
