Attribute VB_Name = "Function_Log"
Option Explicit
'These are creating and handling arrays and log files

Public fso As Scripting.FileSystemObject ' add MS scripting Runtime

Sub AddToItemArray(OutlookFolderInput As Outlook.MAPIFolder, OutlookItemInput) 'As Outlook.MailItem)
'Creates ArrayLine for Outlook item with attributes to be validated and with statuses
Dim Status As String
Dim ErrorType As String
Dim OLFolder As Outlook.MAPIFolder
Dim OLFolderTypeValid As Boolean
Dim OLItemInFolderCount As Double
Dim OLItem As Object 'Outlook.MailItem
Dim OLItemTitle As String
Dim OLItemDate As String
Dim OLItemUnRead As Boolean
Dim OLItemTitleShort As String
Dim OLItemType As String
Dim OLItemTypeValid As Boolean
Dim OLItemSize As Double
Dim OLItemSizeValid As Boolean
Dim OLItemFrom As String
Dim OLItemTo As String
Dim OLItemToCount As Double
Dim OLItemToCountValid As Boolean
Dim HDDFolderPath As String
Dim HDDItemValidLenght As Boolean
Dim FullItemPath As String
Dim CleanItemName As String
Dim i As Double
Dim ClassCheckChar As Double
    
    On Error GoTo NetworkError
    
    Set OLFolder = OutlookFolderInput
    Set OLItem = OutlookItemInput
    ClassCheckChar = 8
    
    OLFolderTypeValid = ValidOutlookFolder(OLFolder)
'Check validity
    If OLFolderTypeValid = False Then
        If ErrorType = "" Then
            ErrorType = "Invalid Folder"
        Else
            ErrorType = ErrorType & "; " & "Invalid Folder"
        End If
    End If

    OLItemInFolderCount = OLFolder.Items.Count
    HDDFolderPath = DefultBackupLocation & "\" & CleanOutlookFullPathName(OLFolder) & "\"
    OLItemTypeValid = ValidOutlookItem(OLItem)
    OLItemTitle = OLItem.Subject
    OLItemSize = OLItem.Size
    OLItemType = OLItem.MessageClass 'Class
    
    Select Case Left(OLItem.MessageClass, ClassCheckChar)
        Case Left("IPM.Appointment", ClassCheckChar) 'Appointment
'Debug.Print "RecurrenceState: " & OLItem.RecurrenceState
            OLItemType = "Appointment"
            OLItemDate = Format(OLItem.Start, "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
                Format(OLItem.Start, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem)
            OLItemUnRead = OLItem.UnRead
            OLItemFrom = OLItem.Organizer
            OLItemToCount = OLItem.Recipients.Count
            If OLItemToCount = 0 Then
                OLItemTo = "-"
            Else
                If OLItemToCount > MaxItemTo Then
                    For i = 1 To MaxItemTo
                        If i = 1 Then
                            OLItemTo = OLItem.Recipients(i)
                        Else
                            OLItemTo = OLItemTo & ", " & OLItem.Recipients(i)
                        End If
                    Next
                Else
                    For i = 1 To OLItemToCount
                        If i = 1 Then
                            OLItemTo = OLItem.Recipients(i)
                        Else
                            OLItemTo = OLItemTo & ", " & OLItem.Recipients(i)
                        End If
                    Next
                End If
            End If
        Case Left("IPM.Note", ClassCheckChar) 'Mail
            OLItemType = "Mail"
            OLItemDate = Format(OLItem.ReceivedTime, "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
                Format(OLItem.ReceivedTime, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem)
            OLItemUnRead = OLItem.UnRead
            OLItemFrom = OLItem.SenderName
            OLItemToCount = OLItem.Recipients.Count
            If OLItemToCount = 0 Then
                OLItemTo = "-"
            Else
                If OLItemToCount > MaxItemTo Then
                    For i = 1 To MaxItemTo
                        If i = 1 Then
                            OLItemTo = OLItem.Recipients(i)
                        Else
                            OLItemTo = OLItemTo & ", " & OLItem.Recipients(i)
                        End If
                    Next
                Else
                    For i = 1 To OLItemToCount
                        If i = 1 Then
                            OLItemTo = OLItem.Recipients(i)
                        Else
                            OLItemTo = OLItemTo & ", " & OLItem.Recipients(i)
                        End If
                    Next
                End If
            End If
        Case Left("IPM.Schedule.Meeting.Resp.Tent", ClassCheckChar), _
            Left("IPM.Schedule.Meeting.Resp.Pos", ClassCheckChar), _
            Left("IPM.Schedule.Meeting.Resp.Neg", ClassCheckChar), _
            Left("IPM.Schedule.Meeting.Request", ClassCheckChar), _
            Left("IPM.Schedule.Meeting.Canceled", ClassCheckChar) 'Meeting
            OLItemType = "Meeting"
            OLItemDate = Format(OLItem.ReceivedTime, "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
                Format(OLItem.ReceivedTime, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem)
            OLItemUnRead = OLItem.UnRead
            OLItemFrom = OLItem.SenderName
            OLItemToCount = OLItem.Recipients.Count
            If OLItemToCount = 0 Then
                OLItemTo = "-"
            Else
                If OLItemToCount > MaxItemTo Then
                    For i = 1 To MaxItemTo
                        If i = 1 Then
                            OLItemTo = OLItem.Recipients(i)
                        Else
                            OLItemTo = OLItemTo & ", " & OLItem.Recipients(i)
                        End If
                    Next
                Else
                    For i = 1 To OLItemToCount
                        If i = 1 Then
                            OLItemTo = OLItem.Recipients(i)
                        Else
                            OLItemTo = OLItemTo & ", " & OLItem.Recipients(i)
                        End If
                    Next
                End If
            End If
        Case Left("IPM.StickyNote", ClassCheckChar) 'Note
            OLItemType = "Note"
            OLItemDate = Format(OLItem.LastModificationTime, "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
                Format(OLItem.LastModificationTime, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem)
        Case Left("IPM.Task", ClassCheckChar) 'Task
            OLItemType = "Task"
            OLItemDate = Format(OLItem.StartDate, "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
                Format(OLItem.StartDate, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem)
            OLItemUnRead = OLItem.UnRead
            OLItemFrom = OLItem.Owner
            OLItemToCount = OLItem.Recipients.Count
            If OLItemToCount = 0 Then
                OLItemTo = "-"
            Else
                If OLItemToCount > MaxItemTo Then
                    For i = 1 To MaxItemTo
                        If i = 1 Then
                            OLItemTo = OLItem.Recipients(i)
                        Else
                            OLItemTo = OLItemTo & ", " & OLItem.Recipients(i)
                        End If
                    Next
                Else
                    For i = 1 To OLItemToCount
                        If i = 1 Then
                            OLItemTo = OLItem.Recipients(i)
                        Else
                            OLItemTo = OLItemTo & ", " & OLItem.Recipients(i)
                        End If
                    Next
                End If
            End If
        Case Else
            OLItemTypeValid = False
    End Select
'Check validity
    If OLItemTypeValid = False Then
        If ErrorType = "" Then
            ErrorType = "Invalid Item"
        Else
            ErrorType = ErrorType & "; " & "Invalid Item"
        End If
        GoTo NotValid:
    Else
    End If
    
    CleanItemName = ReplaceIllegalCharsFileFolderName(OLItemTitle, ReplaceCharBy, MaxFileNameLenght, False)
    FullItemPath = HDDFolderPath & OLItemDate & " - " & CleanItemName & ".msg"
    If Len(FullItemPath) > MaxPathLenght Then
        If MaxPathLenght - Len(HDDFolderPath) - Len(OLItemDate & " - " & ".msg") < MinFileNameLenght Then
            OLItemTitleShort = "-"
            HDDItemValidLenght = False
        Else
            CleanItemName = ReplaceIllegalCharsFileFolderName(OLItemTitle, ReplaceCharBy, _
                MaxPathLenght - Len(HDDFolderPath) - Len(OLItemDate & " - " & ".msg"), False)
            OLItemTitleShort = OLItemDate & " - " & CleanItemName & ".msg"
            HDDItemValidLenght = True
        End If
    Else
        OLItemTitleShort = OLItemDate & " - " & CleanItemName & ".msg"
        HDDItemValidLenght = True
    End If
'Check validity
    If HDDItemValidLenght = False Then
        If ErrorType = "" Then
            ErrorType = "File Name lenght"
        Else
            ErrorType = ErrorType & "; " & "File Name lenght"
        End If
    Else
    End If

    If OLItemSize > MaxItemSize Then
        OLItemSizeValid = False
    Else
        OLItemSizeValid = True
    End If
'Check validity
    If OLItemSizeValid = False Then
        If ErrorType = "" Then
            ErrorType = "Email size"
        Else
            ErrorType = ErrorType & "; " & "Email size"
        End If
    Else
    End If
    
    If OLItemToCount > MaxItemTo Then
        OLItemToCountValid = False
    Else
        OLItemToCountValid = True
    End If
'Check validity
    If OLItemToCountValid = False Then
        If ErrorType = "" Then
            ErrorType = "Number of recipients"
        Else
            ErrorType = ErrorType & "; " & "Number of recipients"
        End If
    Else
    End If
        
NotValid:
    If ErrorType = "" Then
        Status = "Listed"
        ErrorType = "OK"
    Else
        Status = "Error"
    End If

    ItemArray = Array(Status, ErrorType, _
        OLFolder, OLFolderTypeValid, OLItemInFolderCount, _
        OLItemTitle, OLItemDate, OLItemUnRead, OLItemFrom, OLItemTo, OLItemTitleShort, _
        OLItemType, OLItemTypeValid, OLItemSize, OLItemSizeValid, _
        OLItemToCount, OLItemToCountValid, _
        HDDFolderPath, HDDItemValidLenght)
        
NetworkError:
If Err.Number <> 0 And Left(Err.Description, Len("Network")) = "Network" Then
    MsgBox "Hmmm..., AddToItemArray"
End If
'Debug.Print "############# " & "AddToItemArray"
'Debug.Print "OutlookFolderInput: " & OutlookFolderInput
'Debug.Print "OutlookItemInput: " & OutlookItemInput
'Debug.Print "Status: " & Status
'Debug.Print "ErrorType: " & ErrorType
'Debug.Print "OLFolder: " & OLFolder
'Debug.Print "OLFolderTypeValid: " & OLFolderTypeValid
'Debug.Print "OLItemInFolderCount: " & OLItemInFolderCount
'Debug.Print "OLItemTitle: " & OLItemTitle
'Debug.Print "OLItemDate: " & OLItemDate
'Debug.Print "OLItemFrom: " & OLItemFrom
'Debug.Print "OLItemTo: " & OLItemTo
'Debug.Print "OLItemTitleShort: " & OLItemTitleShort
'Debug.Print "OLItemType: " & OLItemType
'Debug.Print "OLItemTypeValid: " & OLItemTypeValid
'Debug.Print "OLItemSize: " & OLItemSize
'Debug.Print "OLItemSizeValid: " & OLItemSizeValid
'Debug.Print "OLItemToCount: " & OLItemToCount
'Debug.Print "OLItemToCountValid: " & OLItemToCountValid
'Debug.Print "HDDFolderPath: " & HDDFolderPath
'Debug.Print "HDDItemValidLenght: " & HDDItemValidLenght
'Debug.Print "############# " & "AddToItemArray"
'    For i = LBound(ItemArray) To UBound(ItemArray)
'    Debug.Print ItemArray(i)
'    Next
'Debug.Print "############# " & "AddToItemArray"
End Sub

Sub AddToShortItemArray(OutlookFolderInput As Outlook.MAPIFolder, OutlookItemInput) 'As Outlook.MailItem)
'Creates Short ArrayLine for Outlook item with attributes to be checked against saved files
Dim OLFolder As Outlook.MAPIFolder
Dim OLItem As Object 'Outlook.MailItem
Dim OLItemTitle As String
Dim OLItemDate As String
Dim OLItemTitleShort As String
Dim OLItemType As String
Dim HDDFolderPath As String
Dim CleanItemName As String
Dim i As Double
Dim ClassCheckChar As Double

    On Error GoTo NetworkError
    
    Set OLFolder = OutlookFolderInput
    Set OLItem = OutlookItemInput
    ClassCheckChar = 8
    
    OLItemTitle = OLItem.Subject
    OLItemType = OLItem.MessageClass 'Class
    HDDFolderPath = DefultBackupLocation & "\" & CleanOutlookFullPathName(OLFolder) & "\"
    
    Select Case Left(OLItem.MessageClass, ClassCheckChar)
        Case Left("IPM.Appointment", ClassCheckChar) 'Appointment
'Debug.Print "RecurrenceState: " & OLItem.RecurrenceState
            OLItemDate = Format(OLItem.Start, "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
                Format(OLItem.Start, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem)
        Case Left("IPM.Note", ClassCheckChar) 'Mail
            OLItemDate = Format(OLItem.ReceivedTime, "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
                Format(OLItem.ReceivedTime, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem)
        Case Left("IPM.Schedule.Meeting.Resp.Tent", ClassCheckChar), _
            Left("IPM.Schedule.Meeting.Resp.Pos", ClassCheckChar), _
            Left("IPM.Schedule.Meeting.Resp.Neg", ClassCheckChar), _
            Left("IPM.Schedule.Meeting.Request", ClassCheckChar), _
            Left("IPM.Schedule.Meeting.Canceled", ClassCheckChar) 'Meeting
            OLItemDate = Format(OLItem.ReceivedTime, "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
                Format(OLItem.ReceivedTime, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem)
        Case Left("IPM.StickyNote", ClassCheckChar) 'Note
            OLItemDate = Format(OLItem.LastModificationTime, "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
                Format(OLItem.LastModificationTime, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem)
        Case Left("IPM.Task", ClassCheckChar) 'Task
            OLItemDate = Format(OLItem.StartDate, "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
                Format(OLItem.StartDate, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem)
        Case Else
    End Select
    
    CleanItemName = ReplaceIllegalCharsFileFolderName(OLItemTitle, ReplaceCharBy, MaxFileNameLenght, False) & ".msg"

    ItemShortArray = Array(OLItemDate, CleanItemName, HDDFolderPath)
    
NetworkError:
If Err.Number <> 0 And Left(Err.Description, Len("Network")) = "Network" Then
    MsgBox "Hmmm..., AddToShortItemArray"
End If
'Debug.Print "############# " & "AddToShortItemArray"
'Debug.Print "OutlookFolderInput: " & OutlookFolderInput
'Debug.Print "OutlookItemInput: " & OutlookItemInput
'Debug.Print "OLItemType: " & OLItemType
'Debug.Print "OLItemDate: " & OLItemDate
'Debug.Print "OLItemTitleShort: " & OLItemTitleShort
'Debug.Print "HDDFolderPath: " & HDDFolderPath
'Debug.Print "############# " & "AddToShortItemArray"
'    For i = LBound(ItemShortArray) To UBound(ItemShortArray)
'       Debug.Print ItemShortArray(i)
'    Next
'Debug.Print "############# " & "AddToShortItemArray"
End Sub

Sub LogFileOpen(Filename As String)
'Open the selected file
Dim Where As String
    FileNumber = FreeFile
    Where = Filename
    Open Where For Append Access Write As #FileNumber
'Debug.Print "############# " & "LogFileOpen"
'Debug.Print "FileName: " & FileName
'Debug.Print "############# " & "LogFileOpen"
End Sub

Sub LogFileCreateWithHeading(Filename As String, Heading As Variant)
'Open the selected file or creates it if doesn't exit with header
Dim Rows As Double
Dim Columns As Double
Dim ArrayValue As String
Dim WholeLine As String
Dim Where As String
Dim ArrayHeading As Variant
    Set fso = New Scripting.FileSystemObject
    FileNumber = FreeFile

    Where = Filename
    ReDim ArrayHeading(UBound(Heading))
    ArrayHeading = Heading
    
    If fso.FileExists(Filename) = False Then
        FileNumber = FreeFile
        Where = Filename
        Open Where For Output Access Write As #FileNumber
        WholeLine = ""
        For Rows = LBound(ArrayHeading) To UBound(ArrayHeading)
            ArrayValue = ArrayHeading(Rows)
            If WholeLine = "" Then
                WholeLine = ArrayValue
            Else
                WholeLine = WholeLine & SeparatorInFile & ArrayValue
            End If
        Next Rows
        Print #FileNumber, WholeLine
        Close #FileNumber
    End If
'Debug.Print "############# " & "LogFileCreateWithHeading"
'Debug.Print "FileName: " & FileName
'Dim i As Double
'    For i = LBound(ArrayHeading) To UBound(ArrayHeading)
'        Debug.Print "Heading" & i & ": " & Heading(i)
'    Next
'Debug.Print "############# " & "LogFileCreateWithHeading"
End Sub

Sub LogFileAddLine(ArrayToAdd As Variant)
'Add an array as line to opened text file
Dim Rows As Double
Dim Columns As Double
Dim ArrayValue As String
Dim WholeLine As String
Dim Where As String
Dim ArrayLine As Variant

    WholeLine = ""
    ReDim ArrayLine(UBound(ArrayToAdd))
    ArrayLine = ArrayToAdd

    For Columns = LBound(ArrayLine) To UBound(ArrayLine)
        ArrayValue = ArrayLine(Columns)
        If WholeLine = "" Then
            WholeLine = ArrayValue
        Else
            WholeLine = WholeLine & SeparatorInFile & ArrayValue
        End If
    Next Columns
    Print #FileNumber, WholeLine
'Debug.Print "############# " & "LogFileAddLine"
'Debug.Print "WholeLine: " & WholeLine
'Debug.Print "############# " & "LogFileAddLine"
End Sub

Sub LogFileClose()
'Close an open txt file
    Close #FileNumber
'Debug.Print "############# " & "LogFileClose"
End Sub

Sub BuildHDDArray(Optional FolderToCheck As String, Optional Filename As String)
'Build initial log file of already saved Outlook items
Dim What As String
Dim Where As String

    Call WipeMeClean
    Call SetConfig

    If Filename = Empty Then
        Where = DefultBackupLocationLog & "\" & LogFileSum & ".txt"
    Else
        Where = Filename
    End If

    If FolderToCheck = Empty Then
        What = DefultBackupLocation '& "\"
    Else
        What = FolderToCheck
    End If

    Call SetBackupPgogressBarData
    Call HDDFolderItemCount(What)
    Call HDDItemToArray(What)
    Call RebuildLogFile
    Unload BackupBar

'Debug.Print "############# " & "BuildHDDArray"
'Debug.Print "FolderToCheck: " & FolderToCheck
'Debug.Print "FileName: " & Filename
'Debug.Print "UBound(ArchivedFileArray, 2): " & UBound(ArchivedFileArray, 2)
'Debug.Print "HDDFolderCount: " & HDDFolderCount
'Debug.Print "HDDFileCount: " & HDDFileCount
'Debug.Print "############# " & "BuildHDDArray"
End Sub

Sub HDDItemToArray(HDDFolderInput As Variant)
'Loop through folders (and subfolders) and files from selected folder and append to HDD array
Dim FolderLoop As Scripting.Folder
Dim SubFolderLoop As Scripting.Folder
Dim HDDFolder As Variant
Dim HDDFile As Variant
Dim HDDFileName As String
Dim HDDDate As String
Dim HDDSubject As String

Dim i As Double
    Set fso = New Scripting.FileSystemObject
    Set FolderLoop = fso.GetFolder(HDDFolderInput)
    If fso.FolderExists(FolderLoop) = True Then
        HDDFolderCountToday = HDDFolderCountToday + 1
        Set HDDFolder = FolderLoop.Files
        For Each HDDFile In HDDFolder
            HDDFileCountToday = HDDFileCountToday + 1
            HDDFileName = HDDFile.Name
            If Len(HDDFileName) > 21 And IsNumeric(Left(HDDFileName, 2)) = True Then
                HDDDate = Left(HDDFileName, 17) 'TextToDateTime(HDDFileName)
                HDDSubject = Mid(HDDFileName, 21, 9999)
                FileArray = Array(HDDDate, HDDSubject, FolderLoop)
                Call AddToHDDArray(FileArray)
                ProgressNowTime = Now()
                Call UpdateHDDProgressBar(FolderLoop, HDDFileName, "Create array of existsing files")
                DoEvents
            End If
        Next
    Else
    End If
'Process all folders and subfolders recursively
    If FolderLoop.SubFolders.Count Then
       For Each SubFolderLoop In FolderLoop.SubFolders
           Call HDDItemToArray(SubFolderLoop)
       Next
    End If
'Debug.Print "############# " & "HDDItemToArray"
'Debug.Print "HDDFolderInput: " & HDDFolderInput
'Debug.Print "UBound(ArchivedFileArray, 1): " & UBound(ArchivedFileArray, 1)
'Debug.Print "############# " & "HDDItemToArray"
End Sub

Sub AddToHDDArray(NewArrayLine As Variant)
'Add 'ArrayLine' to HDD array (horizontal)
Dim NewLine As Variant
Dim Row As Double
Dim Col As Double
Dim TempArray As Variant
Dim ArrayRowSize As Double
Dim ArrayColumnSize As Double
    
    NewLine = NewArrayLine
    ArrayRowSize = UBound(NewLine)
    If IsEmpty(ArchivedFileArray) Then
        ReDim ArchivedFileArray(UBound(FileArrayHeading), 1)
        For Row = 0 To ArrayRowSize
            ArchivedFileArray(Row, Col) = FileArrayHeading(Row)
        Next
    Else
        ArrayColumnSize = UBound(ArchivedFileArray, 2)
        ReDim Preserve ArchivedFileArray(ArrayRowSize, ArrayColumnSize + 1)
    End If
    ArrayColumnSize = UBound(ArchivedFileArray, 2)
    For Row = 0 To ArrayRowSize
        ArchivedFileArray(Row, ArrayColumnSize) = NewLine(Row)
    Next
'Debug.Print "############# " & "AddToHDDArray"
'Debug.Print "UBound(NewLine, 2): " & UBound(NewLine, 2)
'Debug.Print "UBound(ArchivedFileArray, 2): " & UBound(ArchivedFileArray, 2)
'Debug.Print "############# " & "AddToLogArray"
'Dim i As Double
'Dim j As Double
'    For i = LBound(ArchivedFileArray, 2) To UBound(ArchivedFileArray, 2)
'        Debug.Print "                                             New Line"
'        For j = LBound(ArchivedFileArray, 1) To UBound(ArchivedFileArray, 1)
'            Debug.Print ArchivedFileArray(j, i)
'        Next
'    Next
'Debug.Print "############# " & "AddToHDDArray"
End Sub

Sub LogHDDFileInOneVertical(Filename As String)
'Creates a text log file with date and headings and a list of files saved (vertical log file from horizontal array)
'Overwrites existing file
Dim Rows As Double
Dim Columns As Double
Dim ArrayValue As String
Dim WholeLine As String
Dim Where As String
Dim FileNumber As Double
    
    FileNumber = FreeFile
    Where = Filename
    Open Where For Output Access Write As #FileNumber
    
    HDDFileCount = UBound(ArchivedFileArray, 2) + 1
    For Columns = LBound(ArchivedFileArray, 2) To UBound(ArchivedFileArray, 2)
        WholeLine = ""
        For Rows = LBound(ArchivedFileArray, 1) To UBound(ArchivedFileArray, 1)
            ArrayValue = ArchivedFileArray(Rows, Columns)
            If WholeLine = "" Then
                WholeLine = ArrayValue
            Else
                WholeLine = WholeLine & SeparatorInFile & ArrayValue
            End If
        Next Rows
        Print #FileNumber, WholeLine
        HDDFileCountToday = Columns + 1
        ProgressNowTime = Now()
        Call UpdateHDDProgressBar(Where, WholeLine, "Save log file")
        DoEvents
    Next Columns
    Close #FileNumber
'Debug.Print "############# " & "LogHDDFileInOneVertical"
'Debug.Print "Where: " & Where
'Debug.Print "############# " & "LogHDDFileInOneVertical"
End Sub

Sub ReadHDDInAsArray(Filename As String) 'FileName As String)
'Reads log file into array (vertical file into horizontal array)
Dim Where As String
Dim WholeLine As String
Dim NewLine As Variant
Dim Row As Double
Dim Col As Double
Dim RowStart As Double
Dim RowEnd As Double
Dim ColStart As Double
Dim ColEnd As Double
Dim TempArray As Variant
Dim ArrayRowSize As Double
Dim ArrayColumnSize As Double

    FileNumber = FreeFile
    Where = Filename
    
    ArrayRowSize = UBound(FileArrayHeading)
    ReDim NewLine(ArrayRowSize)
    
    Open Where For Input As #FileNumber
    Do Until EOF(1)
        Line Input #FileNumber, WholeLine

        For Col = 0 To ArrayRowSize
            If Col = 0 Then
                ColStart = 1
            Else
                ColStart = ColEnd + 1
            End If
            If Col = ArrayRowSize Then
                ColEnd = Len(WholeLine) + 1
            Else
                ColEnd = InStr(ColStart, WholeLine, SeparatorInFile)
            End If
            NewLine(Col) = Mid(WholeLine, ColStart, ColEnd - ColStart)
        Next

        If IsEmpty(ArchivedFileArray) Then
            ReDim ArchivedFileArray(UBound(FileArrayHeading), 1)
            For Row = 0 To ArrayRowSize
                ArchivedFileArray(Row, 0) = FileArrayHeading(Row)
            Next
        Else
            ArrayColumnSize = UBound(ArchivedFileArray, 2)
            ReDim Preserve ArchivedFileArray(ArrayRowSize, ArrayColumnSize + 1)
        End If
        ArrayColumnSize = UBound(ArchivedFileArray, 2)
        For Row = 0 To ArrayRowSize
            ArchivedFileArray(Row, ArrayColumnSize) = NewLine(Row)
        Next
'Dim i As Double
'Dim j As Double
'        i = 1 + i
'        Debug.Print "                                                              New Line"
'        For j = LBound(ArchivedFileArray, 1) To UBound(ArchivedFileArray, 1)
'            Debug.Print ArchivedFileArray(j, i)
'        Next
    Loop
    Close #FileNumber
'Debug.Print "ReadIn: end: " & Now()
'Debug.Print "############# " & "ReadHDDInAsArray"
'Debug.Print "UBound(ArchivedFileArray, 1): " & UBound(ArchivedFileArray, 1)
'Debug.Print "UBound(ArchivedFileArray, 2): " & UBound(ArchivedFileArray, 2)
'Debug.Print "############# " & "ReadHDDInAsArray"
'Dim i As Double
'Dim j As Double
'    For i = LBound(ArchivedFileArray, 2) To UBound(ArchivedFileArray, 2)
'        Debug.Print "                                                              New Line"
'        For j = LBound(ArchivedFileArray, 1) To UBound(ArchivedFileArray, 1)
'            Debug.Print ArchivedFileArray(j, i)
'        Next
'    Next
'Debug.Print "############# " & "ReadHDDInAsArray"
End Sub

Sub RebuildLogFile(Optional Filename As String)
'Reads log file into array sort it by date and resave
Dim i As Long
Dim Where As String
Dim TempArray As Variant
Dim NewLine As Variant
Dim Row As Double
Dim Col As Double
   
    If IsEmpty(ArchivedFileArray) Then
        Call WipeMeClean
        Call SetConfig
        If Filename = Empty Then
            Where = DefultBackupLocationLog & "\" & LogFileSum & ".txt"
        Else
            Where = Filename
        End If
        Call SetBackupPgogressBarData
        Call ReadHDDInAsArray(Where)
    Else
        If Filename = Empty Then
            Where = DefultBackupLocationLog & "\" & LogFileSum & ".txt"
        Else
            Where = Filename
        End If
    End If
    
    Call QuickSort2(ArchivedFileArray, 1, 0)

    For i = UBound(ArchivedFileArray, 2) To 1 Step -1
        If IsNumeric(Left(ArchivedFileArray(0, i), 1)) = False Then
            ReDim Preserve ArchivedFileArray(UBound(ArchivedFileArray, 1), UBound(ArchivedFileArray, 2) - 1)
        Else
            GoTo ReSaveFile
        End If
    Next

ReSaveFile:
    TempArray = ArchivedFileArray
    ArchivedFileArray = Empty

    ReDim NewLine(UBound(TempArray, 1))
'    Debug.Print UBound(NewLine)

    For Col = LBound(TempArray, 2) To UBound(TempArray, 2)
        For Row = LBound(TempArray, 1) To UBound(TempArray, 1)
             NewLine(Row) = TempArray(Row, Col)
        Next
        Call AddToHDDArray(NewLine)
    Next
    
    Call LogHDDFileInOneVertical(Where)
    Unload BackupBar
    Call WipeMeClean
'Debug.Print "############# " & "RebuildLogFile"
'Debug.Print "UBound(ArchivedFileArray, 1): " & UBound(ArchivedFileArray, 1)
'Debug.Print "UBound(ArchivedFileArray, 2): " & UBound(ArchivedFileArray, 2)
'Debug.Print "############# " & "RebuildLogFile"
'Dim i As Double
'Dim j As Double
'    For i = LBound(ArchivedFileArray, 2) To UBound(ArchivedFileArray, 2)
'        Debug.Print "                                                              New Line"
'        For j = LBound(ArchivedFileArray, 1) To UBound(ArchivedFileArray, 1)
'            Debug.Print ArchivedFileArray(j, i)
'        Next
'    Next
'Debug.Print "############# " & "RebuildLogFile"
End Sub

Sub LogLastItemChecked(Filename As String)
'Create file with a date of item last checked
Dim Where As String
Dim WholeLine As String
    FileNumber = FreeFile
    Where = Filename
    WholeLine = LastItemCheckedDate
    Open Where For Output Access Write As #FileNumber
    Print #FileNumber, WholeLine
    Close #FileNumber
'Debug.Print "############# " & LogLastItemChecked
'Debug.Print "FileName: " & FileName
'Debug.Print "LastItemCheckedDate: " & LastItemCheckedDate
'Debug.Print "############# " & LogLastItemChecked
End Sub

Sub ReadLastItemDateLog(Filename As String)
'Reads last item date log file and sets LastItemCheckedDate
Dim Where As String
Dim WholeLine As String
    Set fso = New Scripting.FileSystemObject
    FileNumber = FreeFile
    Where = Filename
    If fso.FileExists(Where) Then
        Open Where For Input As #FileNumber
        Do Until EOF(1)
            Line Input #FileNumber, WholeLine
            LastItemCheckedDate = WholeLine
        Loop
        Close #FileNumber
    End If
'Debug.Print "############# " & "ReadLastItemDateLog"
'Debug.Print "FileName: " & FileName
'Debug.Print "LastItemCheckedDate: " & LastItemCheckedDate
'Debug.Print "############# " & "ReadLastItemDateLog"
End Sub

Sub AddToShortItemDate(OutlookFolderInput As Outlook.MAPIFolder, OutlookItemInput) 'As Outlook.MailItem)
'Creates ShortDateArray for Outlook item
Dim OLFolder As Outlook.MAPIFolder
Dim OLItem As Object 'Outlook.MailItem
Dim OLItemDate As String
Dim OLItemType As String
Dim ClassCheckChar As Double
Dim i As Double

    On Error GoTo NetworkError
    
    Set OLFolder = OutlookFolderInput
    Set OLItem = OutlookItemInput
    ClassCheckChar = 8

    OLItemType = OLItem.MessageClass 'Class

    Select Case Left(OLItem.MessageClass, ClassCheckChar)
        Case Left("IPM.Appointment", ClassCheckChar) 'Appointment
'Debug.Print "RecurrenceState: " & OLItem.RecurrenceState
            OLItemDate = OLItem.Start
        Case Left("IPM.Note", ClassCheckChar) 'Mail
            OLItemDate = Format(OLItem.ReceivedTime, "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
                Format(OLItem.ReceivedTime, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem)
        Case Left("IPM.Schedule.Meeting.Resp.Tent", ClassCheckChar), _
            Left("IPM.Schedule.Meeting.Resp.Pos", ClassCheckChar), _
            Left("IPM.Schedule.Meeting.Resp.Neg", ClassCheckChar), _
            Left("IPM.Schedule.Meeting.Request", ClassCheckChar), _
            Left("IPM.Schedule.Meeting.Canceled", ClassCheckChar) 'Meeting
            OLItemDate = Format(OLItem.ReceivedTime, "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
                Format(OLItem.ReceivedTime, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem)
        Case Left("IPM.StickyNote", ClassCheckChar) 'Note
            OLItemDate = Format(OLItem.LastModificationTime, "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
                Format(OLItem.LastModificationTime, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem)
        Case Left("IPM.Task", ClassCheckChar) 'Task
            OLItemDate = Format(OLItem.StartDate, "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
                Format(OLItem.StartDate, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem)
        Case Else
    End Select

    ItemDateOnly = TextToDateTime(OLItemDate)
    
NetworkError:
If Err.Number <> 0 And Left(Err.Description, Len("Network")) = "Network" Then
    MsgBox "Hmmm..., AddToShortItemDateArray"
End If
'Debug.Print "############# " & "AddToShortItemDateArray"
'Debug.Print "UBound(ItemShortArray): " & UBound(ItemShortArray)
'Debug.Print "############# " & "AddToShortItemDateArray"
End Sub
