'https://docs.microsoft.com/en-us/office/vba/api/overview/outlook
Option Explicit

Public fso As Scripting.FileSystemObject ' add MS scripting Runtime
Public HDDMainFolder As Scripting.Folder
Public HDDSubFolder As Scripting.Folder

Public OutlookAccountFolder As Outlook.MAPIFolder
Public OutlookMainFolder As Outlook.MAPIFolder
Public OutlookCurrentFolder As Outlook.MAPIFolder
Public OutlookSubFolder As Outlook.MAPIFolder

Public LastFoundWasAt As Double

Public OutlookFolderCount As Double
Public OutlookItemCount As Double
Public HDDFolderCount As Double
Public HDDFileCount As Double
Public HDDFileCountToday As Double
Public HDDFolderCountToday As Double
Public SavedAlreadyCounter As Double
Public SavedAlreadyFisrt As Double
Public SavedAlreadyCurrent As Double
Public OverlapDays As Double
Public OverlapResaved As Double
Public OverlapSubject As Double

Public DefultBackupLocation As String
Public DefultBackupLocationLog As String
Public InvalidFolders As Variant
Public ValidItems As Variant
Public ArchivedArray As Variant
Public DefultFolder As String
Public LogedinUserFileLocation As String
Public MaxFolderNameLenght As Double
Public MaxFileNameLenght As Double
Public MinFileNameLenght As Double
Public MaxPathLenght As Double
Public MaxItemTo As Double
Public MaxItemSize As Double
Public ReplaceCharBy As String
Public SuffixText As String
Public ForceResave As Boolean
Public SaveResult As String
Public FileNumber As Double
Public SaveItemsToHDD As Boolean
Public LogFileFlow As String
Public LogFileSum As String
Public ErrorSkipCode As String

Public OutlookFolderCurrentCount  As Double
Public OutlookItemCurrentCount  As Double
Public OutlookItemCurrentCountToday  As Double
Public OutlookItemCurrentCountTodayInFolder  As Double
Public OutlookFolderCurrentCountToday As Double
Public ProgressStartTime  As Double
Public ProgressNowTime  As Double
Public EndCode As Double
Public RGBStepCount As Double
Public NoOfLinesInFile As Double
Public SeparatorInFile As String
Public OutlookItemSavedAlready As Boolean

Public LogArray As Variant
Public ItemArray As Variant
Public ItemShortArray As Variant
Public FileArray As Variant
Public ArchivedLogArray As Variant
Public ArchivedFileArray As Variant
Public EndMessage As Boolean
Public FromNewToOld As Boolean
Public AutoRun As Boolean
Public ItemArrayHeading As Variant
Public FileArrayHeading As Variant
Public LinkToGitHub As String
Public UndeliverableError As String
Public DateError As Boolean

Sub WipeMeClean()
'Cleans variables (in case of previous unfinished runs)
    Set HDDMainFolder = Nothing
    Set HDDSubFolder = Nothing
    Set OutlookAccountFolder = ActiveOutlookAccount
    Set OutlookMainFolder = Nothing
    Set OutlookCurrentFolder = Nothing
    Set OutlookSubFolder = Nothing
    
    OutlookFolderCount = 0
    OutlookItemCount = 0
    HDDFolderCount = 0
    HDDFileCount = 0
    EndCode = 0
    SavedAlreadyCounter = 0
    LastFoundWasAt = 0
    LogArray = Empty
    ItemArray = Empty
    ArchivedLogArray = Empty
    ArchivedFileArray = Empty
'Debug.Print "############# " & "WipeMeClean"
End Sub

Sub SetConfig()
'Sets basic boundaries
    Set fso = New Scripting.FileSystemObject
    
    SuffixText = "..."
    ReplaceCharBy = "_"
    SeparatorInFile = Chr(9)
    MaxFolderNameLenght = 100
    MaxFileNameLenght = 200
    MinFileNameLenght = 40
    MaxPathLenght = 240
    MaxItemTo = 250
    MaxItemSize = 25000000 '25MB
    OverlapDays = 7
    OverlapResaved = 100
    OverlapSubject = 20

    DefultFolder = "Desktop\eMails"
    LogedinUserFileLocation = CStr(Environ("USERPROFILE"))
    DefultBackupLocation = LogedinUserFileLocation & "\" & DefultFolder
    DefultBackupLocationLog = DefultBackupLocation & "\" & "Logs"
    SaveItemsToHDD = True
    ForceResave = False
    LogFileSum = "Log of Saved Outlook Items"
    LinkToGitHub = "https://github.com/licyp/SaveOutlookEmails"

'https://docs.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
'Name    Value   Folder Name Description
'olFolderConflicts   19  Conflicts   The Conflicts folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
'olFolderContacts    10  Contacts    The Contacts folder.
'olFolderDeletedItems    3   Deleted Items   The Deleted Items folder.
'olFolderJournal 11  Journal The Journal folder.
'olFolderJunk    23  Junk E-Mail The Junk E-Mail folder.
'olFolderLocalFailures   21  Local Failures  The Local Failures folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
'olFolderRssFeeds    25  RSS Feeds   The RSS Feeds folder.
'olFolderServerFailures  22  Server Failures The Server Failures folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
'olFolderSuggestedContacts   30  Suggested Contacts  The Suggested Contacts folder.
'olFolderSyncIssues  20  Sync Issues The Sync Issues folder. Only available for an Exchange account.
'olFolderManagedEmail    29      The top-level folder in the Managed Folders group. For more information on Managed Folders, see the Help in Microsoft Outlook. Only available for an Exchange account.
'olPublicFoldersAllPublicFolders 18      The All Public Folders folder in the Exchange Public Folders store. Only available for an Exchange account.
'olFolderCalendar    9   Calendar    The Calendar folder.
'olFolderDrafts  16  Drafts  The Drafts folder.
'olFolderInbox   6   Inbox   The Inbox folder.
'olFolderNotes   12  Notes   The Notes folder.
'olFolderOutbox  4   Outbox  The Outbox folder.
'olFolderSentMail    5   Sent Mail   The Sent Mail folder.
'olFolderTasks   13  Tasks   The Tasks folder.
'olFolderToDo    28  To Do   The To Do folder.
    
    InvalidFolders = Array("Conflicts", "Contacts", "Journal", _
        "Junk E-Mail", "Local Failures", "RSS Feeds", "Server Failures", _
        "Suggested Contacts", "Sync Issues", "Recipient Cache") '"Deleted Items",
    ValidItems = Array("IPM.Appointment", "IPM.Schedule", "IPM.Note", "IPM.Task", "IPM.StickyNote") ' Start with
    ArchivedArray = Array("EAS") ' End with
    ItemArrayHeading = Array("Backup Status", "Error", _
        "Fodler", "Fodler Validity", "Item Count", _
        "Title", "Date", "Unread", "From", "To", "Shortened Title", _
        "Type", "Type Validity", "Size", "Size Validity", _
        "Recipients Count", "Recipients Validity", _
        "Path on Drive", "Path Validity")
    FileArrayHeading = Array("Date", "Subject", "Path")
    
    Call CreateHDDFolder(DefultBackupLocationLog)
'    If fso.FileExists(DefultBackupLocationLog & "\" & LogFileSum & ".txt") = False Then
'        Call BuildHDDArray(DefultBackupLocation, DefultBackupLocationLog & "\" & LogFileSum & ".txt")
'    End If
        
'Debug.Print "############# " & "SetConfig"
End Sub

Sub QuickAccessSaveEmails()
'Sub to be used on QuickLunch
    EndMessage = True
    FromNewToOld = False
    AutoRun = False
    Call BackUpOutlookFolder(AutoRun)
End Sub

Sub BackUpOutlookFolder(Optional AutoRun As Boolean)
    Call WipeMeClean
    Call SetConfig
Dim MsgBoxTitle As String
Dim MsgBoxButtons As String
Dim MsgBoxText As String
Dim MsgBoxResponse As Double

If AutoRun = True Then
    MsgBoxResponse = 6
    GoTo BackUpMainAccount
End If
'Set folder to back up
StartAgain:
    MsgBoxButtons = vbYesNoCancel + vbQuestion + vbDefaultButton1
    MsgBoxTitle = "Backup Outlook Folder"
    MsgBoxText = "Back up '" & OutlookAccountFolder & "' folder instead?"

    Set OutlookCurrentFolder = Outlook.Application.Session.PickFolder
    If OutlookCurrentFolder Is Nothing Then
        MsgBoxResponse = MsgBox(MsgBoxText, MsgBoxButtons, MsgBoxTitle)

BackUpMainAccount:
        Select Case MsgBoxResponse
        Case 6 'Yes
            Set OutlookCurrentFolder = OutlookAccountFolder
        Case 7 'No
            GoTo StartAgain
        Case 2 'Cancel
            Call WipeMeClean
            Exit Sub
        End Select
    Else
    End If
    
'Is chosen folder valid folder for backup?
    MsgBoxButtons = vbOKOnly + vbExclamation + vbDefaultButton1
    MsgBoxTitle = "Backup Outlook Folder"
    MsgBoxText = "Selected '" & OutlookCurrentFolder & "' folder is not valid for backup."
    If ValidOutlookFolder(OutlookCurrentFolder) = False Then
        MsgBox MsgBoxText, MsgBoxButtons, MsgBoxTitle
        GoTo StartAgain
    Else
    End If

    Call SetBackupPgogressBarData
    Set OutlookMainFolder = TopOutlookFolder(OutlookCurrentFolder)
    
    Call OutlookFolderItemCount(OutlookCurrentFolder)
    Call CreateHDDFolderForOutlookFolder(OutlookCurrentFolder)
    Call HDDFolderItemCount(DefultBackupLocation & "\" & CleanOutlookFullPathName(OutlookCurrentFolder))
    Call LogFileCreateWithHeading(DefultBackupLocationLog & "\" & LogFileSum & ".txt", FileArrayHeading)
    Call ReadHDDInAsArray(DefultBackupLocationLog & "\" & LogFileSum & ".txt")
    Call LogFileOpen(DefultBackupLocationLog & "\" & LogFileSum & ".txt")
    Call LoopOutlookFolders(OutlookCurrentFolder, SaveItemsToHDD)
    Call LogFileClose
    
    Select Case EndCode
        Case 1
            MsgBox "Cancelled"
        Case 2
            MsgBox "Red cross"
        Case Else
            If EndMessage = False Then
            Else
                MsgBox "All done"
            End If
    End Select

    Call WipeMeClean
    Unload BackupBar
    
'Debug.Print "############# " & "BackUpOutlookFolder"
'Debug.Print "OutlookAccountFolder: " & OutlookAccountFolder
'Debug.Print "MsgBoxResponse: " & MsgBoxResponse
'Debug.Print "OutlookCurrentFolder: " & OutlookCurrentFolder
'Debug.Print "OutlookMainFolder: " & OutlookMainFolder
'Debug.Print "FullPathOutlookFolder: " & FullPathOutlookFolder(OutlookCurrentFolder)
'Debug.Print "CleanOutlookFullPathName: " & CleanOutlookFullPathName(OutlookCurrentFolder)
'Debug.Print "OutlookFolderCount: " & OutlookFolderCount
'Debug.Print "OutlookItemCount: " & OutlookItemCount
'Debug.Print "DefultBackupLocation: " & DefultBackupLocation
'Debug.Print "HDDFolderCount: " & HDDFolderCount
'Debug.Print "HDDFileCount: " & HDDFileCount
'Debug.Print "############# " & "BackUpOutlookFolder"
End Sub

Sub LoopOutlookFolders(LoopOutlookFoldersInput As Outlook.MAPIFolder, SaveItem As Boolean)
Dim FolderLoop As Outlook.MAPIFolder
Dim SubFolderLoop As Outlook.MAPIFolder
Dim i As Double
        
    Set FolderLoop = LoopOutlookFoldersInput
    If ValidOutlookFolder(FolderLoop) = True Then
        OutlookFolderCurrentCountToday = OutlookFolderCurrentCountToday + 1
        If FromNewToOld = False Then
            For i = 3000 To FolderLoop.Items.Count 'from old to new
                DateError = False
'Checks autorun overlap and stop scanning Outlook folder
                If ForceResave = False And AutoRun = True And _
                    SavedAlreadyCounter > OverlapResaved And Abs(SavedAlreadyCurrent - SavedAlreadyFisrt) > OverlapDays Then
                    Exit Sub
                End If
'If cancelled or red cross exist then stop
                If EndCode = 1 Or EndCode = 2 Then
                    Exit Sub
                End If
'Check file existence
                Call AddToShortItemArray(FolderLoop, FolderLoop.Items(i))
                Call FileExistsInLogOrHDD
                If OutlookItemSavedAlready = True And ForceResave = False Then
                    SaveResult = "Saved Already"
                    SavedAlreadyCounter = SavedAlreadyCounter + 1
                    If DateError = False Then
                        If SavedAlreadyCounter = 1 Then
                            SavedAlreadyFisrt = TextToDateTime(ItemShortArray(0) & " ")
                        Else
                            SavedAlreadyCurrent = TextToDateTime(ItemShortArray(0) & " ")
                        End If
                    End If
                Else
                    SavedAlreadyCounter = 0
'Read full Outlook item data
                    Call AddToItemArray(FolderLoop, FolderLoop.Items(i))
                    If SaveItem = True And ItemArray(1) = "OK" Then
'Save item
                        Call SaveOutlookItem(FolderLoop.Items(i), ForceResave, ItemArray)
                        ItemArray(0) = SaveResult
                        If ErrorSkipCode <> "" Then
                            ItemArray(1) = ErrorSkipCode
                        Else
'Add successful file save to log file
                            Call LogFileAddLine(ItemShortArray)
                        End If
                    End If
                End If
                OutlookItemCurrentCountTodayInFolder = i
                OutlookItemCurrentCountToday = OutlookItemCurrentCountToday + 1
'Update progress bar
                ProgressNowTime = Now()
                Call UpdateBackupProgressBar(FolderLoop, FolderLoop.Items(i))
                DoEvents
'Debug.Print "SavedAlreadyCounter: " & SavedAlreadyCounter
            Next
        Else
            For i = FolderLoop.Items.Count To 1 Step -1 'new old to old
                DateError = False
'Checks autorun overlap and stop scanning Outlook folder
                If ForceResave = False And AutoRun = True And _
                    SavedAlreadyCounter > OverlapResaved And Abs(SavedAlreadyCurrent - SavedAlreadyFisrt) > OverlapDays Then
                    Exit Sub
                End If
'If cancelled or red cross exist then stop
                If EndCode = 1 Or EndCode = 2 Then
                    Exit Sub
                End If
'Check file existence
                Call AddToShortItemArray(FolderLoop, FolderLoop.Items(i))
                Call FileExistsInLogOrHDD
                If OutlookItemSavedAlready = True And ForceResave = False Then
                    SaveResult = "Saved Already"
                    SavedAlreadyCounter = SavedAlreadyCounter + 1
                    If DateError = False Then
                        If SavedAlreadyCounter = 1 Then
                            SavedAlreadyFisrt = TextToDateTime(ItemShortArray(0) & " ")
                        Else
                            SavedAlreadyCurrent = TextToDateTime(ItemShortArray(0) & " ")
                        End If
                    End If
                Else
                    SavedAlreadyCounter = 0
'Read full Outlook item data
                    Call AddToItemArray(FolderLoop, FolderLoop.Items(i))
                    If SaveItem = True And ItemArray(1) = "OK" Then
'Save item
                        Call SaveOutlookItem(FolderLoop.Items(i), ForceResave, ItemArray)
                        ItemArray(0) = SaveResult
                        If ErrorSkipCode <> "" Then
                            ItemArray(1) = ErrorSkipCode
                        Else
'Add successful file save to log file
                            Call LogFileAddLine(ItemShortArray)
                        End If
                    End If
                End If
                OutlookItemCurrentCountTodayInFolder = i
                OutlookItemCurrentCountToday = OutlookItemCurrentCountToday + 1
'Update progress bar
                ProgressNowTime = Now()
                Call UpdateBackupProgressBar(FolderLoop, FolderLoop.Items(i))
                DoEvents
'Debug.Print "SavedAlreadyCounter: " & SavedAlreadyCounter
            Next
        End If
    Else
    End If
'Process all folders and subfolders recursively
    If FolderLoop.Folders.Count Then
       For Each SubFolderLoop In FolderLoop.Folders
           Call LoopOutlookFolders(SubFolderLoop, SaveItem)
       Next
    End If
'Debug.Print "############# " & "LoopOutlookFolders"
'Debug.Print "LoopOutlookFoldersInput: " & LoopOutlookFoldersInput
'Debug.Print "FolderLoop.Items(i): " & FolderLoop.Items(i).Subject
'Debug.Print "UBound(LogArray, 1): " & UBound(LogArray, 1)
'Debug.Print "SavedAlreadyCounter: " & SavedAlreadyCounter
'Debug.Print "############# " & "LoopOutlookFolders"
End Sub

Sub SaveOutlookItem(OutlookItemInput, ReSave As Boolean, ItemData)
'Saves Outlook item if not exists or force resave=true; other attributes are used from ItemArray related to selected Outlook item
Dim OutlookApp As Outlook.Application
Dim ObjectInspector As Outlook.Inspector
Dim ItemToBeSaved As Object
Dim ItemToBeSavedOpen As Object
Dim UnRead As Boolean
Dim ItemStatus As String
Dim SaveFileName As String
Dim SavePathName As String
Dim FileExists As Boolean
Dim ArchivedItem As Boolean

    ErrorSkipCode = ""
    Set OutlookApp = Outlook.Application
    Set ItemToBeSaved = OutlookItemInput
    Set fso = New Scripting.FileSystemObject
    UnRead = ItemData(7)
    ItemStatus = ItemData(0)
    SavePathName = ItemData(17)
    SaveFileName = ItemData(10)
    ArchivedItem = ArchivedOutlookItem(ItemToBeSaved)

    If ItemStatus = "Error" Then
        SaveResult = ItemStatus
    Else
        If ReSave = True Then
            If ArchivedItem = True Then
                Set ObjectInspector = Nothing
                ItemToBeSaved.Display
                Do While ObjectInspector Is Nothing
                    Set ObjectInspector = OutlookApp.ActiveInspector
                Loop
                Set ItemToBeSavedOpen = ObjectInspector.CurrentItem
            Else
                Set ItemToBeSavedOpen = ItemToBeSaved
            End If
            ItemToBeSavedOpen.SaveAs SavePathName & SaveFileName, olMSG
            If UnRead = True Then
                ItemToBeSavedOpen.UnRead = True
            End If
            If ArchivedItem = True Then
                ItemToBeSavedOpen.Close olDiscard
            Else
            End If
            SaveResult = "Resaved"
        Else
            If ArchivedItem = True Then
                Set ObjectInspector = Nothing
                ItemToBeSaved.Display
                Do While ObjectInspector Is Nothing
                    Set ObjectInspector = OutlookApp.ActiveInspector
                Loop
                Set ItemToBeSavedOpen = ObjectInspector.CurrentItem
            Else
                Set ItemToBeSavedOpen = ItemToBeSaved
            End If
On Error GoTo SkipError
            ItemToBeSavedOpen.SaveAs SavePathName & SaveFileName, olMSG
            If UnRead = True Then
                ItemToBeSavedOpen.UnRead = True
            End If
            If ArchivedItem = True Then
                ItemToBeSavedOpen.Close olDiscard
            Else
            End If
            SaveResult = "Saved"
        End If
    End If
    
SkipError:
If Err.Number <> 0 Then
    ErrorSkipCode = Err.Number & " " & Err.Description
    SaveResult = "Error"
End If
    Set OutlookApp = Nothing
    Set ObjectInspector = Nothing
    Set fso = Nothing
    Set ItemToBeSaved = Nothing
    Set ItemToBeSavedOpen = Nothing
'Debug.Print "############# " & "SaveOutlookItem"
'Debug.Print "OutlookItemInput: " & OutlookItemInput
'Debug.Print "ReSave: " & ReSave
'Debug.Print "ItemStatus: " & ItemStatus
'Debug.Print "UnRead: " & UnRead
'Debug.Print "SaveResult: " & SaveResult
'Debug.Print SavePathName & SaveFileName
'Debug.Print "############# " & "SaveOutlookItem"
End Sub
