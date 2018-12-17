Attribute VB_Name = "Function_Validation"
Option Explicit
'These are the used validations

Public fso As Scripting.FileSystemObject ' add MS scripting Runtime

Function ValidOutlookFolder(ValidOutlookFolderInput As Outlook.MAPIFolder) As Boolean
'Checks Outlook folder validity based on invalid folders defined in Config sub
Dim FolderName As String
Dim i As Double
Dim NumberOfInvalidFolders As Double
    NumberOfInvalidFolders = UBound(InvalidFolders) - LBound(InvalidFolders)
    FolderName = ValidOutlookFolderInput.Name
    ValidOutlookFolder = True
    For i = 0 To NumberOfInvalidFolders
        If UCase(FolderName) = UCase(InvalidFolders(i)) Then
            ValidOutlookFolder = False
        End If
    Next
'Debug.Print "############# " & "ValidOutlookFolder"
'Debug.Print "ValidOutlookFolderInput: " & ValidOutlookFolderInput
'Debug.Print "ValidOutlookFolder: " & ValidOutlookFolder
'Debug.Print "############# " & "ValidOutlookFolder"
End Function

Function ValidOutlookItem(ValidOutlookItemInput) As Boolean 'As Outlook.MailItem)
'Checks Outlook item validity based on valid items defined in Config sub
Dim ItemName As String
Dim i As Double
Dim NumberOfValidItems As Double
    NumberOfValidItems = UBound(ValidItems) - LBound(ValidItems)
    ItemName = ValidOutlookItemInput.MessageClass
    ValidOutlookItem = False
    For i = 0 To NumberOfValidItems
        If UCase(Left(ItemName, Len(ValidItems(i)))) = UCase(ValidItems(i)) Then
            ValidOutlookItem = True
        End If
    Next
'Debug.Print "############# " & "ValidOutlookItem"
'Debug.Print "ValidOutlookItemInput: " & ValidOutlookItemInput
'Debug.Print "ValidOutlookItem: " & ValidOutlookItem
'Debug.Print "############# " & "ValidOutlookItem"
End Function

Function ArchivedOutlookItem(ArchivedOutlookItemInput) As Boolean 'As Outlook.MailItem)
'Checks Archived Outlook item based on archived items defined in Config sub
Dim ItemName As String
Dim i As Double
Dim NumberOfArchivedItems As Double
    NumberOfArchivedItems = UBound(ArchivedArray) - LBound(ArchivedArray)
    ItemName = ArchivedOutlookItemInput.MessageClass
    ArchivedOutlookItem = False
    For i = 0 To NumberOfArchivedItems
        If UCase(Right(ItemName, Len(ArchivedArray(i)))) = UCase(ArchivedArray(i)) Then
            ArchivedOutlookItem = True
        End If
    Next
'Debug.Print "############# " & "ValidOutlookItem"
'Debug.Print "ArchivedOutlookItemInput: " & ArchivedOutlookItemInput
'Debug.Print "ValidOutlookItem: " & ValidOutlookItem
'Debug.Print "############# " & "ValidOutlookItem"
End Function

Sub FileExistsInLogOrHDD()
'Checks HDD for selected Outlook item (exact match) or log file (date time and partial subject match)
Dim i As Double
Dim ItemDate As String
Dim LogDate As String

    Set fso = New Scripting.FileSystemObject
    OutlookItemSavedAlready = False
    If AutoRun = False Then
        If LastFoundWasAt = 0 Then
            If FromNewToOld = False Then
                LastFoundWasAt = LBound(ArchivedFileArray, 2)
            Else
                LastFoundWasAt = UBound(ArchivedFileArray, 2)
            End If
        End If
    Else
        OutlookItemSavedAlready = fso.FileExists(ItemShortArray(2) & ItemShortArray(0) & " - " & ItemShortArray(1))
        Exit Sub
    End If
    
    ItemDate = TextToDateTime(ItemShortArray(0) & " ")
    
Debug.Print "Looking for: " & ItemDate
    If FromNewToOld = False Then
        For i = LastFoundWasAt To UBound(ArchivedFileArray, 2)
            If ArchivedFileArray(0, i) <> FileArrayHeading(0) Then
                LogDate = TextToDateTime(ArchivedFileArray(0, i) & " ")
            Else
                LogDate = ArchivedFileArray(0, i)
            End If
Debug.Print "Is it this one? " & ArchivedFileArray(0, i)
            If IsDate(LogDate) Then
                If DateValue(LogDate) = DateValue(ItemDate) And _
                    TimeValue(LogDate) = TimeValue(ItemDate) Then
                    If Left(ReplaceIllegalCharsFileFolderName(ItemShortArray(1) & " ", ReplaceCharBy, MaxFileNameLenght, False), OverlapSubject) _
                        = Left(ArchivedFileArray(2, i), OverlapSubject) Then
                        OutlookItemSavedAlready = True
                        LastFoundWasAt = i
                        Exit For
                    End If
                End If
            End If
        Next
    Else
        For i = LastFoundWasAt To LBound(ArchivedFileArray, 2) Step -1
Debug.Print "Is it this one? " & ArchivedFileArray(1, i)
            If IsDate(ArchivedFileArray(1, i)) Then
                If DateValue(ArchivedFileArray(1, i)) = DateValue(TextToDateTime(ItemShortArray(0) & " ")) And _
                    TimeValue(ArchivedFileArray(1, i)) = TimeValue(TextToDateTime(ItemShortArray(0) & " ")) Then
                    If Left(ReplaceIllegalCharsFileFolderName(ItemShortArray(1) & " ", ReplaceCharBy, MaxFileNameLenght, False), OverlapSubject) _
                        = Left(ArchivedFileArray(2, i), OverlapSubject) Then
                        OutlookItemSavedAlready = True
                        LastFoundWasAt = i
                        Exit For
                    End If
                End If
            End If
        Next
    End If
Debug.Print "############# " & "FileExistsInLogOrHDD"
Debug.Print "OutlookItem: " & ItemShortArray(2) & ItemShortArray(0) & " - " & ItemShortArray(1)
Debug.Print "OutlookItemSavedAlready: " & OutlookItemSavedAlready
Debug.Print "############# " & "FileExistsInLogOrHDD"
End Sub



