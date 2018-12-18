Attribute VB_Name = "Progress"
Option Explicit
'These are used for progress bar updates

'From https://stackoverflow.com/questions/3166265/open-an-html-page-in-default-browser-with-vba
Private Declare Function ShellExecute _
  Lib "shell32.dll" Alias "ShellExecuteA" ( _
  ByVal hWnd As Long, _
  ByVal Operation As String, _
  ByVal Filename As String, _
  Optional ByVal Parameters As String, _
  Optional ByVal Directory As String, _
  Optional ByVal WindowStyle As Long = vbMinimizedFocus _
  ) As Long
  
Public Sub OpenUrl(Hyperlink As String)
    Dim lSuccess As Long
    lSuccess = ShellExecute(0, "Open", Hyperlink)
End Sub

Sub SetBackupPgogressBarData()
'Set up user form for progress bar
    OutlookFolderCurrentCount = 0
    OutlookItemCurrentCount = 0
    OutlookItemCurrentCountToday = 0
    ProgressStartTime = Now()
    BackupBar.Show vbModeless
End Sub

Sub UpdateBackupProgressBar(OutlookFolderInput As Outlook.MAPIFolder, OutlookItemInput)
'Update user form based on progress of Outlook item backup
Dim TopBorder As Double
Dim LeftBorder As Double
Dim MarginBorder As Double
Dim MarginText As Double
Dim TextBorder As Double
Dim ItemText As String
Dim TimeToGo
    MarginBorder = 5
    MarginText = 4 'space between texts
    TextBorder = 3 'double for button
    TopBorder = 31
    LeftBorder = 14
    
    If OutlookItemCurrentCountToday = 1 Then
'        With BackupBar
'            .Top = Application.ActiveWindow.Top + 25
'            .Left = Application.ActiveWindow.Left + 25
'        End With
    Else
        ItemText = OutlookItemInput.Subject
    End If
    
'Title of user form
    With BackupBar
        .Caption = "Backing up: " & OutlookCurrentFolder 'update
    End With
'Text and size of progress bar
    With BackupBar.Header
        .Caption = Format(OutlookItemCurrentCountToday / OutlookItemCount, "0.00%") & _
            " (" & OutlookItemCurrentCountToday & " of " & OutlookItemCount & ")" 'update
    End With
    With BackupBar.HeaderFill
        .BackColor = BarColour 'RGB(106, 224, 56)  'update########################################################
        .Width = BackupBar.Header.Width * OutlookItemCurrentCountToday / OutlookItemCount 'update
    End With
'First line of details
    With BackupBar.TopRight
        .Caption = "(" & OutlookFolderCurrentCountToday & " of " & OutlookFolderCount & ")" 'update
    End With
    With BackupBar.TopLeft
        .Caption = "Folder: " & OutlookFolderInput 'update
    End With
'Second line of details
    With BackupBar.BottomRight
        .Caption = "(" & OutlookItemCurrentCountTodayInFolder & " of " & OutlookFolderInput.Items.Count & ")" 'update
    End With
    With BackupBar.BottomLeft
        .Caption = "Item: " & ItemText 'update
    End With
'Last line of details
    TimeToGo = (OutlookItemCount - OutlookItemCurrentCountToday) * (ProgressNowTime - ProgressStartTime) / OutlookItemCurrentCountToday

    If TimeToGo <= 1 Then
        TimeToGo = Format(TimeToGo, "hh:nn:ss", vbUseSystemDayOfWeek, vbUseSystem)
    Else
        TimeToGo = Format(TimeToGo, "dd hh:nn:ss", vbUseSystemDayOfWeek, vbUseSystem)
    End If

    With BackupBar.Footer
        .Caption = "Time remaining: " & TimeToGo 'update
    End With
End Sub

Sub CloseBackupProgressBar()
'Hmm, not really used apart from dummy endpoint for user form cancellation
'    Select Case EndCode
'        Case 1
'            MsgBox "Cancelled"
'        Case 2
'            MsgBox "Red cross"
'        Case Else
'            MsgBox "All done"
'    End Select
'    Unload BackupBar
End Sub

Sub UpdateHDDProgressBar(HDDFolderInput, HDDFileInput, ListOfWhat As String)
'Update user form based on progress of Log file creation
Dim TopBorder As Double
Dim LeftBorder As Double
Dim MarginBorder As Double
Dim MarginText As Double
Dim TextBorder As Double
Dim ItemText As String
Dim TimeToGo
    MarginBorder = 5
    MarginText = 4 'space between texts
    TextBorder = 3 'double for button
    TopBorder = 31
    LeftBorder = 14
'Title of user form
    With BackupBar
        .Caption = "Build List: " & ListOfWhat
    End With
'Text and size of progress bar
    With BackupBar.Header
        .Caption = Format(HDDFileCountToday / HDDFileCount, "0.00%") & _
            " (" & HDDFileCountToday & " of " & HDDFileCount & ")" 'update
    End With
    With BackupBar.HeaderFill
        .BackColor = BarColour 'RGB(106, 224, 56)  'update########################################################
        .Width = BackupBar.Header.Width * HDDFileCountToday / HDDFileCount 'update
    End With
'First line of details
    With BackupBar.TopRight
        .Caption = " " ' "(" & HDDFolderCountToday & " of " & HDDFolderCount & ")" 'update
    End With
    With BackupBar.TopLeft
        .Caption = "Folder: " & "..." & Right(HDDFolderInput, 31) 'update
    End With
'Second line of details
    With BackupBar.BottomRight
        .Caption = " " '"(" & HDDFileCountToday & " of " & HDDFileCount & ")" 'update
    End With
    With BackupBar.BottomLeft
        .Caption = "Item: " & HDDFileInput 'update
    End With
'Last line of details
    TimeToGo = (HDDFileCount - HDDFileCountToday) * (ProgressNowTime - ProgressStartTime) / HDDFileCountToday

    If TimeToGo <= 1 Then
        TimeToGo = Format(TimeToGo, "hh:nn:ss", vbUseSystemDayOfWeek, vbUseSystem)
    Else
        TimeToGo = Format(TimeToGo, "dd hh:nn:ss", vbUseSystemDayOfWeek, vbUseSystem)
    End If

    With BackupBar.Footer
        .Caption = "Time remaining: " & TimeToGo 'update
    End With
End Sub

Function BarColour()
'Wow, changing colour! I had too much time
Dim GreenStart As Double
Dim GreenEnd As Double
Dim GreenRGB As Double
Dim GreenScale As Double
Dim RedStart As Double
Dim RedEnd As Double
Dim RedRGB As Double
Dim RedScale As Double
Dim BlueStart As Double
Dim BlueEnd As Double
Dim BlueRGB As Double
Dim BlueScale As Double
Dim RGBStep As Double

    RGBStep = 50
    GreenStart = 211
    GreenEnd = 236
    GreenScale = (GreenEnd - GreenStart) / RGBStep
    RedStart = 1
    RedEnd = 102
    RedScale = (RedEnd - RedStart) / RGBStep
    BlueStart = 41
    BlueEnd = 118
    BlueScale = (BlueEnd - BlueStart) / RGBStep
    
    RGBStepCount = RGBStepCount + 1
    If RGBStepCount <= RGBStep Then
        GreenRGB = GreenStart + GreenScale * RGBStepCount
    Else
        If RGBStepCount < RGBStep * 2 Then
            GreenRGB = GreenEnd - GreenScale * (RGBStepCount - RGBStep)
        Else
            GreenRGB = GreenEnd - GreenScale * (RGBStepCount - RGBStep)
            RGBStepCount = 0
        End If
    End If
    
    If RGBStepCount <= RGBStep Then
        RedRGB = RedStart + RedScale * RGBStepCount
    Else
        If RGBStepCount < RGBStep * 2 Then
            RedRGB = RedEnd - RedScale * (RGBStepCount - RGBStep)
        Else
            RedRGB = RedEnd - RedScale * (RGBStepCount - RGBStep)
            RGBStepCount = 0
        End If
    End If
    
    If RGBStepCount <= RGBStep Then
        BlueRGB = BlueStart + BlueScale * RGBStepCount
    Else
        If RGBStepCount < RGBStep * 2 Then
            BlueRGB = BlueEnd - BlueScale * (RGBStepCount - RGBStep)
        Else
            BlueRGB = BlueEnd - BlueScale * (RGBStepCount - RGBStep)
            RGBStepCount = 0
        End If
    End If

 BarColour = RGB(RedRGB, GreenRGB, BlueRGB)
End Function


