Option Explicit
'These are used for progress bar updates

'From https://stackoverflow.com/questions/3166265/open-an-html-page-in-default-browser-with-vba
Private Declare Function ShellExecute _
  Lib "shell32.dll" Alias "ShellExecuteA" ( _
  ByVal hWnd As Long, _
  ByVal Operation As String, _
  ByVal File_Name As String, _
  Optional ByVal Parameters As String, _
  Optional ByVal Directory As String, _
  Optional ByVal WindowStyle As Long = vbMinimizedFocus _
  ) As Long

Public Sub Open_Url(Hyperlink As String)
    Dim lSuccess As Long
    lSuccess = ShellExecute(0, "Open", Hyperlink)
End Sub

Sub Set_Backup_Progress_Bar_Data()
'Set up user form for progress bar
    Outlook_Folder_Current_Count = 0
    Outlook_Item_Current_Count = 0
    Outlook_Item_Current_Count_Today = 0
    Progress_Start_Time = Now()
    BackupBar.Show vbModeless
End Sub

Sub Update_Backup_Progress_Bar(Outlook_Folder_Input As Outlook.MAPIFolder, Outlook_Item_Input)
'Update user form based on progress of Outlook item backup
Dim Top_Border As Double
Dim Left_Border As Double
Dim Margin_Border As Double
Dim Margin_Text As Double
Dim Text_Border As Double
Dim Item_Text As String
Dim Time_To_Go
    Margin_Border = 5
    Margin_Text = 4 'space between texts
    Text_Border = 3 'double for button
    Top_Border = 31
    Left_Border = 14

    If Outlook_Item_Current_Count_Today = 1 Then
'        With BackupBar
'            .Top = Application.ActiveWindow.Top + 25
'            .Left = Application.ActiveWindow.Left + 25
'        End With
    Else
        Item_Text = Outlook_Item_Input.Subject
    End If
    
    If Outlook_Item_Current_Count_Today = 0 Then
        Outlook_Item_Current_Count_Today = 1
    End If

'Title of user form
    With BackupBar
        .Caption = "Backing up: " & Outlook_Current_Folder 'update
    End With
'Text and size of progress bar
    With BackupBar.Header
        .Caption = Format(Outlook_Item_Current_Count_Today / Outlook_Item_Count, "0.00%") & _
            " (" & Outlook_Item_Current_Count_Today & " of " & Outlook_Item_Count & ")" 'update
    End With
    With BackupBar.HeaderFill
        .BackColor = Bar_Colour 'RGB(106, 224, 56)  'update########################################################
        .Width = BackupBar.Header.Width * Outlook_Item_Current_Count_Today / Outlook_Item_Count 'update
    End With
'First line of details
    With BackupBar.TopRight
        .Caption = "(" & Outlook_Folder_Current_Count_Today & " of " & Outlook_Folder_Count & ")" 'update
    End With
    With BackupBar.TopLeft
        .Caption = "Folder: " & Outlook_Folder_Input 'update
    End With
'Second line of details
    With BackupBar.BottomRight
        .Caption = "(" & Outlook_Item_Current_Count_Today_In_Folder & " of " & Outlook_Folder_Input.Items.Count & ")" 'update
    End With
    With BackupBar.BottomLeft
        .Caption = "Item: " & Item_Text 'update
    End With
'Last line of details
    Time_To_Go = (Outlook_Item_Count - Outlook_Item_Current_Count_Today) * (Progress_Now_Time - Progress_Start_Time) / Outlook_Item_Current_Count_Today

    If Time_To_Go <= 1 Then
        Time_To_Go = Format(Time_To_Go, "hh:mm:ss", vbUseSystemDayOfWeek, vbUseSystem)
    Else
        Time_To_Go = Format(Time_To_Go, "dd hh:mm:ss", vbUseSystemDayOfWeek, vbUseSystem)
    End If

    With BackupBar.Footer
        .Caption = "Time remaining: " & Time_To_Go 'update
    End With
End Sub

Sub Close_Backup_Progress_Bar()
'Hmm, not really used apart from dummy endpoint for user form cancellation
'    Select Case End_Code
'        Case 1
'            MsgBox "Cancelled"
'        Case 2
'            MsgBox "Red cross"
'        Case Else
'            MsgBox "All done"
'    End Select
'    Unload BackupBar
End Sub

Sub Update_HDD_Progress_Bar(HDD_Folder_Input, HDD_File_Input, List_Of_What As String)
'Update user form based on progress of Log file creation
Dim Top_Border As Double
Dim Left_Border As Double
Dim Margin_Border As Double
Dim Margin_Text As Double
Dim Text_Border As Double
Dim Item_Text As String
Dim Time_To_Go
    Margin_Border = 5
    Margin_Text = 4 'space between texts
    Text_Border = 3 'double for button
    Top_Border = 31
    Left_Border = 14
'Title of user form
    With BackupBar
        .Caption = "Build List: " & List_Of_What
    End With
'Text and size of progress bar
    With BackupBar.Header
        .Caption = Format(HDD_File_Count_Today / HDD_File_Count, "0.00%") & _
            " (" & HDD_File_Count_Today & " of " & HDD_File_Count & ")" 'update
    End With
    With BackupBar.HeaderFill
        .BackColor = Bar_Colour 'RGB(106, 224, 56)  'update########################################################
        .Width = BackupBar.Header.Width * HDD_File_Count_Today / HDD_File_Count 'update
    End With
'First line of details
    With BackupBar.TopRight
        .Caption = " " ' "(" & HDD_Folder_Count_Today & " of " & HDD_Folder_Count & ")" 'update
    End With
    With BackupBar.TopLeft
        .Caption = "Folder: " & "..." & Right(HDD_Folder_Input, 31) 'update
    End With
'Second line of details
    With BackupBar.BottomRight
        .Caption = " " '"(" & HDD_File_Count_Today & " of " & HDD_File_Count & ")" 'update
    End With
    With BackupBar.BottomLeft
        .Caption = "Item: " & HDD_File_Input 'update
    End With
'Last line of details
    Time_To_Go = (HDD_File_Count - HDD_File_Count_Today) * (Progress_Now_Time - Progress_Start_Time) / HDD_File_Count_Today

    If Time_To_Go <= 1 Then
        Time_To_Go = Format(Time_To_Go, "hh:mm:ss", vbUseSystemDayOfWeek, vbUseSystem)
    Else
        Time_To_Go = Format(Time_To_Go, "dd hh:mm:ss", vbUseSystemDayOfWeek, vbUseSystem)
    End If

    With BackupBar.Footer
        .Caption = "Time remaining: " & Time_To_Go 'update
    End With
End Sub

Function Bar_Colour()
'Wow, changing colour! I had too much time
Dim Green_Start As Double
Dim Green_End As Double
Dim Green_RGB As Double
Dim Green_Scale As Double
Dim Red_Start As Double
Dim Red_End As Double
Dim Red_RGB As Double
Dim Red_Scale As Double
Dim Blue_Start As Double
Dim Blue_End As Double
Dim Blue_RGB As Double
Dim Blue_Scale As Double
Dim RGB_Step As Double

    RGB_Step = 50
    Green_Start = 211
    Green_End = 236
    Green_Scale = (Green_End - Green_Start) / RGB_Step
    Red_Start = 1
    Red_End = 102
    Red_Scale = (Red_End - Red_Start) / RGB_Step
    Blue_Start = 41
    Blue_End = 118
    Blue_Scale = (Blue_End - Blue_Start) / RGB_Step

    RGB_Step_Count = RGB_Step_Count + 1
    If RGB_Step_Count <= RGB_Step Then
        Green_RGB = Green_Start + Green_Scale * RGB_Step_Count
    Else
        If RGB_Step_Count < RGB_Step * 2 Then
            Green_RGB = Green_End - Green_Scale * (RGB_Step_Count - RGB_Step)
        Else
            Green_RGB = Green_End - Green_Scale * (RGB_Step_Count - RGB_Step)
            RGB_Step_Count = 0
        End If
    End If

    If RGB_Step_Count <= RGB_Step Then
        Red_RGB = Red_Start + Red_Scale * RGB_Step_Count
    Else
        If RGB_Step_Count < RGB_Step * 2 Then
            Red_RGB = Red_End - Red_Scale * (RGB_Step_Count - RGB_Step)
        Else
            Red_RGB = Red_End - Red_Scale * (RGB_Step_Count - RGB_Step)
            RGB_Step_Count = 0
        End If
    End If

    If RGB_Step_Count <= RGB_Step Then
        Blue_RGB = Blue_Start + Blue_Scale * RGB_Step_Count
    Else
        If RGB_Step_Count < RGB_Step * 2 Then
            Blue_RGB = Blue_End - Blue_Scale * (RGB_Step_Count - RGB_Step)
        Else
            Blue_RGB = Blue_End - Blue_Scale * (RGB_Step_Count - RGB_Step)
            RGB_Step_Count = 0
        End If
    End If

 Bar_Colour = RGB(Red_RGB, Green_RGB, Blue_RGB)
End Function
