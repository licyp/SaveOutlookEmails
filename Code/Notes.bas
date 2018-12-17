Attribute VB_Name = "Notes"
'############################
'check spelling in word
'chske all private variabales are used
'DONE add short array for date subject checekc
'DONE i now it is fully DRY, room for improvemenet
'Make a Pull ReqDONE uest 
'http://makeapullrequest.com/


'log file cration progress bar: red to arrayc and save to file!!!!




Sub AddToLogArray(NewArrayLine)
'Add 'ArrayLine' to main log array
Dim NewLine As Variant
Dim Row As Double
Dim Col As Double
Dim TempArray As Variant
Dim ArrayRowSize As Double
Dim ArrayColumnSize As Double

'    ArrayColumnSize = UBound(ItemArray)
'    NewLine = NewArrayLine
'    If IsEmpty(LogArray) Then
'        ReDim LogArray(0, UBound(ItemArray))
'    Else
'        ArrayRowSize = UBound(LogArray, 1)
'        ReDim TempArray(ArrayRowSize, ArrayColumnSize)
'        For Row = 0 To ArrayRowSize
'            For Col = 0 To ArrayColumnSize
'                TempArray(Row, Col) = LogArray(Row, Col)
'            Next
'        Next
'
'        ReDim LogArray(ArrayRowSize + 1, ArrayColumnSize)
'        For Row = 0 To ArrayRowSize
'            For Col = 0 To ArrayColumnSize
'                LogArray(Row, Col) = TempArray(Row, Col)
'            Next
'        Next
'    End If
'    ArrayRowSize = UBound(LogArray, 1)
'    For Col = 0 To ArrayColumnSize
'        LogArray(ArrayRowSize, Col) = NewLine(Col)
'    Next
'Debug.Print "############# " & "AddToLogArray"
'Debug.Print "UBound(LogArray, 1): " & UBound(LogArray, 1)
'Debug.Print "############# " & "AddToLogArray"
'Dim i As Double
'Dim j As Double
'    For i = LBound(LogArray, 1) To UBound(LogArray, 1)
'        Debug.Print "New Line"
'        For j = LBound(LogArray, 2) To UBound(LogArray, 2)
'            Debug.Print LogArray(i, j)
'        Next
'    Next
'Debug.Print "############# " & "AddToLogArray"
End Sub

Sub LogFileLastLine(FileName As String)
'Counts number of lines in selected text file
Dim Where As String
Dim WholeLine As String
    FileNumber = FreeFile

    Where = FileName
On Error GoTo NoFile
    Open Where For Input As #FileNumber
    Do Until EOF(1)
        Line Input #FileNumber, WholeLine
        NoOfLinesInFile = NoOfLinesInFile + 1
    Loop
    Close #FileNumber
    If ForceResave = True Then
        GoTo NoFile
    End If
    Exit Sub

NoFile:
    NoOfLinesInFile = 0
'Debug.Print "############# " & "LogFileLastLine"
End Sub


Sub LogFileInOne(FileName As String)
'Creates a text log file with date and headdins and a list of items processed
'http://www.cpearson.com/excel/ImpText.aspx
Dim Rows As Double
Dim Columns As Double
Dim ArrayValue As String
Dim WholeLine As String
Dim Where As String
Dim FileNumber As Double

    FileNumber = FreeFile

    Where = DefultBackupLocationLog & "\" & FileName & " " & _
        Format(Now(), "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
        Format(Now(), "hhnnss", vbUseSystemDayOfWeek, vbUseSystem) & ".txt"
    Open Where For Output Access Write As #FileNumber

'    Call AddToLogArray(ItemArray)
    
    For Rows = LBound(LogArray, 1) To UBound(LogArray, 1)
        WholeLine = ""
        For Columns = LBound(LogArray, 2) To UBound(LogArray, 2)
            ArrayValue = LogArray(Rows, Columns)
            If WholeLine = "" Then
                WholeLine = ArrayValue
            Else
                WholeLine = WholeLine & SeparatorInFile & ArrayValue
            End If
        Next Columns
        Print #FileNumber, WholeLine
    Next Rows
    Close #FileNumber
    Erase LogArray
'Debug.Print "############# " & "LogFileInOne"
'Debug.Print "Where: " & Where
'Debug.Print "############# " & "LogFileInOne"
End Sub

Sub LogHDDFileInOneHorizontal() 'Horizontal
'Creates a text log file with date and headdins and a list of files saved
Dim Rows As Double
Dim Columns As Double
Dim ArrayValue As String
Dim WholeLine As String
Dim Where As String
Dim FileNumber As Double
    
    FileNumber = FreeFile

    Where = DefultBackupLocationLog & "\" & "HDDH.txt" ' & " " & _
        Format(Now(), "yyyy.mm.dd", vbUseSystemDayOfWeek, vbUseSystem) & "-" & _
        Format(Now(), "hhnnss", vbUseSystemDayOfWeek, vbUseSystem) & ".txt"
    Open Where For Output Access Write As #FileNumber
    
    For Rows = LBound(ArchivedFileArray, 1) To UBound(ArchivedFileArray, 1)
        WholeLine = ""
        For Columns = LBound(ArchivedFileArray, 2) To UBound(ArchivedFileArray, 2)
            ArrayValue = ArchivedFileArray(Rows, Columns)
            If WholeLine = "" Then
                WholeLine = ArrayValue
            Else
                WholeLine = WholeLine & SeparatorInFile & ArrayValue
            End If
        Next Columns
        Print #FileNumber, WholeLine
    Next Rows
    Close #FileNumber
'Debug.Print "############# " & "LogHDDFileInOneHorizontal"
'Debug.Print "Where: " & Where
'Debug.Print "############# " & "LogHDDFileInOneHorizontal"
End Sub


'####################
'Outlook:
'Progress
'Startup
'Archive
'
'Select folder
'
'lomitation (object tile:email, calndar note etc
'max email size, max no of to/cc/bc)
'
'Count all itmes under
'    all/ quliafied (by limitataion type)
'
'loop folders
'
'onHDD create folder if dosn;t exist
'
'save meassgae (by limitioan, size, no of addresses)
'[over ride or skip if file exist]
'
'Email save naming
'Folder naming
''
'DONE file name limits
'DONE File shorntening
'select saved location+ waring if path >=50 char
'force resave option
'save file: date create as email date-time' from ' to title and other exif
'addreport: to draft email to self: saved, failed etc
'#####################

'Save emails from longed-in user's Outlook acccounts (active, acrhcived or choosed)onto loged-in user Desktop under 'eMails' folder (dafault) _
'unless other location is selected selected (recomneded to leave on desktop as other locatacion may reach Windos maxinm path lenght) (warning if reached 50 char)
'Emails saved under 'Outlook account' folder:
'   received and sent emails saved at startup or sent/ receved evenet under 'active' Outlook folder
'   when email moved from 'active' to 'archavied' Outlook foler and backup run then email will be saved again under 'archived' Outlook folder
'   when choosen folder is used to back up emails then emails will b esaved under the relevant Outlook folder
'It is sugsteste to tur on auto back-up od 'active' folder and run bacuup on 'archive' only once (it may take long time)
'when back-up is restarted saved emilas will NOT be resaved (defult settings) unless 'Force resave' selected
'Illegal characters (*, /, \, |, ?, :, ", %, <, >, non-printables) replaced by '_' in folder and file names
'Maximum lenght of folder and file names is 100 charaters
'backround (backup location is assumed no longer than 50ch, account folder name no longer than 50 cha and folder strauate within account folder 50 ch,_
'whihc leaved 110 char. the date prefix is 21ch making email file name 100 cha as safe option as windows max limit is 245)
'Backed-up fodlers
'Inbox, sent, draft, calandar event
'archaibd by EVS
'indovual folder names are choped down to 100 characters, resonalbe shoul be no motre than 50 warning loged in log
'emai file name lenght is the leftover afert pathnmae, if it cannot be at laeast 25 char not saved but recoded in the log





