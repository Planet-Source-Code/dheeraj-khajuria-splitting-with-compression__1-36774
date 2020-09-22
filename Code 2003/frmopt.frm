VERSION 5.00
Begin VB.Form frmopt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "General Options"
   ClientHeight    =   5565
   ClientLeft      =   4830
   ClientTop       =   3330
   ClientWidth     =   4935
   Icon            =   "frmopt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Compression"
      Height          =   735
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton Command1 
         Caption         =   "O&K"
         Height          =   372
         Left            =   3360
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame Frame5 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4932
      Begin VB.Frame Frame3 
         Height          =   2535
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   3375
         Begin VB.CheckBox Check4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Delete the Split files after Joining"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   1080
            Width           =   2655
         End
         Begin VB.CheckBox chkOpenExplorer 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Open Explorer Window to Split Directory"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   1320
            Value           =   1  'Checked
            Width           =   2655
         End
         Begin VB.CheckBox chkDeleteParts 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "(Delete) Orignal Splitting  files "
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Width           =   2415
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Self Joining Batch File"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   9
            ToolTipText     =   "Can Split up  127 files"
            Top             =   480
            Width           =   2055
         End
         Begin VB.CheckBox crc32 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "Check CRC"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Value           =   2  'Grayed
            Width           =   1335
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "#"
            Height          =   195
            Left            =   1560
            TabIndex        =   21
            Top             =   2160
            Width           =   90
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Compressed Bytes:"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   2160
            Width           =   1350
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "#"
            Height          =   195
            Left            =   1080
            TabIndex        =   19
            Top             =   1800
            Width           =   90
         End
         Begin VB.Label Label1 
            Caption         =   "Total Bytes:"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   1800
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Split"
         Height          =   1935
         Left            =   120
         TabIndex        =   13
         Top             =   3480
         Width           =   3375
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   940
            Width           =   1095
         End
         Begin VB.ComboBox cmbSize 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmopt.frx":08CA
            Left            =   1320
            List            =   "frmopt.frx":08CC
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   940
            Width           =   1815
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   280
            Left            =   120
            TabIndex        =   1
            Top             =   480
            Width           =   3015
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000007&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   1560
            Width           =   1455
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000007&
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Segment  Size"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   1020
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Each Segment Size:"
            Height          =   195
            Left            =   1680
            TabIndex        =   15
            Top             =   1320
            Width           =   1440
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. of Segment:"
            Height          =   240
            Left            =   120
            TabIndex        =   14
            Top             =   1320
            Width           =   1185
         End
      End
   End
End
Attribute VB_Name = "frmopt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////
' FrmOpt.frm
' Splits and joins files, with optional compression
' Original by Dheeraj Khajuria Copyright 2002
'
'//////////////////////////////////////////////
Option Explicit
Private Const CB_SHOWDROPDOWN = &H14F
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Private Sub Command1_Click()
Dim X As Integer
Dim sMsg As String, sMsgDis As Long
Dim SegmentSize As Double
Dim Segments As Integer
Dim FName               As String
Dim FNameNoExt          As String
Dim stmp                As String
Dim FPath               As String
Dim filename            As String
WritePrivateProfileString "Power Splitter", "Compression level", _
CStr(Combo2.ListIndex), App.Path & "\Settings.ini"
WritePrivateProfileString "Power Splitter", "Type", _
CStr(Combo1.ListIndex), App.Path & "\Settings.ini"
WritePrivateProfileString "Power Splitter", "Custimize", _
CStr(cmbSize.ListIndex), App.Path & "\Settings.ini"
WritePrivateProfileString "Power Splitter", "Size", _
Text2.Text, App.Path & "\Settings.ini"
WritePrivateProfileString "Power Splitter", "archieve", _
CStr(frmMain.Text6.Text), App.Path & "\Settings.ini"
WritePrivateProfileString "Power Splitter", "folder", _
CStr(frmMain.txtFile(0).Text), App.Path & "\Settings.ini"
WritePrivateProfileString "Power Splitter", "batch", _
CStr(Check2.Value), App.Path & "\Settings.ini"
WritePrivateProfileString "Power Splitter", "delete1", _
CStr(chkDeleteParts.Value), App.Path & "\Settings.ini"
WritePrivateProfileString "Power Splitter", "delete2", _
CStr(Check4.Value), App.Path & "\Settings.ini"
WritePrivateProfileString "Power Splitter", "explorer", _
CStr(chkOpenExplorer.Value), App.Path & "\Settings.ini"
Unload Me
 Select Case modMain.Compress
 Case 0
 
    FName = GetFileName(desFile)                     '   Extract the file name
    FPath = GetFilePath(desFile)                     '   Extract the path name
    FNameNoExt = GetFileNoExtension(FName)           '   File name without extension
    filename = Tempop & FNameNoExt & ".tmp"
      X = SplitFile(filename, SegmentSize, Segments)
      If X <> 0 Then
      modMain.Compress = -1
      End If
 Case 1
    X = zipFile(desFile)
    If X = 0 Then
      MsgBox "The process completed successfully." & vbCrLf & _
      "The file was Compressed " & vbCrLf & frmMain.Text6.Text & vbCrLf _
     & "Compressed Size " & filelength(FileLen(frmMain.txtFile(0) & _
     frmMain.Text6)) & vbCrLf _
     & "Compressed Ratio " & Int((FileLen(frmMain.txtFile(0) & frmMain.Text6) _
     / totbytes) * 100) & " %", vbInformation, "PowerSplitter v." & App.Major _
     & "." & App.Minor
     Else
      If X <> Z_OK Then
        Select Case X
        Case Z_STREAM_END    '1
         err.Raise 801, , "Error compressing file " & " : Stream end"
        Case Z_NEED_DICT     '2
         err.Raise 802, , "Error compressing file " & " : Dictionary required"
        Case Z_ERRNO        '-1
          err.Raise 803, , "Error compressing file " & " : Unknown error"
        Case Z_STREAM_ERROR  '-2
          err.Raise 804, , "Error compressing file " & " : Stream error"
        Case Z_DATA_ERROR    '-3
          err.Raise 805, , "Error compressing file " & " : Input data corrupted"
        Case Z_MEM_ERROR     '-4
          err.Raise 806, , "Error compressing file " & " : Insufficient memory"
        Case Z_BUF_ERROR     '-5
         err.Raise 804, , "Error compressing file " & " : Insufficient space in output buffer"
        Case Z_VERSION_ERROR '-6
          err.Raise 805, , "Error compressing file " & " : zlib version error"
        Case Else '-6
          err.Raise 806, , "Error compressing file " & err.Description
        End Select
        End If
    End If
    Call frmMain.clear
    modMain.Compress = -1
 Case 2
    On Error GoTo ErrorHandler
    Dim Errorcode           As Integer
    Dim SourceBytes         As Long
    Dim SourceFile          As String
    Dim DestinationFile     As String
    Dim SegmentNumber       As Integer
    Dim bytesdone           As Long
    Dim RemainingBytes      As Long
    Dim para()              As String
    Dim bytes()             As Byte
    Dim i                   As Integer
    Dim hdr                 As header
       
     SegmentSize = segment_size()
    If SegmentSize = 0 Then                          'Ensure that the segment size is valid
        Errorcode = 2
        GoTo ErrorHandler
    End If
    FName = GetFileName(desFile)                     '   Extract the file name
    FPath = GetFilePath(desFile)                     '   Extract the path name
    FNameNoExt = GetFileNoExtension(FName)           '   File name without extension
    filename = Tempop & FNameNoExt & ".tmp"
    'Get total number or bytes in the source file
    SourceBytes = FileLen(filename)
    bytesdone = 0
    'Ensure that the resultant file segments will not exceed 999 segments
    'because otherwise we will have incorrect file extensions
    If SourceBytes / SegmentSize >= 1000 Then
        Errorcode = 3
        GoTo ErrorHandler
    End If
    ReDim para(1 To Int(SourceBytes / SegmentSize) + 1)
    ReDim bytes(1 To SegmentSize)
    'Open the source file for binary read
     frmMain.MousePointer = 13
    Open filename For Binary Access Read As #1 Len = 1

    SegmentNumber = 0
    Do
        'Increase the number of segments counter by 1
        SegmentNumber = SegmentNumber + 1

        'Compose the file name of the new file to be created (file segment)
        Select Case SegmentNumber
            Case Is < 10
          ' for batch file
            Call CopyMem(para(SegmentNumber), FNameNoExt & ".00" & CStr(SegmentNumber), _
            Len(FNameNoExt & ".00" & CStr(SegmentNumber)))
                DestinationFile = FPath & FNameNoExt & ".00" & CStr(SegmentNumber)
            Case 10 To 99
             Call CopyMem(para(SegmentNumber), FNameNoExt & ".0" & CStr(SegmentNumber), _
             Len(FNameNoExt & ".00" & CStr(SegmentNumber)))
                DestinationFile = FPath & FNameNoExt & ".0" & CStr(SegmentNumber)
            Case 100 To 999
             Call CopyMem(para(SegmentNumber), FNameNoExt & "." & CStr(SegmentNumber), _
             Len(FNameNoExt & ".00" & CStr(SegmentNumber)))
                DestinationFile = FPath & FNameNoExt & "." & CStr(SegmentNumber)
        End Select
        para(SegmentNumber) = Trim(para(SegmentNumber))
        Open DestinationFile For Binary Access Write As #2 Len = 1
        'Check whether the remaining bytes to process in the source file are
        'less than Segment bytes
        Select Case SourceBytes - bytesdone
        Case Is < SegmentSize
            RemainingBytes = SourceBytes - bytesdone
            ReDim bytes(1 To RemainingBytes)
        Case Else
            RemainingBytes = SegmentSize
        End Select
        frmMain.lblstatus = "Writing...."
        frmMain.lblCurrentFile = GetFileName(DestinationFile)
        'Read bytes from the source file and write them to
        'the destination file (the current segment file)
        'Depending on the remaining bytes to read and write,
        'the routine below will read the largest possible
        'burst mode copy ............
                    Get #1, , bytes        'copy in ms
                    Put #2, , bytes
                    '   Update the bytes done counter
                    bytesdone = bytesdone + RemainingBytes
                    Close #2
            'Update the percent control on the form
            DrawPercent Int((bytesdone / SourceBytes) * 100)
             frmMain.lblFileCount = "Finished (" & SegmentNumber & " Files)"
             time2 = Timer
             frmMain.lblElapsed = Format(time2 - time1, "#.00") & " seconds"
            'Refresh the form and yield to windows
    Loop Until bytesdone = SourceBytes
    'Close the source file
    Close 1
    frmMain.lblstatus = "Writing..."
    'header and file info in .000 file
    stmp = FNameNoExt & ".000"
    frmMain.lblCurrentFile = stmp
    stmp = FPath & FNameNoExt & ".000"
    hdr.name = "[Power Splitter]"
    hdr.totfiles = start
    hdr.zip = level  ' check for zip split
    hdr.splitfiles = SegmentNumber
    Open stmp For Binary Access Write As #1
    Put #1, , hdr
    For i = 1 To start
    Put #1, , filearr(i)
    Next i
    Close #1
   ' Make a batch file
    If Check2.Value = 1 And start = 1 Then
    FName = GetFileName(filearr(1).name)
    Dim str1 As String
    Dim Filenr As Integer
    str1 = Join(para, "+")
    If Right(FPath, 1) <> "\" Then
    stmp = FPath & "\" & FNameNoExt & "000" & ".bat"
    Else
    stmp = FPath & FNameNoExt & "000" & ".bat"
    End If
    Filenr = FreeFile
    'Save the destination string
     Open stmp For Output As #Filenr
     Print #Filenr, "@Echo off"
     Print #Filenr, , "Echo    *****************************"
     Print #Filenr, , "Echo      Power Spliter batch file "
     Print #Filenr, , "Echo        by Dheeraj Khajuria"
     Print #Filenr, , "Echo    *****************************"
     Print #Filenr, "Echo joining files........in Progess"
     Print #Filenr, "Copy /b " & str1 & " " & FName
     If frmopt.Check4.Value = 1 Then 'delete the splitted file's
     Print #Filenr, "Ren " & FName & " temp$.$$$" ' rename
     Print #Filenr, "Del " & FNameNoExt & ".*"
     Print #Filenr, "Ren   temp$.$$$  " & FName     ' rename again
     Print #Filenr, "Echo " & FName & " has been recreated for you... "
    ''bug
     Print #Filenr, "goto end"
     End If
     Print #Filenr, "Echo " & FName & " has been recreated for you... "
     Print #Filenr, ": End"
     Close #Filenr
    End If
    'When the code reaches this point, everything went OK.
    'Acknowledge the number of segments, assign the value '0' to the function and exit
    If frmopt.chkDeleteParts = 1 Then
    If MsgBox("Delete the Orignal file(s)", vbQuestion + vbYesNo, _
    "Conform Delete") = vbYes Then
    For i = 1 To start
    Kill filearr(i).name
    Next i
    End If
    End If
    DrawPercent 100
    frmMain.lblstatus = "Idle..."
    frmMain.MousePointer = 0
ErrorHandler:
 If modMain.Compress <> -1 Then
   'Inform the user about the call success or failure
    Select Case Errorcode
    Case 0
        sMsg = "The process completed successfully." & vbCr & _
        "The file was split to " & SegmentNumber & " segments."
        sMsgDis = 64
    Case 1
        sMsg = "File does not exists."
        sMsgDis = 16
    Case 2
        sMsg = "Invalid segment size."
        sMsgDis = 16
    Case 3
        sMsg = "Unable to create more than 999 segments." & vbCr & _
        "Please raise segment size and try again."
        sMsgDis = 16
    Case 5
        sMsg = "Resulting file is smaller than requested segment size." & _
        vbCr & "Please try another size or don't split the file at all :)"
        sMsgDis = 64
    Case 4
        sMsg = "Unknown Error!!"
        sMsgDis = 16
    End Select
    MsgBox sMsg, sMsgDis, "PowerSplitter v." & App.Major & "." & App.Minor
    Call frmMain.clear
    frmMain.Toolbar1.Buttons(3).Enabled = False
    frmMain.add_mnu.Enabled = False
    modMain.Compress = -1
End If
End Select
End Sub
Private Sub Form_Load()
 On Error Resume Next
     With Combo2
  .AddItem "Medium compression(Normal)" '= -1
  .AddItem "No compression"             '= 0
  .AddItem "Low compression(Fastest)"   '= 1
  .AddItem "Light compression(Fast)"    '= 3
  .AddItem "Light compression(Medium)"  '= 4
  .AddItem "High compression(Slow)"     '= 6
  .AddItem "Highest compression(Slowest)" '= 9
  .ListIndex = read000("Power Splitter", App.Path & "\settings.ini", _
  "Compression level", 0)
     End With
    With Combo1
         .AddItem "Bytes"
         .AddItem "Kb"
         .AddItem "Mb"
         .AddItem "Segments"
   .ListIndex = read000("Power Splitter", App.Path & "\settings.ini", "Type", 2)
     End With
    With cmbSize
        .AddItem "Custimize"
        .AddItem "100 Kb"
        .AddItem "250 Kb"
        .AddItem "500 Kb"
        .AddItem "720 Kb"                 ' 713 kb
        .AddItem "1.00 Mb"
        .AddItem "1.44 Mb"                ' 1.38 kb
        .AddItem "2.88 Mb"
        .AddItem "5.00 Mb"
        .AddItem "7.50 Mb"
        .AddItem "10.0 Mb"
        .AddItem "25.0 Mb"
        .ListIndex = read000("Power Splitter", _
        App.Path & "\settings.ini", "Custimize", 0)
    End With
Check2.Value = read000("Power Splitter", App.Path & "\settings.ini", "batch", 0)
Check4.Value = read000("Power Splitter", App.Path & "\settings.ini", "delete2", 0)
chkDeleteParts.Value = read000("Power Splitter", App.Path & "\settings.ini", _
"delete1", 0)
chkOpenExplorer.Value = read000("Power Splitter", App.Path & "\settings.ini", _
"explorer", 0)
'set default (1.44Mb)
 Text2.Text = read000("Power Splitter", App.Path & "\settings.ini", "Size", "1.44")
 
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
' its works
 Select Case KeyAscii
  Case 48 And (Text2.Text) = "0"     'for zeros
           KeyAscii = 0
  Case 8   ' backSpace
  Case 46 And InStr(1, Text2.Text, ".", vbTextCompare) = 0 _
  And Combo1.Text <> "Segments" And Len(Text2.Text) <> 0 ' . dot
 Case Else
If (KeyAscii) < 48 Or (KeyAscii) > 57 Then
MsgBox "You should Enter valid Integer", vbExclamation, "Integer"
KeyAscii = 0
End If
End Select
End Sub
Private Sub Text2_Change()
If Combo1.Text = "Segments" And Len(Text2.Text) >= 4 Then
Text2.Text = ""
End If
Call frmMain.info
End Sub


Private Sub Combo2_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then Call SendMessage(Combo2.hWnd, CB_SHOWDROPDOWN, True, 0&)
End Sub


Private Sub Combo2_click()
'On Error Resume Next
Frame6.Enabled = False
Select Case Combo2.ListIndex
Case 0
     level = -1
Case 1
     level = 0  ' no compression
     Frame6.Enabled = True
Case 2
     level = 1
Case 3
     level = 3
Case 4
     level = 4
Case 5
     level = 6
Case 6
     level = 9
 End Select
End Sub




Private Sub cmbSize_Click()
Call frmMain.info
If (cmbSize.Text) = "Custimize" Then
 Combo1.Visible = True
 Text2.Visible = True
 Label2.Visible = True
 Else
 Combo1.Visible = False
 Text2.Visible = False
 Label2.Visible = False
End If
End Sub


Private Sub cmbSize_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call SendMessage(cmbSize.hWnd, CB_SHOWDROPDOWN, True, 0&)
End Sub
Private Sub Combo1_Click()
If Combo1.Text = "Segments" Then
If Len(Text2.Text) >= 4 Then
Text2.Text = 999
End If
End If
Call frmMain.info
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then Call SendMessage(Combo1.hWnd, CB_SHOWDROPDOWN, True, 0&)
End Sub

