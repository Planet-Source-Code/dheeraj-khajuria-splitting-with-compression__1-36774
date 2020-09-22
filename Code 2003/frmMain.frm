VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmMain 
   Caption         =   "Power Splitter"
   ClientHeight    =   6315
   ClientLeft      =   3615
   ClientTop       =   2760
   ClientWidth     =   6930
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   6930
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   7
      Top             =   4920
      Width           =   3615
      Begin VB.TextBox txtFile 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   200
         Width           =   2175
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
         ButtonWidth     =   1588
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgToolbar"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Folder"
               Object.ToolTipText     =   "Browse for  Output path "
               ImageKey        =   "Folder"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label4 
         Caption         =   "Archive Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   255
         Width           =   1245
      End
   End
   Begin MSComctlLib.ListView listfiles 
      Height          =   4125
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   7276
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   -9
         Key             =   "Name"
         Text            =   "Name"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Type"
         Text            =   "Type"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   "Size"
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Modified"
         Text            =   "Modified"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Key             =   "Ratio"
         Text            =   "Ratio"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Key             =   "Packed"
         Text            =   "Packed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "CRC"
         Text            =   "CRC"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "Path"
         Text            =   "Path"
         Object.Width           =   6068
      EndProperty
   End
   Begin VB.Frame Frame7 
      Caption         =   "Working Statistics"
      Height          =   1095
      Left            =   3720
      TabIndex        =   10
      Top             =   4920
      Width           =   3135
      Begin VB.Label Label13 
         Caption         =   "Elapsed Time:"
         Height          =   180
         Index           =   4
         Left            =   1560
         TabIndex        =   17
         Top             =   120
         Width           =   1365
      End
      Begin VB.Label Label13 
         Caption         =   "File Count:"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1050
      End
      Begin VB.Label Label13 
         Caption         =   "Current file:"
         Height          =   180
         Index           =   1
         Left            =   1560
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblCurrentFile 
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1560
         TabIndex        =   14
         Top             =   840
         Width           =   1530
      End
      Begin VB.Label lblstatus 
         AutoSize        =   -1  'True
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   90
      End
      Begin VB.Label lblFileCount 
         AutoSize        =   -1  'True
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   90
      End
      Begin VB.Label lblElapsed 
         AutoSize        =   -1  'True
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1560
         TabIndex        =   11
         Top             =   360
         Width           =   90
      End
   End
   Begin VB.PictureBox pic1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4920
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   9
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3840
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1440
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar P1 
      Height          =   180
      Left            =   1440
      TabIndex        =   6
      Top             =   1920
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList imgToolbar 
      Left            =   2640
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":171E
            Key             =   "split"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1FFA
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E4E
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":32A2
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":383E
            Key             =   "Export"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3DDA
            Key             =   "New"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4C2E
            Key             =   "zip"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":51CA
            Key             =   "order"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5766
            Key             =   "add"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":65BA
            Key             =   "option"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6B56
            Key             =   "info"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2160
      Top             =   2160
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6060
      Width           =   6930
      _ExtentX        =   12224
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1940
            MinWidth        =   1940
            Text            =   "Power Splitter"
            TextSave        =   "Power Splitter"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4817
            MinWidth        =   4817
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2011
            MinWidth        =   2011
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   9698
            MinWidth        =   9698
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   1085
      BandCount       =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   11640
      _CBHeight       =   615
      _Version        =   "6.0.8169"
      Child1          =   "Toolbar1"
      MinHeight1      =   555
      Width1          =   12030
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   555
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   11520
         _ExtentX        =   20320
         _ExtentY        =   979
         ButtonWidth     =   1164
         ButtonHeight    =   979
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "imgToolbar"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "New"
               Object.ToolTipText     =   "Browse for  Output path "
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Open"
               Object.ToolTipText     =   "open a Archive"
               ImageKey        =   "Open"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Add"
               Object.ToolTipText     =   "Add the file(s)"
               ImageKey        =   "add"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Delete"
               Object.ToolTipText     =   "Delete a file(s)"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "SPlit"
               Object.ToolTipText     =   "Split or  Compress the file(s)"
               ImageKey        =   "split"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Zip"
               Object.ToolTipText     =   "Compress the files"
               ImageKey        =   "zip"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Option"
               Object.ToolTipText     =   "Options"
               ImageKey        =   "option"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "About"
               Object.ToolTipText     =   "About the author"
               ImageKey        =   "info"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Exit"
               Object.ToolTipText     =   "Press to Exit"
               ImageKey        =   "exit"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu file_mnu 
      Caption         =   "&File"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu New_mnu 
         Caption         =   "&New              "
         Shortcut        =   ^N
      End
      Begin VB.Menu Open_mnu 
         Caption         =   "&Open...    "
         Shortcut        =   ^O
      End
      Begin VB.Menu mnu0 
         Caption         =   "-"
      End
      Begin VB.Menu add_mnu 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Shortcut        =   ^A
      End
      Begin VB.Menu mnu_del 
         Caption         =   "Delete"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_exit 
         Caption         =   "&Exit"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnu_opt 
      Caption         =   "Op&tion"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu mnu_Split 
         Caption         =   "S&plit        "
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu_zip 
         Caption         =   "Zip"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnu3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_option 
         Caption         =   "Options"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnu_a 
      Caption         =   "&About"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu mnu_about 
         Caption         =   "A&bout"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'///////////////////////////////////////////////
' FrmMain.frm
' Splits and joins files, with optional compression
' Original by Dheeraj Khajuria Copyright 2002
'
'//////////////////////////////////////////////

Option Explicit
Dim SH              As New Shell  'reference to shell32.dll class
Dim ShBFF           As Folder     'Shell Browse For Folder
Dim cap             As String
Dim cap1            As String
Dim selected        As Integer
Dim desin           As String
Dim i               As Integer
Public LF           As ListItem
Public LFS          As ListSubItem

Private Type SHFILEINFO            'As required by ShInfo
  hIcon             As Long
  iIcon             As Long
  dwAttributes      As Long
  szDisplayName     As String * 255
  szTypeName        As String * 80
End Type

'Functions to extract icons & place them in a picture box
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
 (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, _
  ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
    
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, _
ByVal hDCDest&, ByVal X&, ByVal Y&, ByVal flags&) As Long
'Icon Sizes in pixels
Private Const SMALL_ICON As Integer = 16    'Icon size
'ShellInfo Flags
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000 'System icon index
Private Const SHGFI_LARGEICON = &H0       'Large icon
Private Const SHGFI_SMALLICON = &H1       'Small icon
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or _
SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
Private ShInfo     As SHFILEINFO
Private mIml       As ImageList                 'Imagelist containing small icons
Private mpic       As PictureBox                'Temporay container for small icon
Private mPicDef    As PictureBox

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal _
hWnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As _
String, ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Public Sub info()
If start <> 0 Then
StatusBar1.Panels.Item(2).Text = Format(start, "#,##0") & " file(s) " & FormatNumber( _
totbytes, 0, , , vbTrue) & " Byte (" & filelength(totbytes) & ")"  'size of file
Dim Size As Double
Size = segment_size()
If Size <> 0 Then
If totbytes / Size >= 1000 Then
frmopt.Text4 = "Can't Have More Then (999) Segments"
frmopt.Text3 = ""
Else
If (totbytes / Size) <> 0 Then
If Size >= 1024 Then
frmopt.Text3 = Int(Size / 1024) & " Kb"
Else
frmopt.Text3 = Int(Size) & "  bytes"
End If
frmopt.Text4 = Round(totbytes / Size) & " Segment"
Else
frmopt.Text4 = "1 Segment"
End If
End If
End If
Else
StatusBar1.Panels.Item(2).Text = ""
End If
End Sub

Private Function Find(filename As String) As files
  On Error Resume Next
  Set f = fs.GetFile(filename)
  Find.name = filename
  Find.Size = FileLen(filename)
  Find.attr = f.Attributes
  Find.type = f.type
  Find.Mdate = FormatDateTime(f.DateLastModified, vbGeneralDate)
  Find.icon = GetIconKey(filename)
  End Function

Private Sub cmdBrowse_Click(filename As String)
    Dim numberofsfiles As Integer
    Dim Compressed As Integer
    Dim sDestinationPath As String
    Dim sfilename As String
    Dim stmp As String
    Dim hdr As header
    Dim i As Integer
    Dim pinfo As Long
    Dim getinfo As Long
On Error Resume Next
'set object
   If Len(filename) = 0 Then
        With CommonDialog1
            .DialogTitle = "Select the Split or Zip .000 File"
            .Filter = "Info File (*.000)|*.000"
            .filename = vbNullString
            .ShowOpen
            If LenB(.filename) <> 0 Then
            Text6.Text = GetFileName(.filename)
            txtFile(0) = GetFilePath(.filename)
            desin = GetFilePath(.filename)
            Else
            Exit Sub
            End If
        End With
    Else
    Text6.Text = GetFileName(filename)
    txtFile(0) = GetFilePath(filename)
    desin = GetFilePath(filename)
   End If
lblstatus = "Retrieving Info..."

  If txtFile(0).Text <> "" Then
    sDestinationPath = desin
    sfilename = Text6.Text                             ' filename
    If Right(sDestinationPath, 1) <> "\" Then
     sDestinationPath = sDestinationPath & "\"
    End If
    stmp = sDestinationPath & sfilename
    Select Case LenB(sfilename)
    Case 0                                             ' Field is empty
        MsgBox "Invalid file Selected.", 16, "PowerSplitter v." & App.Major _
        & "." & App.Minor
        Exit Sub
    Case Else
        If Right$(sfilename, 3) = "000" Then          ' File is not a valid info file
        frmMain.lblCurrentFile = GetFileName(stmp)
        Open stmp For Binary Access Read As #1
        Get #1, , hdr
        getinfo = Loc(1)
        start = hdr.totfiles
        Compressed = hdr.zip
        numberofsfiles = hdr.splitfiles
        time1 = Timer
        Erase filearr
        ReDim filearr(start)
    
        Select Case hdr.name
        Case "[Power Splitter]"                       ' Split files
        totbytes = 0
        For i = 1 To start
        Get #1, , filearr(i)
        totbytes = totbytes + filearr(i).Size
        Next i
        Close #1
        Call cmdJoin_Click(sfilename, _
        numberofsfiles, start, Compressed, sDestinationPath)
        
        Case "[Power Compressed]"                    ' compressed files
        Seek #1, LOF(1) - 3
        Get #1, , pinfo
        Seek #1, pinfo + 1
        totbytes = 0
        For i = 1 To start
        Get #1, , filearr(i)
        totbytes = totbytes + filearr(i).Size
        Next i
        Close #1
        DecompressFile stmp, sDestinationPath, start, getinfo, Compressed
        Call showdetail(1, sDestinationPath)
        ' Every thing is ok
         MsgBox "file's Decompressed!", vbInformation, "PowerSplitter v." & _
         App.Major & "." & App.Minor
         Call clear
         Call info
        Case Else
           MsgBox GetFileName(sfilename) & " is not a valid File, please select " & _
           "using the dialog!", 16, "PowerSplitter v." & App.Major & "." & App.Minor
        End Select
     End If
    End Select
  End If
End Sub
Private Sub showdetail(check As Integer, Path As String)
 On Error Resume Next
 Dim status As Boolean
 
 Dim i As Integer
    ' info  is ok..
       listfiles.ListItems.clear
        For i = 1 To start
       ' set the attributes
        Call SetAttr(Path & GetFileName(filearr(i).name), filearr(i).attr)
        filearr(i).icon = GetIconKey(Path & GetFileName(filearr(i).name))
        status = False
        If check <> 0 And filearr(i).Size <> 0 Then status = True
        showInfo filearr(i), status
        Next i
End Sub

Private Sub add_mnu_Click()
frmMain.MousePointer = 11
        Call add_Click
frmMain.MousePointer = 0
End Sub
Private Sub cmdClose_Click()
 On Error Resume Next
Dim f As Form
Dim modname, Key As String
Dim wdth As Integer
Kill Tempop & "*.bmp"
Kill Tempop & "*.ino"
Kill Tempop & "*.tmp"
modname = App.Path & "\Settings.ini"
If Me.WindowState = vbNormal Then
WritePrivateProfileString "Power Splitter", "left", CStr(frmMain.Left), modname
WritePrivateProfileString "Power Splitter", "Top", CStr(frmMain.Top), modname
WritePrivateProfileString "Power Splitter", "Width", CStr(frmMain.Width), modname
WritePrivateProfileString "Power Splitter", "Height", CStr(frmMain.Height), modname
End If
For i = 1 To listfiles.ColumnHeaders.count
  Key = "ColumnHeader " & CStr(i)
  wdth = listfiles.ColumnHeaders(i).Width
  WritePrivateProfileString "Power Splitter", Key, CStr(wdth), modname
Next
For Each f In Forms
  If f.name <> "frmMain" Then
    Unload f
  End If
Next
Unload frmMain
End Sub

Private Sub cmdJoin_Click(sfilename As String, numberofsfiles As Integer, _
Numberoffiles As Integer, Compressed As Integer, sDestinationPath As String)
    Call frmMain.clear
    With frmMain
       .lblCurrentFile = GetFileNoExtension(sfilename) & ".001"
       .lblstatus = "ReadingFile..."
    End With
    i = JoinFile(GetFileNoExtension(sfilename) & ".000", Val(numberofsfiles), _
    Val(Compressed), Val(Numberoffiles), sDestinationPath)
    If frmopt.chkOpenExplorer.Value = vbChecked Then SH.Open (desin)
    Call showdetail(Compressed, sDestinationPath)
    If i = 0 Then
       MsgBox " Ok! File's are Joined:" & GetFileName(sfilename), _
       vbInformation, cap1 & " Join!"
    Else
      MsgBox "Error in Joining!:" & GetFileName(sfilename), vbInformation, cap1 & " Error!"
    End If
End Sub

Private Sub split_Click()
     If start <> 0 Then
     time1 = Timer
    Dim X As Integer
    cap1 = "PowerSplitter v." & App.Major & "." & App.Minor & " " & Text6.Text
    If Right(txtFile(0).Text, 1) <> "\" Then txtFile(0).Text = txtFile(0).Text & "\"
    If GetFileExtension(Text6.Text) <> ".000" Then Text6.Text = Text6.Text & ".000"
    desFile = txtFile(0).Text & Text6.Text
    Call h_Click  ' make dir's
    lblstatus = "Retrieving Info..."
    'Call the option
    modMain.Compress = 0
    frmopt.Frame5.Visible = False
    frmopt.Height = 1300
    frmopt.Show , Me
    Else
    MsgBox "Please Add the files to Split", vbInformation, cap1
    End If
End Sub

Private Sub add_Click()
   'Initialize the common dialog control and show it
   On Error GoTo MS: ' Error control
    Dim vFiles       As Variant
    Dim lFile        As Integer
    Dim res          As Integer
    Dim upper        As Integer
    Dim Arcname      As String
    Dim str          As String
    With CommonDialog1
         .CancelError = True
        .filename = vbNullString
        .CancelError = True 'Gives an error if cancel is pressed
        .DialogTitle = "Select file(s) to split"
        .flags = cdlOFNAllowMultiselect Or cdlOFNFileMustExist Or cdlOFNExplorer _
        Or cdlOFNNoDereferenceLinks
        'Flags, allows Multi select, Explorer style and hide the Read only tag
        .Filter = "All Files (*.*)|*.*"
        .MaxFileSize = 30000 'number of files buffer
        .ShowOpen
     End With
    frmMain.MousePointer = 13
    If Right(txtFile(0).Text, 1) <> "\" Then txtFile(0).Text = txtFile(0).Text & "\"
    Arcname = txtFile(0).Text & Text6
    vFiles = Split(CommonDialog1.filename, vbNullChar) 'Splits the filename up in segments
    ReDim Preserve filearr(1 To UBound(vFiles) + start + 1)
    'File is part/same of the current archive (.000)
    i = 1
    If UBound(vFiles) = 0 Then
    upper = 1     ' If there is only 1 file then do this
    str = CStr(vFiles(0))
    Else
    If Right(vFiles(0), 1) <> "\" Then
    vFiles(0) = vFiles(0) & "\"
    End If
    ' More than 1 file then do this until there are no more files
    upper = UBound(vFiles)
    End If
    For lFile = start + 1 To (upper + start)
    If upper <> 1 Then str = vFiles(0) & GetFileName(CStr(vFiles(i)))
    filearr(lFile) = Find(str)               'call find
    res = check(filearr(lFile), lFile - 1)
    If GetFileExtension(filearr(lFile).name) = ".000" Then
    If LCase(filearr(lFile).name) = LCase(Arcname) Then
    MsgBox "Can't Include the Archieve name : " & GetFileName(Arcname), vbCritical, cap1
    Call new_Click
    Exit Sub
    End If
    End If
    If res = -1 Then
       totbytes = totbytes + filearr(i).Size
      start = start + 1
      Else
      ' update the files if res <> -1
      totbytes = totbytes - filearr(res).Size
      totbytes = totbytes + filearr(lFile).Size
      filearr(res) = filearr(lFile) ' updated
         End If
     i = i + 1
    Next
listfiles.ListItems.clear
For i = 1 To start
showInfo filearr(i), False
Next i
Select Case start
Case 1
frmopt.Check2.Enabled = True
Case Else
frmopt.Check2.Enabled = False
End Select
Text6.Text = GetFileNoExtension(GetFileName(filearr(1).name)) & ".000"
Call info
frmMain.MousePointer = 0
Exit Sub
MS:
End Sub

Private Sub delete_Click()
Dim i, j As Integer
Dim Item As ListItem
On Error GoTo CmdDelErr
Screen.MousePointer = vbArrowHourglass
For i = listfiles.ListItems.count To 1 Step -1
Set Item = listfiles.ListItems(i)
If Item.selected Then
listfiles.ListItems.Remove i
If i <> start Then
For j = i To start - 1
filearr(j) = filearr(j + 1)
Next j
End If
start = start - 1
End If
Next i
totbytes = 0
For i = 1 To start
totbytes = totbytes + filearr(i).Size
Next i
listfiles.Refresh
Call info
Screen.MousePointer = vbNormal
Exit Sub
CmdDelErr:
  Resume Next
End Sub



Private Sub Form_Activate()
Dim commandString As String
 On Error GoTo e_next
    commandString = Command
    If commandString <> "" Then
        If Dir(commandString) <> "" Then
       If StrComp(GetFileExtension(commandString), ".000", vbTextCompare) = 0 Then
       Call cmdBrowse_Click(commandString)
       End If
        End If
         Else
    End If
e_next:
End Sub

Private Sub Form_Load()
Dim l, t, w, h
Dim i, Key As String, wdth
Dim modname As String
totbytes = 0
cap1 = "PowerSplitter v." & App.Major & "." & App.Minor
Call setreg
modname = App.Path & "\Settings.ini"
If Me.WindowState = vbNormal Then
  l = read000("Power Splitter", modname, "Left", (Screen.Width - Me.Width) / 2)
  t = read000("Power Splitter", modname, "Top", (Screen.Height - Me.Height) / 2)
  w = read000("Power Splitter", modname, "Width", Me.Width)
  h = read000("Power Splitter", modname, "Height", Me.Height)
  Me.Move l, t, w, h
End If
For i = 1 To listfiles.ColumnHeaders.count
  Key = "ColumnHeader " & CStr(i)
  wdth = listfiles.ColumnHeaders(i).Width
  listfiles.ColumnHeaders(i).Width = read000("Power Splitter", modname, Key, Int(wdth))
Next
   Set fs = CreateObject("Scripting.FileSystemObject")
   start = 0
   Dim TempPath As String * 255
   GetTempPath 254, TempPath
   Tempop = Mid(TempPath, 1, InStr(1, TempPath, Chr(0), vbTextCompare) - 1)
   txtFile(0) = read000("Power Splitter", modname, "folder", "")
   Text6 = read000("Power Splitter", modname, "archieve", "")
   Toolbar1.Buttons(3).Enabled = False
   Set mIml = ImageList1
   Set mpic = Picture1
   'Initialise picture box
   mpic.Width = (SMALL_ICON) * Screen.TwipsPerPixelX
   mpic.Height = (SMALL_ICON) * Screen.TwipsPerPixelY
   mpic.AutoRedraw = True
   'Initialise Default Picture
    Set mPicDef = pic1
    modMain.Compress = -1
End Sub



Private Sub Form_Resize()
On Error Resume Next
If frmMain.Width >= 7950 Then
If frmMain.Height >= 3240 Then
CoolBar1.Width = frmMain.Width - 130
Frame7.Width = frmMain.Width - 240 - Frame1.Width
listfiles.Width = frmMain.Width - 130
listfiles.Height = frmMain.Height - Frame7.Height - CoolBar1.Height - _
StatusBar1.Height - 880
Frame7.Top = listfiles.Height + listfiles.Top - 10
Frame1.Top = listfiles.Height + listfiles.Top - 10
StatusBar1.Top = listfiles.Height + Frame1.Height + CoolBar1.Height
StatusBar1.Panels(4).Width = frmMain.Width - StatusBar1.Panels(1).Width - _
StatusBar1.Panels(2).Width - StatusBar1.Panels(3).Width
P1.Top = StatusBar1.Top + 50
P1.Left = StatusBar1.Panels(4).Left + 20
P1.Width = StatusBar1.Panels(4).Width - 350
lblCurrentFile.Width = Frame7.Width / 2
Else
frmMain.Height = 3240
End If
Else
frmMain.Width = 7950
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call cmdClose_Click
End Sub

Private Sub ListFiles_DblClick()
If start <> 0 Then
On Error Resume Next
SH.MinimizeAll
 Dim Scr_hDC As Long
    Scr_hDC = GetDesktopWindow()
   ShellExecute Scr_hDC, "Open", _
   filearr(listfiles.SelectedItem.Index).name, "", "", 1
   End If
End Sub
Private Sub ListFiles_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode And Shift = 0
 Case 46
 Call delete_Click
 Case 13
 Call ListFiles_DblClick
 End Select
End Sub

Private Sub mnu_about_Click()
frmMain.MousePointer = 11
frmabout.Show
frmMain.MousePointer = 0
End Sub

Private Sub mnu_del_Click()
frmMain.MousePointer = 11
Call delete_Click
frmMain.MousePointer = 0
End Sub

Private Sub mnu_exit_Click()
Call cmdClose_Click
End Sub

Private Sub mnu_option_Click()
frmMain.MousePointer = 11
modMain.Compress = -1
frmopt.Show
frmMain.MousePointer = 0
End Sub

Private Sub mnu_Split_Click()
frmMain.MousePointer = 11
Call split_Click
frmMain.MousePointer = 0
End Sub

Private Sub mnu_zip_Click()
frmMain.MousePointer = 11
Call zip_Click
frmMain.MousePointer = 0
End Sub

Private Sub New_mnu_Click()
frmMain.MousePointer = 11
Call new_Click
frmMain.MousePointer = 0
End Sub

Private Sub Open_mnu_Click()
frmMain.MousePointer = 11
Call cmdBrowse_Click("")
frmMain.MousePointer = 0
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
  Case 8   ' backSpace
  Case 46
  KeyAscii = 0
  Case 58
  KeyAscii = 0
  Case 92
  KeyAscii = 0
  Case 32   ' space
  KeyAscii = 0
End Select
End Sub

Private Sub Timer1_Timer()
If Len(cap1) >= 30 Then
If i = Len(cap1) Then
i = 0
Else
i = i + 1
cap = Mid(cap1, i + 1, Len(cap1))
Me.Caption = cap
End If
Else
If StrComp(cap, cap1, vbTextCompare) = -1 Then
Me.Caption = cap1
CopyMem cap, cap1, Len(cap1)
End If
End If
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
 frmMain.MousePointer = 11
Select Case Button.Caption
Case "New"
        Call new_Click
Case "Open"
         Call cmdBrowse_Click("")
Case "Add"
        Call add_Click
Case "Delete"
       Call delete_Click
Case "SPlit"
       Call split_Click
Case "Zip"
        Call zip_Click
Case "Option"
         modMain.Compress = -1
         frmopt.Show
Case "About"
         frmabout.Show
Case "Exit"
       Call cmdClose_Click
         Exit Sub
 End Select
 frmMain.MousePointer = 0
End Sub

Private Function check(filechk As files, chk As Integer) As Integer
Dim k As Integer
'check for file exist already
For k = 1 To chk
With filechk
If .name = filearr(k).name Then
 check = k
 Exit Function
End If
End With
Next k
check = -1
End Function
Public Sub clear()
    frmMain.P1.Value = 0
    frmMain.lblstatus = "#"
    frmMain.lblFileCount = "#"
    frmMain.lblCurrentFile = "#"
    frmMain.lblElapsed = "#"
    StatusBar1.Panels.Item(2).Text = ""
    StatusBar1.Panels.Item(3).Text = ""
End Sub


Private Sub new_Click()
'Clear the list view
Toolbar1.Buttons(3).Enabled = True
frmMain.add_mnu.Enabled = True
listfiles.ListItems.clear
Erase filearr
start = 0
totbytes = 0
Call info
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Set ShBFF = SH.BrowseForFolder(Me.hWnd, "Destination Folder!" & _
            "Please choose a Path and click OK!", 1)
            desin = ShBFF.Items.Item.Path
            txtFile(0).Text = desin
End Sub

Private Sub h_Click()
On Error Resume Next
Dim make As Integer
make = 0
Do
  make = InStr(4 + make, txtFile(0), "\", vbTextCompare)
   If make <> 0 Then
   MkDir (Left(txtFile(0).Text, make))
   End If
Loop While make <> 0
End Sub

Private Sub txtFile_KeyPress(Index As Integer, KeyAscii As Integer)
If Right(txtFile(0).Text, 1) = "\" And KeyAscii = 92 Then
KeyAscii = 0
Exit Sub
End If
If InStr(1, txtFile(0).Text, ":") <> 0 And KeyAscii = 58 Then
KeyAscii = 0
Exit Sub
End If
 Select Case KeyAscii
  Case 47
  KeyAscii = 0
  Case 42
  KeyAscii = 0
  Case 63
  KeyAscii = 0
  Case 34
  KeyAscii = 0
  Case 62
  KeyAscii = 0
  Case 60
  KeyAscii = 0
  Case 124
  KeyAscii = 0
End Select
End Sub


Private Sub zip_Click()
    If start <> 0 Then
    cap1 = "PowerSplitter v." & App.Major & "." & App.Minor & " " & Text6.Text
    If Right(txtFile(0).Text, 1) <> "\" Then txtFile(0).Text = txtFile(0).Text & "\"
    If GetFileExtension(Text6.Text) <> ".000" Then Text6.Text = Text6.Text & ".000"
    desFile = txtFile(0).Text & Text6.Text
    Call h_Click  ' make dir's
    lblstatus = "Retrieving Info..."
    'Call the option
    modMain.Compress = 1
    frmopt.Frame5.Visible = False
    frmopt.Height = 1300
    frmopt.Show , Me
    Else
    MsgBox "Please Add the files to Zip", vbInformation, cap1
    End If
End Sub

Private Function GetIconKey(filename As String) As String
Dim hSmallIcon As Long        'Handle to small Icon
Dim imgObj As ListImage
'Single image object in imagelist.listimages collection
Dim Key As String
Dim Ret As Long               'Return value
Dim Ext As String
On Error Resume Next
'Get a handle to the small icon
Ext = Right(GetFileExtension(filename), 3)
If InStr(1, "exe|pif|lnk|ico|cur", Ext, vbTextCompare) > 0 Then
  Key = LCase(filename) & "icons"
Else
  Key = Ext & "icons"
End If
hSmallIcon = SHGetFileInfo(filename, 0&, ShInfo, Len(ShInfo), _
BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
'If the handle exists, load it into the picture box(es)
If hSmallIcon <> 0 Then
  'Small Icon
  With mpic
    Set .Picture = LoadPicture("")
    Ret = ImageList_Draw(hSmallIcon, ShInfo.iIcon, mpic.hDC, 0, 0, 0)
    .Refresh
  End With
  'Try to access the image in the image list, if it doesn't exist then
  'an error is returned
  Set imgObj = mIml.ListImages(Key)
  If err <> 0 Then
    'Add the image if it doesn't exist.
    err = 0
    Set imgObj = mIml.ListImages.Add(, Key, mpic.Image)
    'Save to temp file
    SavePicture mpic.Image, Tempop & Key & ".bmp"
    nIcons = nIcons + 1
  End If
Else
  'Try to access the image in the image list, if it doesn't exist then
  'an error is returned
  Set imgObj = mIml.ListImages(Key)
  If err <> 0 Then
    'Add the image if it doesn't exist.
    err = 0
    Set imgObj = mIml.ListImages.Add(, Key, mPicDef.Image)
    'Save to temp file
    SavePicture mPicDef.Image, Tempop & "\" & Key & ".bmp"
    nIcons = nIcons + 1
  End If
End If
GetIconKey = Key
End Function

