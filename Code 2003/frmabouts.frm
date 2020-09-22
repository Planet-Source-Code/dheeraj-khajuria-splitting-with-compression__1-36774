VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Power Splitter"
   ClientHeight    =   3090
   ClientLeft      =   4725
   ClientTop       =   4320
   ClientWidth     =   5280
   Icon            =   "frmabouts.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2355
      ScaleWidth      =   4995
      TabIndex        =   1
      Top             =   120
      Width           =   5055
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "dheeraj@mailandnews.com"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1320
         TabIndex        =   6
         Top             =   2160
         Width           =   2040
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "For More Information on this project please contact me at "
         Height          =   435
         Left            =   1320
         TabIndex        =   5
         Top             =   1680
         Width           =   3105
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Application Title"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1320
         TabIndex        =   4
         Top             =   120
         Width           =   3405
      End
      Begin VB.Label lblcopy 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Copy Right"
         Height          =   615
         Left            =   1320
         TabIndex        =   3
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label lblPlatform 
         BackColor       =   &H00FFFFFF&
         Caption         =   "platform"
         Height          =   615
         Left            =   1320
         TabIndex        =   2
         Top             =   600
         Width           =   2775
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         Index           =   0
         X1              =   1200
         X2              =   4800
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   2415
         Left            =   0
         Picture         =   "frmabouts.frx":08CA
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.CommandButton OK 
      Caption         =   "O&k"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////
' Frmabouts.frm
' Splits and joins files, with optional compression
' Original by Dheeraj Khajuria Copyright 2002
'
'//////////////////////////////////////////////

Option Explicit

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblcopy.Caption = App.LegalCopyright & vbCrLf & App.FileDescription
    lblPlatform.Caption = "Windows 95/98/Me/2000/NT"
    lblTitle.Caption = " " & App.Title & App.Major & "." & Format(App.Minor, "00") & "." & Format(App.Revision, "00")
End Sub

Private Sub OK_Click()
 Unload Me
End Sub

