VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ShellBrowse Thumbnail Viewer"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7740
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin Project1.ucShellBrowse ucShellBrowse1 
      Height          =   6090
      Left            =   45
      TabIndex        =   0
      Top             =   -420
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   10742
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EnableUserModeSwitch=   0   'False
      ControlType     =   2
      FilesOnly       =   -1  'True
      BrowserPath     =   "%USERPROFILE%\Pictures"
      ViewMode        =   6
      GroupSubsetLinkText=   ""
      ShowStatusBar   =   0   'False
      DetailsPane     =   0   'False
      LockNavigation  =   -1  'True
      EnablePreview   =   0   'False
      EnableDetails   =   0   'False
      EnableSearch    =   0   'False
      EnableStatusBar =   0   'False
      EnableLayout    =   0   'False
      EnableBookmarks =   0   'False
      EnableNewFolder =   0   'False
      LockView        =   -1  'True
      FollowLinks     =   0   'False
      SimpleSelect    =   0   'False
      AllowPasting    =   0   'False
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0000
      Height          =   585
      Left            =   60
      TabIndex        =   2
      Top             =   6570
      Width           =   7725
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":010E
      Height          =   780
      Left            =   75
      TabIndex        =   1
      Top             =   5760
      Width           =   7530
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'If you haven't set it during design time.
'ucShellBrowse1.BrowserPath = "C:\JPG"
End Sub
