VERSION 5.00
Begin VB.Form frmDirOnly 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Directory Dropdown Only"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   6000
   StartUpPosition =   1  'CenterOwner
   Begin ShellBrowseDemo.ucShellBrowse ucShellBrowse2 
      Height          =   630
      Left            =   45
      TabIndex        =   4
      Top             =   945
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   1111
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ControlType     =   4
      GroupSubsetLinkText=   ""
      ViewButton      =   0   'False
      DetailsPane     =   0   'False
      SimpleSelect    =   0   'False
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test"
      Height          =   360
      Left            =   75
      TabIndex        =   3
      Top             =   45
      Width           =   990
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   30
      TabIndex        =   2
      Top             =   2685
      Width           =   5925
   End
   Begin ShellBrowseDemo.ucShellBrowse ucShellBrowse1 
      Height          =   540
      Left            =   120
      TabIndex        =   0
      Top             =   435
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   953
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ControlType     =   1
      ComboType       =   2
      GroupSubsetLinkText=   ""
      DetailsPane     =   0   'False
      ComboCanEdit    =   0   'False
      SimpleSelect    =   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Events"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   2445
      Width           =   495
   End
End
Attribute VB_Name = "frmDirOnly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
End Sub

Private Sub ucShellBrowse1_DirectoryChanged(ByVal sFullPath As String, siItem As oleexp.IShellItem, pidlFQ As Long)
List1.AddItem "Directory changed to " & sFullPath
End Sub

