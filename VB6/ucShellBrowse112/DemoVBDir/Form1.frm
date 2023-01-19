VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ShellBrowse As VB DirListBox"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6690
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
   ScaleHeight     =   3885
   ScaleWidth      =   6690
   StartUpPosition =   2  'CenterScreen
   Begin Project1.ucShellBrowse ucShellBrowse1 
      Height          =   3540
      Left            =   3165
      TabIndex        =   2
      Top             =   270
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   6244
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ControlType     =   2
      FoldersOnly     =   -1  'True
      ShowParentTreeInList=   -1  'True
      GroupSubsetLinkText=   ""
      ListControlBox  =   0   'False
      DefaultColumns  =   ""
      ShowStatusBar   =   0   'False
      BrowseZip       =   0   'False
      LockColumns     =   1
      HideColumnHeader=   -1  'True
      SimpleSelect    =   0   'False
   End
   Begin VB.DirListBox Dir1 
      Height          =   3465
      Left            =   15
      TabIndex        =   0
      Top             =   345
      Width           =   2970
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ucShellBrowse"
      Height          =   195
      Left            =   3240
      TabIndex        =   3
      Top             =   105
      Width           =   1905
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VB"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   90
      Width           =   675
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Dir1_Change()
ucShellBrowse1.BrowserPath = Dir1.Path
End Sub

Private Sub Form_Load()
'ucShellBrowse1.BrowserPath = Dir1.Path
Dir1.Path = App.Path
End Sub

Private Sub ucShellBrowse1_DirectoryChanged(ByVal sFullPath As String, siItem As oleexp.IShellItem, pidlFQ As Long)
On Error Resume Next
Dir1.Path = sFullPath
End Sub
