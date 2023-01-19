VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ShellBrowse as VB Controls"
   ClientHeight    =   10575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7410
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
   ScaleHeight     =   10575
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Full Window"
      Height          =   3735
      Left            =   45
      TabIndex        =   16
      Top             =   6780
      Width           =   7215
      Begin Project1.ucShellBrowse ucShellBrowse4 
         Height          =   3360
         Left            =   120
         TabIndex        =   17
         Top             =   225
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   5927
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GroupSubsetLinkText=   ""
         ViewButton      =   0   'False
         SearchBox       =   -1  'True
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "FileListBox"
      Height          =   2625
      Left            =   45
      TabIndex        =   10
      Top             =   4125
      Width           =   7305
      Begin VB.CheckBox Check1 
         Caption         =   "Icons"
         Height          =   195
         Left            =   3240
         TabIndex        =   15
         Top             =   600
         Value           =   1  'Checked
         Width           =   870
      End
      Begin Project1.ucShellBrowse ucShellBrowse3 
         Height          =   2220
         Left            =   4170
         TabIndex        =   14
         Top             =   285
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   3942
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
         FilesOnly       =   -1  'True
         GroupSubsetLinkText=   ""
         ListControlBox  =   0   'False
         DefaultColumns  =   ""
         ShowStatusBar   =   0   'False
         BookmarkButton  =   -1  'True
         DetailsPane     =   0   'False
         LockColumns     =   2
         EnableViewMenu  =   0   'False
         HideColumnHeader=   -1  'True
         ActiveDropHoverTime=   2300
      End
      Begin VB.FileListBox File1 
         Height          =   2235
         Left            =   360
         TabIndex        =   12
         Top             =   255
         Width           =   2820
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ucSB"
         Height          =   195
         Left            =   3405
         TabIndex        =   13
         Top             =   330
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VB"
         Height          =   195
         Left            =   105
         TabIndex        =   11
         Top             =   330
         Width           =   180
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "DirListBox"
      Height          =   3225
      Left            =   45
      TabIndex        =   5
      Top             =   810
      Width           =   7320
      Begin Project1.ucShellBrowse ucShellBrowse2 
         Height          =   2790
         Left            =   3930
         TabIndex        =   9
         Top             =   270
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   4921
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
         ThumbnailPreload=   -1  'True
         NavigationButtons=   0
         DetailsPane     =   0   'False
         LockColumns     =   2
         EnableShellMenu =   0   'False
         HideColumnHeader=   -1  'True
         HeaderNoSizing  =   -1  'True
      End
      Begin VB.DirListBox Dir1 
         Height          =   2790
         Left            =   585
         TabIndex        =   7
         Top             =   270
         Width           =   2685
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ucSB"
         Height          =   195
         Left            =   3405
         TabIndex        =   8
         Top             =   270
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VB"
         Height          =   195
         Left            =   270
         TabIndex        =   6
         Top             =   315
         Width           =   180
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "DriveListBox"
      Height          =   660
      Left            =   30
      TabIndex        =   0
      Top             =   105
      Width           =   7320
      Begin Project1.ucShellBrowse ucShellBrowse1 
         Height          =   510
         Left            =   3870
         TabIndex        =   4
         Top             =   120
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   900
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ControlType     =   3
         DetailsPane     =   0   'False
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   600
         TabIndex        =   2
         Top             =   240
         Width           =   2115
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ucSB"
         Height          =   195
         Left            =   3435
         TabIndex        =   3
         Top             =   270
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VB"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   300
         Width           =   180
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
If Check1.value = vbChecked Then
    ucShellBrowse3.HideIcons = False
Else
    ucShellBrowse3.HideIcons = True
End If
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub


Private Sub ucShellBrowse1_DirectoryChanged(ByVal sFullPath As String, siItem As oleexp.IShellItem, pidlFQ As Long)
ucShellBrowse2.BrowserOpenItem siItem
End Sub

Private Sub ucShellBrowse2_DirectoryChanged(ByVal sFullPath As String, siItem As oleexp.IShellItem, pidlFQ As Long)
ucShellBrowse3.BrowserOpenItem siItem
End Sub

Private Sub ucShellBrowse3_DirectoryChanged(ByVal sFullPath As String, siItem As oleexp.IShellItem, pidlFQ As Long)
ucShellBrowse2.BrowserOpenItem siItem
End Sub
