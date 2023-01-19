VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ucShellBrowse Paired Controls"
   ClientHeight    =   2430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6150
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
   ScaleHeight     =   2430
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Button 1"
      Height          =   360
      Left            =   195
      TabIndex        =   3
      Top             =   1215
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Button2"
      Height          =   360
      Left            =   1275
      TabIndex        =   2
      Top             =   1215
      Width           =   990
   End
   Begin ShellBrowseDemoDual.ucShellBrowse ucShellBrowse2 
      Height          =   2325
      Left            =   2730
      TabIndex        =   1
      Top             =   45
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4101
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
      ListControlBox  =   0   'False
   End
   Begin ShellBrowseDemoDual.ucShellBrowse ucShellBrowse1 
      Height          =   405
      Left            =   165
      TabIndex        =   0
      Top             =   105
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   714
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
      ListControlBox  =   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Some explanatory text"
      Height          =   195
      Left            =   510
      TabIndex        =   4
      Top             =   675
      Width           =   1650
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bManualChanging As Boolean

Private Sub Command1_Click()
ucShellBrowse2.StatusSinglePanel = False
End Sub

Private Sub Command2_Click()
ucShellBrowse2.StatusSinglePanel = True
End Sub

Private Sub Form_Resize()
ucShellBrowse2.Height = Me.Height - 750
ucShellBrowse2.Width = Me.Width - 3195
End Sub

Private Sub ucShellBrowse1_DirectoryChanged(ByVal sFullPath As String, siItem As oleexp.IShellItem, pidlFQ As Long)
    ucShellBrowse2.BrowserPath = sFullPath
End Sub

Private Sub ucShellBrowse1_FilterFolder(ByVal sName As String, ByVal sPath As String, fShow As Long)
Debug.Print "Filter " & sPath
If sPath = "C:\" Then fShow = 0
End Sub

Private Sub ucShellBrowse2_DirectoryChanged(ByVal sFullPath As String, siItem As oleexp.IShellItem, pidlFQ As Long)
    ucShellBrowse1.BrowserPath = sFullPath
End Sub
