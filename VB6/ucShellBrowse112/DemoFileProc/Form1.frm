VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "DemoFileProc"
   ClientHeight    =   5985
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10530
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
   ScaleHeight     =   5985
   ScaleWidth      =   10530
   StartUpPosition =   2  'CenterScreen
   Begin DemoFileProc.ucShellBrowse ucShellBrowse1 
      Height          =   5445
      Left            =   2925
      TabIndex        =   6
      Top             =   60
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   9604
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
      MultiSelect     =   0   'False
      ListControlBox  =   0   'False
      DefaultColumns  =   ""
      ShowStatusBar   =   0   'False
      ThumbnailSize   =   216
      DetailsPane     =   0   'False
      AutosizeColumns =   -1  'True
      LockColumns     =   2
      EnableSearch    =   0   'False
      EnableLayout    =   0   'False
      LockView        =   -1  'True
      HideColumnHeader=   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add file"
      Height          =   360
      Left            =   6450
      TabIndex        =   4
      Top             =   105
      Width           =   990
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   6465
      TabIndex        =   3
      Top             =   570
      Width           =   3915
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Process"
      Height          =   495
      Left            =   7335
      TabIndex        =   2
      Top             =   3030
      Width           =   1905
   End
   Begin VB.TextBox Text1 
      Height          =   1695
      Left            =   6465
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   3630
      Width           =   3990
   End
   Begin DemoFileProc.ucShellTree ucShellTree1 
      Height          =   5370
      Left            =   30
      TabIndex        =   0
      Top             =   105
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   9472
      PlayNavigationSound=   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "or double-click one"
      Height          =   255
      Left            =   7500
      TabIndex        =   5
      Top             =   180
      Width           =   1545
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sQueue() As String
Private nQ As Long

Private Sub Command1_Click()
If ucShellBrowse1.SelectedFile <> "" Then
    List1.AddItem ucShellBrowse1.SelectedFile 'Add just the filename
    ReDim Preserve sQueue(nQ)
    sQueue(nQ) = ucShellBrowse1.SelectedFilePath 'Add the full path to the file
    nQ = nQ + 1
End If
End Sub

Private Sub Command2_Click()
Dim i As Long
For i = 0 To nQ - 1
    Text1.Text = Text1.Text & "Reticulating " & sQueue(i) & "..." & vbCrLf
Next i
End Sub

Private Sub Form_Load()
ReDim sQueue(0)
ucShellTree1.OpenToItem ucShellBrowse1.BrowserPathItem, False
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.Width < 7800 Then Me.Width = 7800
If Me.Height < 5400 Then Me.Height = 5400
'ucShellBrowse2.Top = Me.Height - (ucShellBrowse2.Height + 200) - 340
'ucShellBrowse1.Height = (Me.Height) - ucShellBrowse2.Height - 590
'ucShellTree1.Height = (Me.Height) - ucShellBrowse2.Height - 640
List1.Width = Me.Width - (ucShellBrowse1.Width + ucShellTree1.Width) - 403
Text1.Width = Me.Width - (ucShellBrowse1.Width + ucShellTree1.Width) - 403
Text1.Height = Me.Height - Command2.Top - 1300
End Sub

Private Sub ucShellBrowse1_DirectoryChanged(ByVal sFullPath As String, siItem As oleexp.IShellItem, pidlFQ As Long)
ucShellTree1.OpenToItem siItem, False
End Sub

Private Sub ucShellBrowse1_FileExecute(ByVal sFile As String, siFile As oleexp.IShellItem)
List1.AddItem ucShellBrowse1.SelectedFile 'Add just the filename
ReDim Preserve sQueue(nQ)
sQueue(nQ) = ucShellBrowse1.SelectedFilePath 'Add the full path to the file
nQ = nQ + 1
End Sub

Private Sub ucShellTree1_ItemSelectByShellItem(siItem As oleexp.IShellItem, sName As String, sFullPath As String, bFolder As Boolean, hItem As Long)
ucShellBrowse1.BrowserOpenItem siItem
End Sub


