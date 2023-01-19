VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ucShellTree with ucShellBrowse Demo"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8040
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
   ScaleHeight     =   6375
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin Project1.ucShellBrowse ucShellBrowse1 
      Height          =   6285
      Left            =   2430
      TabIndex        =   1
      Top             =   -90
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   11086
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
      DetailsPane     =   -1  'True
   End
   Begin Project1.ucShellTree ucShellTree1 
      Height          =   6390
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   11271
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bChanging As Boolean

Private Sub Form_Load()
'ucShellTree1.OpenToPath App.Path, False '<--- Will not work until after load
ucShellTree1.InitialPath = App.Path 'The 1st dir change from ucShellBrowse is during load, so it won't work to set this either
End Sub

Private Sub Form_Resize()
ucShellTree1.Height = Me.Height - 580
ucShellBrowse1.Height = Me.Height - 580
ucShellBrowse1.Width = Me.Width - 2640
End Sub

Private Sub ucShellBrowse1_DirectoryChanged(ByVal sFullPath As String, siItem As oleexp.IShellItem, pidl As Long)

    ucShellTree1.OpenToItem siItem, False

End Sub

 
Private Sub ucShellTree1_ItemSelect(sName As String, sFullPath As String, bFolder As Boolean, hItem As Long)
    If bFolder Then
        'Opening with the path name won't work for phones/cameras/etc on Windows 10 because of an API bug.
'        ucShellBrowse1.BrowserPath = sFullPath
        'So we'll keep the event trigger but use the shell item instead
        ucShellBrowse1.BrowserOpenItem ucShellTree1.SelectedShellItem
    End If
End Sub

