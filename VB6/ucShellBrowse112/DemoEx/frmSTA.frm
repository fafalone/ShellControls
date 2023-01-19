VERSION 5.00
Begin VB.Form frmSTA 
   Caption         =   "ShellTree Alone"
   ClientHeight    =   7185
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4260
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
   ScaleHeight     =   7185
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Open to path"
      Height          =   270
      Left            =   2895
      TabIndex        =   4
      Top             =   525
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   135
      TabIndex        =   3
      Top             =   525
      Width           =   2655
   End
   Begin ShellBrowseDemoEx.ucShellTree ucShellTree1 
      Height          =   6270
      Left            =   75
      TabIndex        =   2
      Top             =   870
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   11060
      SingleExpand    =   -1  'True
      FullRowSelect   =   0   'False
      InfoTipOnFolders=   -1  'True
      ShowFavorites   =   0   'False
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   360
      Left            =   1215
      TabIndex        =   1
      Top             =   90
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Files"
      Height          =   360
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   990
   End
End
Attribute VB_Name = "frmSTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bSF As Boolean

Private Sub Command1_Click()
If bSF Then
    ucShellTree1.ShowFiles = False
    Command1.Caption = "Show Files"
    bSF = False
Else
    ucShellTree1.ShowFiles = True
    Command1.Caption = "Hide Files"
    bSF = True
End If
End Sub

Private Sub Command2_Click()
ucShellTree1.RefreshTreeView
 
End Sub

Private Sub Command3_Click()
Dim siPath As oleexp.IShellItem

oleexp.SHCreateItemFromParsingName StrPtr(Text1.Text), Nothing, IID_IShellItem, siPath
If (siPath Is Nothing) = False Then
    ucShellTree1.OpenToItem siPath, True
End If

End Sub

Private Sub Form_Load()
Text1.Text = App.Path
End Sub

Private Sub Form_Resize()
On Error Resume Next
ucShellTree1.Width = Me.Width - (24 * Screen.TwipsPerPixelX)
ucShellTree1.Height = Me.Height - (100 * Screen.TwipsPerPixelY)
End Sub


Private Function IID_IShellItem() As oleexp.UUID
Static iid As UUID
If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H43826D1E, CInt(&HE718), CInt(&H42EE), &HBC, &H55, &HA1, &HE2, &H61, &HC3, &H7B, &HFE)
IID_IShellItem = iid
End Function
Private Sub DEFINE_UUID(Name As oleexp.UUID, l As Long, w1 As Integer, w2 As Integer, B0 As Byte, b1 As Byte, b2 As Byte, B3 As Byte, b4 As Byte, b5 As Byte, b6 As Byte, b7 As Byte)
  With Name
    .Data1 = l: .Data2 = w1: .Data3 = w2:
    .Data4(0) = B0: .Data4(1) = b1: .Data4(2) = b2: .Data4(3) = B3: .Data4(4) = b4: .Data4(5) = b5: .Data4(6) = b6: .Data4(7) = b7
  End With
End Sub
