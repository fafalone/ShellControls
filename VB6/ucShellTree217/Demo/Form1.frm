VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ucShellTree Demo"
   ClientHeight    =   5685
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3555
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   3555
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Open To"
      Height          =   360
      Left            =   1440
      TabIndex        =   2
      Top             =   60
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print Checked"
      Height          =   360
      Left            =   75
      TabIndex        =   1
      Top             =   60
      Width           =   1230
   End
   Begin ShellTreeDemo.ucShellTree ucShellTree1 
      Height          =   5040
      Left            =   75
      TabIndex        =   0
      Top             =   600
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   8890
      InfoTipOnFolders=   -1  'True
      SingleClickExpand=   -1  'True
      ShowFavorites   =   0   'False
      Multiselect     =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim sPaths() As String
Dim vPaths As Variant
If ucShellTree1.Checkboxes Then
    vPaths = ucShellTree1.CheckedPaths
    
    Dim i As Long
    For i = 0 To UBound(vPaths)
        Debug.Print vPaths(i)
    Next
Else
    Debug.Print ucShellTree1.SelectedItem
End If
End Sub

Private Sub Command2_Click()
Dim sPath As String
ucShellTree1.FullRowSelect = False
sPath = "C:\Program Files" 'Change to whatever. Can use network paths like \\SERVER\share
ucShellTree1.OpenToPath sPath, False
End Sub

Private Sub Form_Initialize()
'ucShellTree1.CustomRoot = "C:\"
ucShellTree1.OpenToPath App.Path, True
End Sub

Private Sub Form_Resize()
On Error Resume Next
ucShellTree1.Width = Me.Width - 390
ucShellTree1.Height = Me.Height - 1240
End Sub

  
Private Sub ucShellTree1_ItemSelect(sName As String, sFullPath As String, bFolder As Boolean, hItem As Long)
Debug.Print "ItemSelect " & sFullPath

End Sub

