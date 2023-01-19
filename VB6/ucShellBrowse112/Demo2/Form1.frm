VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ucShellBrowse Simple Use"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3885
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
   ScaleHeight     =   2850
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin ShellBrowseDemoMini.ucShellBrowse ucShellBrowse1 
      Height          =   2730
      Left            =   60
      TabIndex        =   0
      Top             =   15
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   4815
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ComboType       =   2
      DefaultColumns  =   ""
      StatusSinglePanel=   -1  'True
      ViewButton      =   0   'False
      LockColumns     =   2
      HideColumnHeader=   -1  'True
      ComboCanEdit    =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

