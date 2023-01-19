VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ShellBrowseDemoEx"
   ClientHeight    =   7005
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11700
   DrawMode        =   14  'Copy Pen
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
   ScaleHeight     =   467
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   780
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "Toggle FilesOnly"
      Height          =   345
      Left            =   8400
      TabIndex        =   9
      Top             =   15
      Width           =   1650
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Toggle Control Bar"
      Height          =   360
      Left            =   210
      TabIndex        =   8
      Top             =   15
      Width           =   1710
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Shell Tree Alone"
      Height          =   345
      Left            =   6990
      TabIndex        =   7
      Top             =   15
      Width           =   1395
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Custom Dir"
      Height          =   345
      Left            =   5940
      TabIndex        =   6
      Top             =   15
      Width           =   1020
   End
   Begin VB.PictureBox pbPreviewSizer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   2820
      MousePointer    =   9  'Size W E
      ScaleHeight     =   337
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   2
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   495
      Width           =   30
   End
   Begin ShellBrowseDemoEx.ucShellTree ucShellTree1 
      Height          =   5910
      Left            =   30
      TabIndex        =   4
      Top             =   480
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   10425
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      InfoTipOnFiles  =   0   'False
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Output Dir List"
      Height          =   360
      Left            =   4545
      TabIndex        =   3
      Top             =   15
      Width           =   1365
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Footer Demo 2"
      Height          =   360
      Left            =   3240
      TabIndex        =   2
      Top             =   15
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Footer Demo 1"
      Height          =   360
      Left            =   1920
      TabIndex        =   1
      Top             =   15
      Width           =   1275
   End
   Begin ShellBrowseDemoEx.ucShellBrowse ucShellBrowse1 
      Height          =   6450
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   11377
      BrowserStartBlank=   -1  'True
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFileControls {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListViewOffsetX =   184
      DetailsPaneOffsetX=   184
      Picture         =   "Form1.frx":0000
      PictureAlignment=   5
      BackColor       =   -2147483645
      ForeColor       =   8388608
      ControlType     =   2
      FullRowSelect   =   -1  'True
      BookmarkButton  =   -1  'True
      SearchBox       =   -1  'True
      ShellTreeInLayout=   -1  'True
      PlaySounds      =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Note: mDemo has been removed since the ComCtl init routine isn't needed on Vista+, which is required for this project.
'
'Note: cLog has been removed from this demo. Write-to-file debugging is now built
'      into the control, see the User Options section below the changelog in
'      ucShellBrowse.ctl

Private Const dbg_IncludeDate As Boolean = True

Private mDX As Single
Private hIML_Footer As Long
Private bChanging As Boolean
Private bPMove As Boolean

Private Const CB_GETDROPPEDWIDTH = &H15F
Private Const HTCAPTION = 2
Private Const HWND_TOP = 0&
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_FRAMECHANGED = &H20

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExW" (ByVal lpszFile As Long, _
                                                                                ByVal nIconIndex As Long, _
                                                                                phiconLarge As Long, _
                                                                                phiconSmall As Long, _
                                                                                ByVal nIcons As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Declare Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long) As Long
Private Declare Function PSFormatPropertyValue Lib "propsys.dll" (ByVal pps As Long, ByVal ppd As Long, ByVal pdff As PROPDESC_FORMAT_FLAGS, ppszDisplay As Long) As Long
Private Declare Function PSGetPropertyDescription Lib "propsys.dll" (PropKey As oleexp.PROPERTYKEY, riid As UUID, ppv As Any) As Long
Private Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function ImageList_Add Lib "comctl32.dll" (ByVal himl As Long, ByVal hbmImage As Long, ByVal hBMMask As Long) As Long
Private Declare Function ImageList_AddMasked Lib "comctl32.dll" (ByVal himl As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long
Private Declare Function ImageList_BeginDrag Lib "comctl32.dll" (ByVal himlTrack As Long, ByVal iTrack As Long, ByVal dxHotspot As Long, ByVal dyHotspot As Long) As Boolean
Private Declare Function ImageList_Copy Lib "comctl32.dll" (ByVal himlDst As Long, ByVal iDst As Long, ByVal himlSrc As Long, ByVal iSrc As Long, ByVal uFlags As IL_CopyFlags) As Boolean
Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal CX As Long, ByVal cy As Long, ByVal Flags As IL_CreateFlags, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Boolean
Private Declare Function ImageList_DragEnter Lib "comctl32.dll" (ByVal hWndLock As Long, ByVal X As Long, ByVal Y As Long) As Boolean
Private Declare Function ImageList_DragLeave Lib "comctl32.dll" (ByVal hWndLock As Long) As Boolean
Private Declare Function ImageList_DragMove Lib "comctl32.dll" (ByVal X As Long, ByVal Y As Long) As Boolean
Private Declare Function ImageList_DragShowNolock Lib "comctl32.dll" (ByVal fShow As Boolean) As Boolean
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal fStyle As IL_DrawStyle) As Boolean
Private Declare Function ImageList_DrawEx Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal rgbBk As Long, ByVal rgbFg As Long, ByVal fStyle As Long) As Boolean
Private Declare Function ImageList_DrawIndirect Lib "comctl32.dll" (pimldp As IMAGELISTDRAWPARAMS) As Boolean
Private Declare Function ImageList_Duplicate Lib "comctl32.dll" (ByVal himl As Long) As Long

Private Declare Sub ImageList_EndDrag Lib "comctl32.dll" ()
Private Declare Function ImageList_GetBkColor Lib "comctl32.dll" (ByVal himl As Long) As Long
Private Declare Function ImageList_GetDragImage Lib "comctl32.dll" (ppt As POINT, pptHotspot As POINT) As Long
Private Declare Function ImageList_GetIconSize Lib "comctl32.dll" (ByVal himl As Long, lpcx As Long, lpcy As Long) As Boolean
Private Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal himl As Long) As Long

Private Declare Function ImageList_GetImageInfo Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, pImageInfo As IMAGEINFO) As Boolean
Private Declare Function ImageList_Merge Lib "comctl32.dll" (ByVal himl1 As Long, ByVal i1 As Long, ByVal himl2 As Long, ByVal i2 As Long, ByVal dx As Long, ByVal dy As Long) As Long

Private Declare Function ImageList_SetIconSize Lib "comctl32.dll" (ByVal himl As Long, ByVal CX As Long, ByVal cy As Long) As Boolean
Private Declare Function ImageList_SetImageCount Lib "comctl32.dll" (ByVal himl As Long, ByVal uNewCount As Long) As Boolean
Private Declare Function ImageList_SetBkColor Lib "comctl32.dll" (ByVal himl As Long, ByVal clrBk As Long) As Long
Private Declare Function ImageList_SetDragCursorImage Lib "comctl32.dll" (ByVal himlDrag As Long, ByVal iDrag As Long, ByVal dxHotspot As Long, ByVal dyHotspot As Long) As Boolean
Private Declare Function ImageList_SetOverlayImage Lib "comctl32.dll" (ByVal himl As Long, ByVal iImage As Long, ByVal iOverlay As Long) As Boolean

Private Declare Function ImageList_CoCreateInstance Lib "comctl32.dll " (refclsid As UUID, ByVal pUnkOuter As Long, riid As UUID, ppv As Any) As Long
Private Declare Function HIMAGELIST_QueryInterface Lib "comctl32.dll" (ByVal himl As Long, riid As UUID, ppv As Any) As Long


Private Declare Sub SysFreeString Lib "oleaut32" (ByVal lpbstr As Long)
 Private Declare Function GetShortPathNameW Lib "kernel32.dll" (ByVal lpszLongPath As Long, Optional ByVal lpszShortPath As Long, Optional ByVal cchBuffer As Long) As Long

Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal lpString As Long) As Long
Private Declare Function PathMakeUniqueName Lib "shell32.dll" (ByVal pszUniqueName As Long, ByVal cchMax As Long, ByVal pszTemplate As Long, Optional ByVal pszLongPlate As Long, Optional ByVal pszDir As Long) As Long
Private Declare Function SysReAllocStringLen Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long, Optional ByVal Length As Long) As Long

Private Declare Function ImageList_Read Lib "comctl32.dll" (pstm As IStream) As Long
Private Declare Function ImageList_Write Lib "comctl32.dll" (ByVal himl As Long, pstm As IStream) As Boolean

Private Enum IL_DrawStyle
  ILD_NORMAL = &H0
  ILD_TRANSPARENT = &H1
  ILD_MASK = &H10
  ILD_IMAGE = &H20
'#If (WIN32_IE >= &H300) Then
  ILD_ROP = &H40
'#End If
  ILD_BLEND25 = &H2
  ILD_BLEND50 = &H4
  ILD_OVERLAYMASK = &HF00
 
  ILD_SELECTED = ILD_BLEND50
  ILD_FOCUS = ILD_BLEND25
  ILD_BLEND = ILD_BLEND50
End Enum

#If False Then
Dim ILD_NORMAL, ILD_TRANSPARENT, ILD_MASK, ILD_IMAGE, ILD_ROP, ILD_BLEND25, ILD_BLEND50, _
ILD_OVERLAYMASK, ILD_SELECTED, ILD_FOCUS, ILD_BLEND
#End If
Private Enum ImageListColor_flags
  CLR_NONE = &HFFFFFFFF
  CLR_DEFAULT = &HFF000000
  CLR_HILIGHT = CLR_DEFAULT
End Enum

Private Type IMAGELISTDRAWPARAMS
  cbSize As Long
  himl As Long
  i As Long
  hdcDst As Long
  X As Long
  Y As Long
  CX As Long
  cy As Long
  xBitmap As Long   ' x offest from the upperleft of bitmap
  yBitmap As Long   ' y offset from the upperleft of bitmap
  rgbBk As Long
  rgbFg As Long
  fStyle As IL_DrawStyle
  dwRop As Long
End Type
 
Private Enum IL_CreateFlags
  ILC_MASK = &H1
  ILC_COLOR = &H0
  ILC_COLORDDB = &HFE
  ILC_COLOR4 = &H4
  ILC_COLOR8 = &H8
  ILC_COLOR16 = &H10
  ILC_COLOR24 = &H18
  ILC_COLOR32 = &H20
  ILC_PALETTE = &H800                  ' (no longer supported...never worked anyway)
  '5.0
  ILC_MIRROR = &H2000
  ILC_PERITEMMIRROR = &H8000
  '6.0
  ILC_ORIGINALSIZE = &H10000
  ILC_HIGHQUALITYSCALE = &H20000
End Enum

 
 
Private Enum LoadResourceFlags
  LR_DEFAULTCOLOR = &H0
  LR_MONOCHROME = &H1
  LR_COLOR = &H2
  LR_COPYRETURNORG = &H4
  LR_COPYDELETEORG = &H8
  LR_LOADFROMFILE = &H10
  LR_LOADTRANSPARENT = &H20
  LR_DEFAULTSIZE = &H40
  LR_VGACOLOR = &H80
  LR_LOADMAP3DCOLORS = &H1000
  LR_CREATEDIBSECTION = &H2000
  LR_COPYFROMRESOURCE = &H4000
  LR_SHARED = &H8000&
End Enum
Private Enum ImageTypes
  IMAGE_BITMAP = 0
  IMAGE_ICON = 1
  IMAGE_CURSOR = 2
  IMAGE_ENHMETAFILE = 3
End Enum
Private Enum IL_CopyFlags
  ILCF_MOVE = &H0
  ILCF_SWAP = &H1
End Enum
Private Type IMAGEINFO
  hbmImage As Long
  hBMMask As Long
  Unused1 As Long
  Unused2 As Long
  rcImage As RECT
End Type

Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Private Declare Function ImageList_GetIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal Flags As IL_DrawStyle) As Long

Private Declare Function ImageList_Remove Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long) As Boolean

Private Declare Function ImageList_Replace Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hbmImage As Long, ByVal hBMMask As Long) As Boolean
Private Declare Function ImageList_ReplaceIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hIcon As Long) As Long

Private Declare Function ImageList_LoadImage Lib "comctl32.dll" Alias "ImageList_LoadImageW" (ByVal hi As Long, ByVal lpbmp As Long, ByVal CX As Long, ByVal cGrow As Long, ByVal crMask As Long, ByVal uType As ImageTypes, ByVal uFlags As LoadResourceFlags) As Long





Private Function IsIDE() As Boolean
'  On Error GoTo Out
'  Debug.Print 1 / 0
'Out:
'  IsIDE = Err
'IsIDE = gide
   Dim buff As String
   Dim Success As Long
   
   buff = Space$(MAX_PATH)
   Success = GetModuleFileName(App.hInstance, buff, Len(buff))
   
   If Success > 0 Then
     'Change the VB exe name here as appropriate
     'for your version. The case change ensures this
     'works regardless as to how the exe is cased on
     'the machine.
      IsIDE = InStr(LCase$(buff), "vb6.exe") > 0
   End If

End Function

' Macros

Private Function ImageList_AddIcon(himl As Long, hIcon As Long) As Long
  ImageList_AddIcon = ImageList_ReplaceIcon(himl, -1, hIcon)
End Function

Private Function ImageList_RemoveAll(himl As Long) As Boolean
  ImageList_RemoveAll = ImageList_Remove(himl, -1)
End Function

' HINSTANCE hi - This parameter is not used and should always be zero.

Private Function ImageList_ExtractIcon(hi As Long, himl As Long, i As Long) As Long
  ImageList_ExtractIcon = ImageList_GetIcon(himl, i, 0)
  Dim X As Long
  X = ILCF_MOVE
  
End Function

Private Function ImageList_LoadBitmap(hi As Long, lpbmp As Long, CX As Long, cGrow As Long, crMask As Long) As Long
  ImageList_LoadBitmap = ImageList_LoadImage(hi, lpbmp, CX, cGrow, crMask, IMAGE_BITMAP, 0)
End Function







Public Function LPWSTRtoStr(lPtr As Long, Optional ByVal fFree As Boolean = True) As String
SysReAllocString VarPtr(LPWSTRtoStr), lPtr
If fFree Then
    Call CoTaskMemFree(lPtr)
End If
End Function
Private Sub DebugAppend(ByVal sMsg As String, Optional ilvl As Integer = 0)
If dbg_IncludeDate Then sMsg = "[F1]" & "[" & Format$(Now, "yyyy-mm-dd Hh:Mm:Ss") & "] " & sMsg
Debug.Print sMsg
End Sub
Private Sub Command1_Click()
Dim tBtnS() As String
Dim tBtnI() As Long
Dim tBtnL() As Long

ReDim tBtnS(2)
ReDim tBtnI(2)
ReDim tBtnL(2)

tBtnS(0) = "Button"
tBtnI(0) = 0
tBtnL(0) = 1000

tBtnS(1) = "Test"
tBtnI(1) = 1
tBtnL(1) = 1001

tBtnS(2) = "Set"
tBtnI(2) = 2
tBtnL(2) = 1002

 ucShellBrowse1.FooterCreate "Hello Footer!", hIML_Footer, tBtnS, tBtnI, tBtnL

End Sub

Private Sub Command2_Click()
 
Dim tBtnS() As String
Dim tBtnI() As String
Dim tBtnL() As Long

ReDim tBtnS(2)
ReDim tBtnI(2)
ReDim tBtnL(2)

tBtnS(0) = "Button"
tBtnI(0) = "ICO_FOOT1"
tBtnL(0) = 1000

tBtnS(1) = "Test"
tBtnI(1) = "ICO_FOOT2"
tBtnL(1) = 1001

tBtnS(2) = "Set"
tBtnI(2) = "ICO_FOOT3"
tBtnL(2) = 1002

 ucShellBrowse1.FooterCreate "Hello Footer!", 0, tBtnS, tBtnI, tBtnL
 
End Sub
 
Private Sub Command3_Click()
Dim pArSel As IShellItemArray
Dim siChild As IShellItem, si2 As IShellItem2
Dim pst As IPropertyStore
Dim nItems As Long
Dim lpsz As Long, sTmp As String
Dim sOut As String
Dim dwAtr As SFGAO_Flags
Dim i As Long
If ucShellBrowse1.ItemsSelectedCount = 0 Then
    Debug.Print "Please select an item first"
    Exit Sub
End If
Set pArSel = ucShellBrowse1.SelectedItems
If pArSel Is Nothing Then
    DebugAppend "Please select all items first"
    Exit Sub
End If

pArSel.GetCount nItems
DebugAppend "         Name                        Size      Date Modified     Attributes"
DebugAppend "---------------------------------------------------------------------------"
For i = 0 To (nItems - 1)
    pArSel.GetItemAt i, siChild
    If siChild Is Nothing Then
        DebugAppend "No child"
        Exit Sub
    End If
    
    Set si2 = siChild
    siChild.GetAttributes SFGAO_FOLDER, dwAtr
    If (dwAtr And SFGAO_FOLDER) = SFGAO_FOLDER Then
        sOut = "<DIR> "
    Else
        sOut = "      "
    End If
    siChild.GetDisplayName SIGDN_NORMALDISPLAY, lpsz
    sTmp = LPWSTRtoStr(lpsz)
    sOut = sOut & sTmp
    If Len(sOut) < 30 Then
        
        sOut = sOut & Space$(30 - Len(sOut))
    Else
        sTmp = Left$(sOut, 27)
        sOut = sTmp & "..."
    End If
    
    si2.GetPropertyStore GPS_OPENSLOWITEM Or GPS_BESTEFFORT, IID_IPropertyStore, pst
    
    sTmp = GetPropertyKeyDisplayString(pst, PKEY_Size)
    sTmp = Space$(13 - Len(sTmp)) & sTmp & "  "
    sOut = sOut & sTmp
        
    sTmp = GetPropertyKeyDisplayString(pst, PKEY_DateModified)
    sTmp = sTmp & Space$(20 - Len(sTmp))
    sOut = sOut & sTmp
        
    sTmp = GetPropertyKeyDisplayString(pst, PKEY_FileAttributes)
    sTmp = "-" & sTmp
    
    sOut = sOut & sTmp
    
    DebugAppend sOut
    Set pst = Nothing
Next i
pArSel.GetPropertyStore GPS_BESTEFFORT Or GPS_OPENSLOWITEM, IID_IPropertyStore, pst
If (pst Is Nothing) = False Then
    sTmp = GetPropertyKeyDisplayString(pst, PKEY_Size)
    Set pst = Nothing
Else
    sTmp = ""
End If
DebugAppend "---------------------------------------------------------------------------"
DebugAppend "                             " & nItems & " items                 " & sTmp
'<EhFooter>
Exit Sub

e0:
    DebugAppend "Form1.Command3_Click->Error: " & Err.Description & ", 0x" & Hex$(Err.Number)
'</EhFooter>
End Sub


Private Sub Command4_Click()
''What you do here is file the list with the full paths to the set of files
''you want to display. You can use any path resolvable by SHCreateItemFromParsingName,
''which includes network paths and virtual locations with GUIDs as path
''The last parameter, for the icon, -1 sets it to the Saved Searches icon, 0 sets
''it to the default folder icon, and a positive value uses that index in the system
''imagelist.
'
'Dim sLst() As String
'ReDim sLst(9)
'
'sLst(0) = "C:\vb6\cgi\VBCGI.vbg"
'sLst(1) = "C:\vb6\cgi-1\VBCGI.vbg"
'sLst(2) = "C:\vb6\cIcoMenu\PMenuTst.vbg"
'sLst(3) = "C:\vb6\ClassicSearch2\ClassicSearch2.vbg"
'sLst(4) = "C:\vb6\com_events\TLBEvents Test.vbg"
'sLst(5) = "C:\vb6\CONTROLES-STYLE-XP\LANCEZ MOI !.vbg"
'sLst(6) = "C:\vb6\gdiplus-vb6-code-8-trunk\Group1.vbg"
'sLst(7) = "C:\vb6\HardcoreVB\HardCore3\AllAbout.vbg"
'sLst(8) = "C:\vb6\HardcoreVB\HardCore3\BitBlast.vbg"
'sLst(9) = "C:\vb6\HardcoreVB\HardCore3\Browse.vbg"
'
'ucShellBrowse1.CreateCustomFolder "Search results", sLst, -1
'ucShellBrowse1.EnableShellMenu = True

End Sub

Private Sub Command5_Click()
frmSTA.Show , Me
End Sub

Private Sub Command6_Click()
If ucShellBrowse1.ControlType = SBCTL_DirAndFiles Then
    ucShellBrowse1.ControlType = SBCTL_FilesOnly
Else
    ucShellBrowse1.ControlType = SBCTL_DirAndFiles
End If
End Sub

Private Sub Command7_Click()
If ucShellBrowse1.FilesOnly = True Then
    ucShellBrowse1.FilesOnly = False
    ucShellBrowse1.RefreshView
Else
    ucShellBrowse1.FilesOnly = True
    ucShellBrowse1.RefreshView
End If

End Sub

Private Sub Form_Load()
InitFooterHIML
'bChanging = True
'ucShellTree1.OpenToPath App.Path, False '<--- Will not work until after load
If ucShellTree1.CustomRoot <> "" Then
    ucShellTree1.InitialPath = ucShellTree1.CustomRoot
Else
    ucShellTree1.InitialPath = App.Path 'The 1st dir change from ucShellBrowse is during load, so it won't work to set this either
End If
ucShellBrowse1.ShellTreeStatus = True
ucShellBrowse1.ListViewOffsetX = ucShellTree1.Width
'bChanging = False
End Sub

Private Sub Form_Resize()
On Error GoTo e0
If Me.Width > 400 Then
    ucShellBrowse1.Height = Me.ScaleHeight - (30 * ucShellBrowse1.DPIScaleY) 'NOTE: the demo project manifest is set to dpiAware=True
    If ucShellTree1.Visible = True Then
        ucShellBrowse1.Width = Me.ScaleWidth - pbPreviewSizer.Left - 6 + pbPreviewSizer.Left
        ucShellTree1.Width = pbPreviewSizer.Left - 2
        ucShellBrowse1.ListViewOffsetX = ucShellTree1.Width
        ucShellBrowse1.DetailsPaneOffsetX = ucShellTree1.Width
        If ucShellBrowse1.ControlType = SBCTL_FilesOnly Then
            ucShellTree1.Top = ucShellBrowse1.Top + (6 * ucShellBrowse1.DPIScaleY)
            '                       v v Subtracting the detail pane height was removed
            ucShellTree1.Height = (ucShellBrowse1.Height) - ucShellBrowse1.StatusBarHeight - (10 * ucShellBrowse1.DPIScaleY)
            pbPreviewSizer.Height = ucShellTree1.Height
        Else
            ucShellTree1.Top = ucShellBrowse1.Top + (34 * ucShellBrowse1.DPIScaleY)
            ucShellTree1.Height = (ucShellBrowse1.Height) - ucShellBrowse1.StatusBarHeight - (40 * ucShellBrowse1.DPIScaleY)
            pbPreviewSizer.Height = ucShellTree1.Height
        End If
        pbPreviewSizer.Top = ucShellTree1.Top
    Else
        ucShellBrowse1.Width = Me.ScaleWidth - 6
    End If
End If
Exit Sub
e0:
DebugAppend Err.Description
Resume Next
End Sub

Private Sub Form_Terminate()
ImageList_Destroy hIML_Footer
End Sub

Private Sub ucShellBrowse1_DebugMessage(sMsg As String, nLevel As Integer)
'Note: The control now has built-in log-to-file.

End Sub

Private Sub InitFooterHIML()
Dim hIco As Long
Dim lPos As Long
Dim hr As Long
On Error GoTo e0
 
hIML_Footer = ImageList_Create(16, 16, ILC_COLOR32, 3, 1)
ExtractIconEx StrPtr(App.Path & "\zip1632.ico"), 0, ByVal 0&, hIco, 1

Call ImageList_AddIcon(hIML_Footer, hIco)
DestroyIcon hIco
ExtractIconEx StrPtr(App.Path & "\search1632.ico"), 0, ByVal 0&, hIco, 1

Call ImageList_AddIcon(hIML_Footer, hIco)
DestroyIcon hIco
ExtractIconEx StrPtr(App.Path & "\jm1632.ico"), 0, ByVal 0&, hIco, 1

Call ImageList_AddIcon(hIML_Footer, hIco)
DestroyIcon hIco
 

On Error GoTo 0
Exit Sub

e0:
DebugAppend "InitFooterHIML.Error->" & Err.Description & " (" & Err.Number & ")"

End Sub
 
Private Sub ucShellBrowse1_DetailPaneHeightChanged()
'Form_Resize <--- YIKES!! Can't do that: Out of stack space in IDE, flat out crash compiled??
'If ucShellBrowse1.ControlType = SBCTL_FilesOnly Then
'    ucShellTree1.Height = (ucShellBrowse1.Height - ucShellBrowse1.DetailsPaneHeight) - ucShellBrowse1.StatusBarHeight - (10 * ucShellBrowse1.DPIScaleY)
'Else
'    ucShellTree1.Height = (ucShellBrowse1.Height - ucShellBrowse1.DetailsPaneHeight) - ucShellBrowse1.StatusBarHeight - (40 * ucShellBrowse1.DPIScaleY)
'End If
'pbPreviewSizer.Height = ucShellTree1.Height

End Sub

Private Sub ucShellBrowse1_DirectoryChanged(ByVal sFullPath As String, siItem As oleexp.IShellItem, pidlFQ As Long)
DebugAppend "Browser.DirChange " & sFullPath
'As of ucShellTree 2.6, we no longer need to worry about manually tracking dir changes:
'If bChanging = False Then
'    bChanging = True
    ucShellTree1.OpenToItem siItem, False
'    bChanging = False
'End If
End Sub

Private Sub ucShellBrowse1_FileExecute(ByVal sFile As String, siFile As oleexp.IShellItem)
DebugAppend "ucShellBrowse->FileExecute " & sFile
End Sub

Private Sub ucShellBrowse1_FooterButtonClick(ByVal idx As Long, ByVal lParam As Long)
'ucShellBrowse1.UpdateStatus "Clicked footer button, index=" & idx & ", lParam=" & lParam
DebugAppend "Form1 footer pingback id=" & lParam

End Sub
Private Function PKEY_Size() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB725F130, &H47EF, &H101A, &HA5, &HF1, &H2, &H60, &H8C, &H9E, &HEB, &HAC, 12)
PKEY_Size = pkk
End Function
Private Function PKEY_DateModified() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB725F130, &H47EF, &H101A, &HA5, &HF1, &H2, &H60, &H8C, &H9E, &HEB, &HAC, 14)
PKEY_DateModified = pkk
End Function
Private Function PKEY_FileAttributes() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB725F130, &H47EF, &H101A, &HA5, &HF1, &H2, &H60, &H8C, &H9E, &HEB, &HAC, 13)
PKEY_FileAttributes = pkk
End Function
Private Sub DEFINE_PROPERTYKEY(Name As PROPERTYKEY, l As Long, w1 As Integer, w2 As Integer, B0 As Byte, b1 As Byte, b2 As Byte, B3 As Byte, b4 As Byte, b5 As Byte, b6 As Byte, b7 As Byte, pid As Long)
  With Name.fmtid
    .Data1 = l
    .Data2 = w1
    .Data3 = w2
    .Data4(0) = B0
    .Data4(1) = b1
    .Data4(2) = b2
    .Data4(3) = B3
    .Data4(4) = b4
    .Data4(5) = b5
    .Data4(6) = b6
    .Data4(7) = b7
  End With
  Name.pid = pid
End Sub
Private Function IID_IShellItem() As UUID
Static iid As UUID
If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H43826D1E, CInt(&HE718), CInt(&H42EE), &HBC, &H55, &HA1, &HE2, &H61, &HC3, &H7B, &HFE)
IID_IShellItem = iid
End Function
Private Function IID_IPropertyStore() As UUID
'DEFINE_GUID(IID_IPropertyStore,0x886d8eeb, 0x8cf2, 0x4446, 0x8d,0x02,0xcd,0xba,0x1d,0xbd,0xcf,0x99);
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H886D8EEB, CInt(&H8CF2), CInt(&H4446), &H8D, &H2, &HCD, &HBA, &H1D, &HBD, &HCF, &H99)
  IID_IPropertyStore = iid
  
End Function
Private Function IID_IPropertyDescription() As UUID
'(IID_IPropertyDescription, 0x6f79d558, 0x3e96, 0x4549, 0xa1,0xd1, 0x7d,0x75,0xd2,0x28,0x88,0x14
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6F79D558, CInt(&H3E96), CInt(&H4549), &HA1, &HD1, &H7D, &H75, &HD2, &H28, &H88, &H14)
  IID_IPropertyDescription = iid
  
End Function
Private Function FOLDERID_ConnectionsFolder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6F0CD92B, CInt(&H2E97), CInt(&H45D1), &H88, &HFF, &HB0, &HD1, &H86, &HB8, &HDE, &HDD)
 FOLDERID_ConnectionsFolder = iid
End Function
Public Function FOLDERID_ComputerFolder() As UUID
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAC0837C, CInt(&HBBF8), CInt(&H452A), &H85, &HD, &H79, &HD0, &H8E, &H66, &H7C, &HA7)
 FOLDERID_ComputerFolder = iid
End Function
Private Sub DEFINE_UUID(Name As UUID, l As Long, w1 As Integer, w2 As Integer, B0 As Byte, b1 As Byte, b2 As Byte, B3 As Byte, b4 As Byte, b5 As Byte, b6 As Byte, b7 As Byte)
  With Name
    .Data1 = l
    .Data2 = w1
    .Data3 = w2
    .Data4(0) = B0
    .Data4(1) = b1
    .Data4(2) = b2
    .Data4(3) = B3
    .Data4(4) = b4
    .Data4(5) = b5
    .Data4(6) = b6
    .Data4(7) = b7
  End With
End Sub
Private Function GetPropertyKeyDisplayString(pps As IPropertyStore, pkProp As oleexp.PROPERTYKEY, Optional bFixChars As Boolean = True) As String
'Gets the string value of the given canonical property; e.g. System.Company, System.Rating, etc
'This would be the value displayed in Explorer if you added the column in details view
Dim lpsz As Long
Dim ppd As IPropertyDescription
PSGetPropertyDescription pkProp, IID_IPropertyDescription, ppd
If ((pps Is Nothing) = False) And ((ppd Is Nothing) = False) Then
    PSFormatPropertyValue ObjPtr(pps), ObjPtr(ppd), PDFF_DEFAULT, lpsz
    SysReAllocString VarPtr(GetPropertyKeyDisplayString), lpsz
    CoTaskMemFree lpsz
    If bFixChars Then
        GetPropertyKeyDisplayString = Replace$(GetPropertyKeyDisplayString, ChrW$(&H200E), "")
        GetPropertyKeyDisplayString = Replace$(GetPropertyKeyDisplayString, ChrW$(&H200F), "")
        GetPropertyKeyDisplayString = Replace$(GetPropertyKeyDisplayString, ChrW$(&H202A), "")
        GetPropertyKeyDisplayString = Replace$(GetPropertyKeyDisplayString, ChrW$(&H202C), "")
    End If
    Set ppd = Nothing
Else
    DebugAppend "GetPropertyKeyDisplayString.Error->PropertyStore or PropertyDescription is not set."
    
End If
End Function

Private Sub ucShellBrowse1_ShowShellTree(bShow As Boolean)
If bShow Then
    ucShellTree1.Visible = True
'    ucShellBrowse1.Left = ucShellTree1.Left + ucShellTree1.Width + 4
    pbPreviewSizer.Visible = True
    pbPreviewSizer.Left = ucShellTree1.Left + ucShellTree1.Width + 2
    ucShellBrowse1.ListViewOffsetX = ucShellTree1.Width
    Form_Resize
Else
    ucShellTree1.Visible = False
    pbPreviewSizer.Visible = False
    ucShellBrowse1.Left = 2
    ucShellBrowse1.ListViewOffsetX = 2
End If
End Sub

Private Sub ucShellBrowse1_ToggleDetailsPane(ByVal bShow As Boolean)
Form_Resize
End Sub

Private Sub ucShellBrowse1_TogglePreviewPane(ByVal bShow As Boolean)
Form_Resize
End Sub

Private Sub ucShellBrowse1_ToggleStatusBar(ByVal bShow As Boolean)
Form_Resize
End Sub



Private Sub pbPreviewSizer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
bPMove = False
End Sub

Private Sub pbPreviewSizer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If bPMove Then
    If Button = 1 Then
        If mDX = 0 Then mDX = ucShellBrowse1.DPIScaleX
        ReleaseCapture
        SendMessage pbPreviewSizer.hWnd, WM_NCLBUTTONDOWN, 2, ByVal 0&
        DebugAppend "left=" & pbPreviewSizer.Left & ",sw=" & Me.ScaleWidth
        If pbPreviewSizer.Left > (Me.ScaleWidth - 38) Then
            pbPreviewSizer.Left = (Me.ScaleWidth - 38)
        End If
        If pbPreviewSizer.Left < (90 * mDX) Then pbPreviewSizer.Left = 90 * mDX
        Form_Resize
    Else
        bPMove = False
    End If
End If
End Sub

Private Sub pbPreviewSizer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
bPMove = True
SetWindowPos pbPreviewSizer.hWnd, HWND_TOP, 0&, 0&, 0&, 0&, SWP_NOSIZE Or SWP_NOMOVE
End If
End Sub


Private Sub ucShellTree1_ItemClick(sName As String, sFullPath As String, bFolder As Boolean, nButton As Long, hItem As Long)
Debug.Print "ItemClick " & sName
End Sub

Private Sub ucShellTree1_ItemSelect(sName As String, sFullPath As String, bFolder As Boolean, hItem As Long)
'If bChanging = False Then
'    bChanging = True
    If bFolder Then
        'Previously we set by path, but this won't work with Win10 virtual devices like phones and cameras
        'ucShellTree had to be extensively modified to not navigate or set the selected ishellitem by just the path either
'        ucShellBrowse1.BrowserPath = sFullPath
         ucShellBrowse1.BrowserOpenItem ucShellTree1.SelectedShellItem
    End If
'    bChanging = False
'End If
'Note: Sanity checking is now done internally, so you no longer need this kind of setup when changing browser dirs
End Sub

Private Sub ucShellTree1_StatusMessage(sMsg As String)
ucShellBrowse1.StatusText = sMsg 'We can set the Status Bar we have in the browser control to show the tree's status updates (mostly menu tips)
End Sub
