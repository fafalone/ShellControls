VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000005&
   Caption         =   "ShellBrowseDemo"
   ClientHeight    =   8880
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18045
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   18045
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command14 
      Caption         =   "DirOnlyForm"
      Height          =   360
      Left            =   16665
      TabIndex        =   14
      Top             =   30
      Width           =   1290
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Col Preload On"
      Height          =   360
      Left            =   15255
      TabIndex        =   13
      Top             =   30
      Width           =   1410
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Custom Col"
      Height          =   360
      Left            =   14220
      TabIndex        =   12
      Top             =   30
      Width           =   1020
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Open FTP"
      Height          =   360
      Left            =   13215
      TabIndex        =   11
      Top             =   30
      Width           =   990
   End
   Begin VB.CommandButton Command10 
      Caption         =   "CustomColor"
      Height          =   360
      Left            =   12000
      TabIndex        =   10
      Top             =   30
      Width           =   1185
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Print Path"
      Height          =   360
      Left            =   10980
      TabIndex        =   9
      Top             =   30
      Width           =   990
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Toggle Search"
      Height          =   360
      Left            =   9735
      TabIndex        =   8
      Top             =   30
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Custom List (EDIT)"
      Height          =   360
      Left            =   8235
      TabIndex        =   7
      Top             =   30
      Width           =   1470
   End
   Begin VB.CommandButton Command6 
      Caption         =   "SetFocusOnFiles"
      Height          =   360
      Left            =   6780
      TabIndex        =   6
      Top             =   30
      Width           =   1425
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Toggle FilterBar"
      Height          =   360
      Left            =   5287
      TabIndex        =   5
      Top             =   30
      Width           =   1485
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Highlight Col"
      Height          =   360
      Left            =   4089
      TabIndex        =   4
      Top             =   30
      Width           =   1170
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Output Dir List"
      Height          =   360
      Left            =   2696
      TabIndex        =   3
      Top             =   30
      Width           =   1365
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Footer Demo 2"
      Height          =   360
      Left            =   1393
      TabIndex        =   2
      Top             =   30
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Footer Demo 1"
      Height          =   360
      Left            =   90
      TabIndex        =   1
      Top             =   30
      Width           =   1275
   End
   Begin ShellBrowseDemo.ucShellBrowse ucShellBrowse1 
      Height          =   8205
      Left            =   90
      TabIndex        =   0
      Top             =   495
      Width           =   13710
      _ExtentX        =   24183
      _ExtentY        =   14473
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
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      ControlBackColor=   -2147483643
      DetailsPaneBackColor=   -2147483645
      ForeColorSubitem=   -2147483630
      ItemFilter      =   ""
      FullRowSelect   =   -1  'True
      SnapToGrid      =   -1  'True
      StatusTextStart =   "Welcome to the ucShellBrowse Demo!"
      ThumbnailSize   =   128
      BookmarkButton  =   -1  'True
      HeaderHotTracking=   -1  'True
      SearchBox       =   -1  'True
      SearchPopupInMenu=   -1  'True
      MaxHistoryDisplayed=   10
      RestrictViewModes=   "7"
      SelectColumnOnSort=   -1  'True
      CustomIconsEnabled=   -1  'True
      CustomIconsSize =   32
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Note: mDemo was removed. The CC init in Sub Main doesn't actually have any effect on Win Vista+, which this project requires anyway
'Note: cLog and the class file has been removed from this demo as log-to-file is now built into the control.
'The DemoEx sample still has it for logging Form1 messages.
Private nColID1 As Long
Private GOCANES As Boolean
Private Declare Function IsUserAdmin Lib "Shell32" Alias "#680" () As Long
Private Const dbg_IncludeDate As Boolean = True
Private hIML_Footer As Long
Private Const CB_GETDROPPEDWIDTH = &H15F
Private lgtc1 As Long, lgtc2 As Long
Private hIcoF As Long, hIcoD As Long

Private Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExW" (ByVal lpszFile As Long, _
                                                                                ByVal nIconIndex As Long, _
                                                                                phiconLarge As Long, _
                                                                                phiconSmall As Long, _
                                                                                ByVal nIcons As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Declare Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long) As Long
Private Declare Function PSFormatPropertyValue Lib "propsys.dll" (ByVal pps As Long, ByVal ppd As Long, ByVal pdff As PROPDESC_FORMAT_FLAGS, ppszDisplay As Long) As Long
Private Declare Function PSGetPropertyDescription Lib "propsys.dll" (PropKey As oleexp.PROPERTYKEY, riid As UUID, ppv As Any) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameW" (ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

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
  ILD_PRESERVEALPHA = &H1000
  ILD_SCALE = &H2000
  ILD_DPISCALE = &H4000
  ILD_ASYNC = &H8000
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

'#If (WIN32_IE >= &H300) Then
Private Type IMAGELISTDRAWPARAMS
  cbSize As Long
  himl As Long
  i As Long
  hdcDst As Long
  x As Long
  y As Long
  cx As Long
  CY As Long
  xBitmap As Long   ' x offest from the upperleft of bitmap
  yBitmap As Long   ' y offset from the upperleft of bitmap
  rgbBk As Long
  rgbFg As Long
  fStyle As Long
  dwRop As Long
End Type
'#End If
 
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
Private Declare Function SHGetPropertyStoreFromParsingName Lib "Shell32" (ByVal pszPath As Long, pbc As IBindCtx, ByVal Flags As GETPROPERTYSTOREFLAGS, riid As UUID, ppv As Any) As Long
Private Declare Function ImageList_Add Lib "comctl32.dll" (ByVal himl As Long, ByVal hbmImage As Long, ByVal hBMMask As Long) As Long
Private Declare Function ImageList_AddMasked Lib "comctl32.dll" (ByVal himl As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long
Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal CY As Long, ByVal Flags As IL_CreateFlags, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Boolean
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal fStyle As IMAGELISTDRAWFLAGS) As Boolean
Private Declare Function ImageList_GetIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal Flags As IMAGELISTDRAWFLAGS) As Long
Private Declare Function ImageList_LoadImage Lib "comctl32.dll" Alias "ImageList_LoadImageW" (ByVal hi As Long, ByVal lpbmp As Long, ByVal cx As Long, ByVal cGrow As Long, ByVal crMask As Long, ByVal uType As ImageTypes, ByVal uFlags As LoadResourceFlags) As Long
Private Declare Function ImageList_Remove Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long) As Boolean
Private Declare Function ImageList_ReplaceIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hIcon As Long) As Long
Private Declare Function ImageList_SetBkColor Lib "comctl32.dll" (ByVal himl As Long, ByVal clrBk As Long) As Long
Private Declare Function ImageList_SetOverlayImage Lib "comctl32.dll" (ByVal himl As Long, ByVal iImage As Long, ByVal iOverlay As Long) As Boolean
Private Declare Function ImageList_CoCreateInstance Lib "comctl32.dll " (refclsid As UUID, ByVal pUnkOuter As Long, riid As UUID, ppv As Any) As Long
Private Declare Function HIMAGELIST_QueryInterface Lib "comctl32.dll" (ByVal himl As Long, riid As UUID, ppv As Any) As Long

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
Private Function LPWSTRtoStr(lPtr As Long, Optional ByVal fFree As Boolean = True) As String
SysReAllocString VarPtr(LPWSTRtoStr), lPtr
If fFree Then
    Call CoTaskMemFree(lPtr)
End If
End Function
Private Sub DebugAppend(ByVal sMsg As String, Optional ilvl As Integer = 0)
If dbg_IncludeDate Then sMsg = "[" & Format$(Now, "yyyy-mm-dd Hh:Mm:Ss") & "] " & "[Form1] " & sMsg
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

Private Sub Command10_Click()
'Enables the response to the CustomColor event and redraws all the items to apply it immediately
If GOCANES = False Then
    GOCANES = True
Else
    GOCANES = False
End If
ucShellBrowse1.RedrawList
End Sub

Private Sub Command11_Click()
ucShellBrowse1.BrowserPath = "ftp://speedtest.tele2.net/"
End Sub

Private Sub Command12_Click()
nColID1 = ucShellBrowse1.AddCustomColumn("Ok?", PDDT_STRING, 1, 40, True)
End Sub
 
Private Sub Command13_Click()
If ucShellBrowse1.ColumnPreload = True Then
    ucShellBrowse1.ColumnPreload = False
    Command13.Caption = "Col Preload On"
Else
    ucShellBrowse1.ColumnPreload = True
    Command13.Caption = "Col Preload Off"
End If
End Sub

Private Sub Command14_Click()
frmDirOnly.Show , Me
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
ucShellBrowse1.HighlightColumn 2, vbGreen
End Sub

Private Sub Command5_Click()
If ucShellBrowse1.FilterBar = True Then
    ucShellBrowse1.FilterBar = False
Else
    ucShellBrowse1.FilterBar = True
End If
End Sub

Private Sub Command6_Click()
'ucShellBrowse1.Enabled = False
'DoEvents
'ucShellBrowse1.BrowserPath = "C:\temp2"
'ucShellBrowse1.SetFocusOnDropdown
 
   ucShellBrowse1.SetFocusOnFiles
End Sub

Private Sub Command7_Click()
'To use this custom file list demo you must substitute your own list of files!
Dim sAr() As String
ReDim sAr(4)
sAr(0) = "C:\temp2\add5.png"
sAr(1) = "C:\download\111.jpg"
sAr(2) = "C:\download\bill2.docx"
sAr(3) = "C:\AMD\gss.txt"
sAr(4) = "C:\vb6\tmp.RES"

'Customize Columns:
Dim apk(1) As oleexp.PROPERTYKEY
apk(0) = PKEY_Size
apk(1) = PKEY_FileAttributes
ucShellBrowse1.CreateCustomFolder "TestDir", sAr, , VarPtr(apk(0)), UBound(apk) + 1


End Sub

Private Sub Command8_Click()
If ucShellBrowse1.SearchBox = True Then
    ucShellBrowse1.SearchBox = False
Else
    ucShellBrowse1.SearchBox = True
End If
End Sub

'Dim hbm As Long
'
'Dim psi As IShellItem
'Set psi = ucShellBrowse1.SelectedItem
'
'Dim isiif As IShellItemImageFactory
'Set isiif = psi
'isiif.GetImage 96, 96, SIIGBF_THUMBNAILONLY Or SIIGBF_RESIZETOFIT, hbm 'substitute whatever width/height for 96
'Set isiif = Nothing
'
''Now you have an HBITMAP. You can set it to your own external PictureBox with the
''HBitmapToPictureBox function from this post, or set it as the preview in the ucShellBrowse
''Preview Box (View->Preview Pane to turn it on at runtime, or .PreviewPane property to set
''via code or in the Design Time properties). Since the preview handler might load a video
''player with the first few seconds instead of a static image.
'
'ucShellBrowse1.SetPreviewPictureWithHBITMAP hbm

Private Sub Command9_Click()
Debug.Print ucShellBrowse1.BrowserPath
End Sub

  
Private Sub Form_Load()
InitFooterHIML
ExtractIconEx StrPtr(App.Path & "\ICO_UP32B.ico"), 0, hIcoD, ByVal 0&, 1
ExtractIconEx StrPtr(App.Path & "\fwd32.ico"), 0, hIcoF, ByVal 0&, 1

End Sub

 
Private Sub Form_Resize()
On Error Resume Next
'If Me.Width > 400 Then
    ucShellBrowse1.Width = Me.Width - 400 '- Text1.Width
    ucShellBrowse1.Height = Me.Height - 1200
'End If
End Sub

Private Sub Form_Terminate()
ImageList_Destroy hIML_Footer
End Sub


Private Sub ucShellBrowse1_CustomColor(ByVal itemIndex As Long, ByVal ItemName As String, ByVal SubItemIndex As Long, ByVal SubItemProp As String, ByVal dwStateCD As Long, ByVal dwStateLV As Long, rgbFore As Long, rgbBack As Long)
If GOCANES Then
    If ((itemIndex Mod 2) = 0) Then
        If (SubItemIndex Mod 2) = 0 Then
            rgbFore = RGB(0, 100, 0)
            rgbBack = RGB(255, 165, 0)
        Else
            rgbBack = RGB(0, 100, 0)
            rgbFore = RGB(255, 165, 0)
        End If
    Else
        If (SubItemIndex Mod 2) = 0 Then
            rgbFore = RGB(255, 165, 0)
            rgbBack = RGB(0, 100, 0)
        Else
            rgbFore = RGB(0, 100, 0)
            rgbBack = RGB(255, 165, 0)
        End If
    End If
End If
End Sub

Private Sub ucShellBrowse1_CustomColumnQueryData(lColID As Long, pItem As oleexp.IShellItem, pidlFQ As Long, sItemName As String, sPath As String, lPos As Long, out_ColText As String, out_nImage As Long)
If lColID = nColID1 Then
    If lPos Mod 2 = 0 Then
        If ucShellBrowse1.ItemIsFolder(lPos) = False Then
            out_ColText = "N"
            out_nImage = 5
        End If
    Else
        If ucShellBrowse1.ItemIsFolder(lPos) = False Then
            out_ColText = "Y"
            out_nImage = 6
        End If
    End If
End If
End Sub

Private Sub ucShellBrowse1_DirectoryChanged(ByVal sFullPath As String, siItem As oleexp.IShellItem, pidlFQ As Long)
Debug.Print "DirChange " & sFullPath
End Sub

Private Sub ucShellBrowse1_DropFiles(sFiles() As String, siaFiles As oleexp.IShellItemArray, doDropped As oleexp.IDataObject, sDropParent As String, siDropParent As oleexp.IShellItem, iEffect As oleexp.DROPEFFECTS, dwKeyState As Long, ptX As Long, ptY As Long)
Debug.Print "Dropped " & Join(sFiles, "|")
End Sub

Private Sub ucShellBrowse1_FileExecute(ByVal sFile As String, siFile As oleexp.IShellItem)
Debug.Print "FileExecute " & sFile
End Sub

Private Sub ucShellBrowse1_FooterButtonClick(ByVal idx As Long, ByVal lParam As Long)
'ucShellBrowse1.UpdateStatus "Clicked footer button, index=" & idx & ", lParam=" & lParam
DebugAppend "Form1 footer pingback id=" & lParam

End Sub

Private Sub ucShellBrowse1_DebugMessage(sMsg As String, nLevel As Integer)
'Control now has its own internal logger; this event is still available if you
'want to still log that way.
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
Private Function FOLDERID_ComputerFolder() As UUID
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
    Debug.Print "GetPropertyKeyDisplayString.Error->PropertyStore or PropertyDescription is not set."
    
End If
End Function

Public Function ImageList_AddIcon(himl As Long, hIcon As Long) As Long
  ImageList_AddIcon = ImageList_ReplaceIcon(himl, -1, hIcon)
End Function

Public Function ImageList_RemoveAll(himl As Long) As Boolean
  ImageList_RemoveAll = ImageList_Remove(himl, -1)
End Function

' HINSTANCE hi - This parameter is not used and should always be zero.

Public Function ImageList_ExtractIcon(hi As Long, himl As Long, i As Long) As Long
  ImageList_ExtractIcon = ImageList_GetIcon(himl, i, 0)
  Dim x As Long
  x = ILCF_MOVE
  
End Function

Public Function ImageList_LoadBitmap(hi As Long, lpbmp As Long, cx As Long, cGrow As Long, crMask As Long) As Long
  ImageList_LoadBitmap = ImageList_LoadImage(hi, lpbmp, cx, cGrow, crMask, IMAGE_BITMAP, 0)
End Function

Private Sub ucShellBrowse1_QueryCustomIcon(ByVal sName As String, ByVal sPath As String, ByVal sFullPath As String, ByVal bIsFolder As Boolean, pidlFQItem As Long, ByVal cxy As Long, out_HBMItem As Long, out_HICOItem As Long, out_Destroy As Boolean)
If bIsFolder Then
    out_HICOItem = hIcoD
Else
    out_HICOItem = hIcoF
End If
End Sub
