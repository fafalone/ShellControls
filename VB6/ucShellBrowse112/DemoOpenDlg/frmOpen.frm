VERSION 5.00
Begin VB.Form frmOpen 
   Caption         =   "Custom Open Dialog w/ ucShellBrowse and ucShellTree"
   ClientHeight    =   7350
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10365
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
   ScaleHeight     =   490
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   691
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   1110
      Left            =   7575
      TabIndex        =   11
      Top             =   5160
      Width           =   2655
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   150
         Width           =   2520
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Open"
         Default         =   -1  'True
         Height          =   435
         Left            =   45
         TabIndex        =   13
         Top             =   555
         Width           =   1185
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   435
         Left            =   1275
         TabIndex        =   12
         Top             =   555
         Width           =   1335
      End
   End
   Begin VB.PictureBox pbPreviewSizer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   3555
      Left            =   2190
      MousePointer    =   9  'Size W E
      ScaleHeight     =   237
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   2
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   600
      Width           =   30
   End
   Begin VB.Frame Frame1 
      Caption         =   "Preview"
      Height          =   1590
      Left            =   60
      TabIndex        =   5
      Top             =   5580
      Width           =   4080
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         Height          =   1320
         Left            =   60
         ScaleHeight     =   84
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   84
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   210
         Width           =   1320
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date taken"
         Height          =   195
         Left            =   1440
         TabIndex        =   9
         Top             =   750
         Width           =   2265
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Colors"
         Height          =   195
         Left            =   1440
         TabIndex        =   8
         Top             =   480
         Width           =   1785
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dimensions"
         Height          =   195
         Left            =   1440
         TabIndex        =   7
         Top             =   225
         Width           =   1500
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2835
      TabIndex        =   2
      Top             =   5310
      Width           =   4695
   End
   Begin ShellOpenDemo.ucShellTree ucShellTree1 
      Height          =   3690
      Left            =   30
      TabIndex        =   1
      Top             =   540
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   6509
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PlayNavigationSound=   0   'False
   End
   Begin ShellOpenDemo.ucShellBrowse ucShellBrowse1 
      Height          =   5115
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   9022
      HighPerformanceMode=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListViewOffsetX =   150
      ForeColor       =   -2147483635
      DetailsPaneBackColor=   -2147483645
      DetailsPaneFileNameColor=   -2147483640
      DetailsPaneForeColor=   -2147483641
      ForeColorSubitem=   16744576
      GroupSubsetLinkText=   ""
      ItemFilter      =   "*.jpg"
      ShowStatusBar   =   0   'False
      ThumbnailSize   =   144
      ColumnPreload   =   -1  'True
      EnableStatusBar =   0   'False
      EnableBookmarks =   0   'False
      SearchBox       =   -1  'True
      ShellTreeInLayout=   -1  'True
      SimpleSelect    =   0   'False
      AlignTop        =   0   'False
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This idea here is you can make much more extensive and precise customizations than even the IFileDialog control can offer."
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   5820
      TabIndex        =   4
      Top             =   6270
      Width           =   4035
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File name:"
      Height          =   195
      Left            =   2040
      TabIndex        =   3
      Top             =   5325
      Width           =   825
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private arSelPaths() As String

Private Type BITMAP
    BMType As Long
    BMWidth As Long
    BMHeight As Long
    BMWidthBytes As Long
    BMPlanes As Integer
    BMBitsPixel As Integer
    BMBits As Long
End Type
Private Declare Function PathIsDirectoryW Lib "shlwapi" (ByVal lpszPath As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal Flags As IL_CreateFlags, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_Add Lib "comctl32.dll" (ByVal himl As Long, ByVal hbmImage As Long, ByVal hBMMask As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal fStyle As IL_DrawStyle) As Boolean
Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Boolean
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

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2

Private Declare Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long) As Long
Private Declare Function PSFormatPropertyValue Lib "propsys.dll" (ByVal pps As Long, ByVal ppd As Long, ByVal pdff As PROPDESC_FORMAT_FLAGS, ppszDisplay As Long) As Long
Private Declare Function PSGetPropertyDescription Lib "propsys.dll" (PropKey As oleexp.PROPERTYKEY, riid As UUID, ppv As Any) As Long

Private Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long

Private bPMove As Boolean
Private mDX As Single
Private bChanging As Boolean
Private bStartup As Boolean

Private Sub Combo2_Click()
Select Case Combo2.ListIndex
     Case 0
        Debug.Print "Set filter to JPG"
        ucShellBrowse1.ItemFilter = "*.jpg"
     Case 1
        Debug.Print "Set filter to PNG"
        ucShellBrowse1.ItemFilter = "*.png"
     Case 2
        Debug.Print "Clear filter"
        ucShellBrowse1.ItemFilter = "*.*"
End Select
End Sub

Private Sub Command1_Click()
Dim sText As String
sText = Combo1.Text
MsgBox "The files you selected (folders were ignored):" & vbCrLf & Join(arSelPaths, vbCrLf)
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
bStartup = True
'bChanging no longer needed as of ucShellTree 2.6
'bChanging = True
ucShellBrowse1.DetailsPaneNoResizing = True
ucShellBrowse1.ShellTreeStatus = True
ucShellTree1.OpenToPath ucShellBrowse1.BrowserPath, False
Combo2.AddItem "JPEG Files"
Combo2.AddItem "PNG Files"
Combo2.AddItem "All Files"
Combo2.ListIndex = 0
'ucShellTree1.SelectNone
bChanging = False
bStartup = False
End Sub

Private Sub Form_Resize()
On Error GoTo e0
If Me.Width > 7110 Then
    ucShellBrowse1.Height = Combo1.Top - (14 * ucShellBrowse1.DPIScaleY) 'Me.ScaleHeight - NOTE: the demo project manifest is set to dpiAware=True
    If ucShellTree1.Visible = True Then
        ucShellBrowse1.Width = Me.ScaleWidth - pbPreviewSizer.Left - 6 + pbPreviewSizer.Left
        ucShellTree1.Width = pbPreviewSizer.Left - 2
        ucShellBrowse1.ListViewOffsetX = ucShellTree1.Width
        If ucShellBrowse1.ControlType = SBCTL_FilesOnly Then
            ucShellTree1.Top = ucShellBrowse1.Top + (6 * ucShellBrowse1.DPIScaleY)
            ucShellTree1.Height = (ucShellBrowse1.Height - ucShellBrowse1.DetailsPaneHeight) - ucShellBrowse1.StatusBarHeight - (10 * ucShellBrowse1.DPIScaleY)
            pbPreviewSizer.Height = ucShellTree1.Height
        Else
            ucShellTree1.Top = ucShellBrowse1.Top + ucShellBrowse1.ControlBarHeight
            ucShellTree1.Height = (ucShellBrowse1.Height - ucShellBrowse1.DetailsPaneHeight) - ucShellBrowse1.StatusBarHeight - ucShellBrowse1.ControlBarHeight - 4
            pbPreviewSizer.Height = ucShellTree1.Height
        End If
        pbPreviewSizer.Top = ucShellTree1.Top
    Else
        ucShellBrowse1.Width = Me.ScaleWidth - 6
    End If
    Frame2.Left = (Me.Width / Screen.TwipsPerPixelX) - Frame2.Width - 20
    Combo1.Width = (Me.Width / Screen.TwipsPerPixelX) - 387
Else
    Me.Width = 7110
End If
Exit Sub
e0:
Debug.Print Err.Description
Resume Next

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
SetWindowPos pbPreviewSizer.hWnd, 0&, 0&, 0&, 0&, 0&, SWP_NOSIZE Or SWP_NOMOVE
End If
End Sub

Private Sub ucShellBrowse1_ControlTypeChange(ByVal nOldType As SB_CTL_TYPE, ByVal nNewType As SB_CTL_TYPE)
Form_Resize
End Sub

Private Sub ucShellBrowse1_DetailPaneHeightChanged()
If ucShellBrowse1.ControlType = SBCTL_FilesOnly Then
    ucShellTree1.Height = (ucShellBrowse1.Height - ucShellBrowse1.DetailsPaneHeight) - ucShellBrowse1.StatusBarHeight - (10 * ucShellBrowse1.DPIScaleY)
Else
    ucShellTree1.Height = (ucShellBrowse1.Height - ucShellBrowse1.DetailsPaneHeight) - ucShellBrowse1.StatusBarHeight - (34 * ucShellBrowse1.DPIScaleY)
End If
pbPreviewSizer.Height = ucShellTree1.Height
End Sub

Private Sub ucShellBrowse1_DirectoryChanged(ByVal sFullPath As String, siItem As oleexp.IShellItem, pidlFQ As Long)
'bChanging no longer needed as of ucShellTree 2.6
'If bChanging = False Then
'    bChanging = True
    ucShellTree1.OpenToItem siItem, False
'    bChanging = False
'End If

End Sub

Private Function JnPaths(ar() As String) As String
'Outputs a quoted list, e.g. "file1.jpg" "file2.jpg" "file3.jpg"
Dim i As Long
Dim sOut As String
Dim sTmp As String

If UBound(ar) = 0 Then
    JnPaths = ar(0)
    Exit Function
End If

For i = 0 To UBound(ar)
    sTmp = Chr$(34) & ar(i) & Chr$(34) & " "
    sOut = sOut & sTmp
Next i
JnPaths = RTrim$(sOut)
End Function
Private Function JnPathsV(ar As Variant) As String
'Outputs a quoted list, e.g. "file1.jpg" "file2.jpg" "file3.jpg"
Dim i As Long
Dim sOut As String
Dim sTmp As String

If UBound(ar) = 0 Then
    JnPathsV = ar(0)
    Exit Function
End If

For i = 0 To UBound(ar)
    sTmp = Chr$(34) & ar(i) & Chr$(34) & " "
    sOut = sOut & sTmp
Next i
JnPathsV = RTrim$(sOut)
End Function

Private Sub ucShellBrowse1_SelectionChanged(arFullPaths() As String, sFocusedItem As String, siFocused As oleexp.IShellItem)
Dim hBmp As Long
Debug.Print "UCSB_SC " & arFullPaths(0)
arSelPaths = ucShellBrowse1.FilesSelectedFull
If (siFocused Is Nothing) = False Then
    Dim lpName As Long
    Dim sName As String
    siFocused.GetDisplayName SIGDN_NORMALDISPLAY, lpName
    sName = LPWSTRtoStr(lpName)
    Dim aNm() As String
    aNm = ucShellBrowse1.FilesSelected
    If UBound(aNm) > 0 Then
        Combo1.Text = JnPaths(aNm)
    Else
        Combo1.Text = aNm(0)
    End If
    If (LCase$(Right$(arFullPaths(0), 4)) = ".jpg") Or (LCase$(Right$(arFullPaths(0), 4)) = ".png") Then
        hBmp = GetFileThumbnail(siFocused, "", 0&, Picture1.ScaleWidth, Picture1.ScaleHeight, SIIGBF_THUMBNAILONLY)
        If hBmp Then
            Picture1.Cls
            hBitmapToPictureBox Picture1, hBmp
            Picture1.Refresh
        End If
        Dim pps As IPropertyStore
        Dim si2 As IShellItem2
        Set si2 = siFocused
        si2.GetPropertyStore GPS_DEFAULT Or GPS_BESTEFFORT Or GPS_OPENSLOWITEM, IID_IPropertyStore, pps
        If (pps Is Nothing) = False Then
            Label3.Caption = "Dimensions: " & GetPropertyKeyDisplayString(pps, PKEY_Image_Dimensions)
            Label4.Caption = "Bit depth: " & GetPropertyKeyDisplayString(pps, PKEY_Image_BitDepth)
            Label5.Caption = "Date Taken: " & GetPropertyKeyDisplayString(pps, PKEY_Photo_DateTaken)
        End If
    Else
        Picture1.Cls
        Picture1.Refresh
    End If
Else
    Combo1.Text = ""
    
End If
End Sub

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


Private Sub DEFINE_UUID(Name As oleexp.UUID, L As Long, w1 As Integer, w2 As Integer, B0 As Byte, b1 As Byte, b2 As Byte, B3 As Byte, b4 As Byte, b5 As Byte, b6 As Byte, b7 As Byte)
  With Name
    .Data1 = L
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
Private Function IID_IShellItemImageFactory() As oleexp.UUID
'{BCC18B79-BA16-442F-80C4-8A59C30C463B}
Static iid As oleexp.UUID
If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBCC18B79, CInt(&HBA16), CInt(&H442F), &H80, &HC4, &H8A, &H59, &HC3, &HC, &H46, &H3B)
IID_IShellItemImageFactory = iid
End Function
Private Function IID_IPropertyStore() As UUID
'DEFINE_GUID(IID_IPropertyStore,0x886d8eeb, 0x8cf2, 0x4446, 0x8d,0x02,0xcd,0xba,0x1d,0xbd,0xcf,0x99);
Static iid As UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H886D8EEB, CInt(&H8CF2), CInt(&H4446), &H8D, &H2, &HCD, &HBA, &H1D, &HBD, &HCF, &H99)
  IID_IPropertyStore = iid
  
End Function
Private Function IID_IPropertyDescription() As oleexp.UUID
'(IID_IPropertyDescription, 0x6f79d558, 0x3e96, 0x4549, 0xa1,0xd1, 0x7d,0x75,0xd2,0x28,0x88,0x14
Static iid As oleexp.UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6F79D558, CInt(&H3E96), CInt(&H4549), &HA1, &HD1, &H7D, &H75, &HD2, &H28, &H88, &H14)
  IID_IPropertyDescription = iid
End Function


Private Function PKEY_Image_Dimensions() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6444048F, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 13)
PKEY_Image_Dimensions = pkk
End Function
Private Function PKEY_Image_BitDepth() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6444048F, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 7)
PKEY_Image_BitDepth = pkk
End Function
Private Function PKEY_Photo_DateTaken() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H14B81DA1, &H135, &H4D31, &H96, &HD9, &H6C, &HBF, &HC9, &H67, &H1A, &H99, 36867)
PKEY_Photo_DateTaken = pkk
End Function
Private Sub DEFINE_PROPERTYKEY(Name As PROPERTYKEY, L As Long, w1 As Integer, w2 As Integer, B0 As Byte, b1 As Byte, b2 As Byte, B3 As Byte, b4 As Byte, b5 As Byte, b6 As Byte, b7 As Byte, pid As Long)
  With Name.fmtid
    .Data1 = L
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

Private Function GetFileThumbnail(siItem As oleexp.IShellItem, sFile As String, pidlFQ As Long, cx As Long, cy As Long, Optional dwFlags As SIIGBF = SIIGBF_THUMBNAILONLY Or SIIGBF_RESIZETOFIT) As Long
Dim isiif As oleexp.IShellItemImageFactory
On Error GoTo e0

If (siItem Is Nothing) Then
    If pidlFQ Then
        Call oleexp.SHCreateItemFromIDList(pidlFQ, IID_IShellItemImageFactory, isiif)
    Else
        oleexp.SHCreateItemFromParsingName StrPtr(sFile), Nothing, IID_IShellItemImageFactory, isiif
    End If
Else
    Set isiif = siItem
End If

isiif.GetImage cx, cy, dwFlags, GetFileThumbnail
Set isiif = Nothing
On Error GoTo 0
Exit Function

e0:
Debug.Print "GetFileThumbnail.Error->" & Err.Description & " (" & Err.Number & ")"

End Function

Private Sub hBitmapToPictureBox(picturebox As Object, hBitmap As Long, Optional X As Long = 0&, Optional Y As Long = 0&)
Debug.Print "hBitmapToPictureBox"
'This or similar is always given as the example on how to do this
'But it results in transparency being lost
'So the below method seems a little ackward, but it works. It can
'be done without the ImageList trick, but it's much more code
Dim himlBmp As Long
Dim tBMP As BITMAP
Dim cx As Long, cy As Long
Call GetObject(hBitmap, LenB(tBMP), tBMP)
cx = tBMP.BMWidth
cy = tBMP.BMHeight
If cx = 0 Then
    Debug.Print "no width"
    Exit Sub
End If
himlBmp = ImageList_Create(cx, cy, ILC_COLOR32, 1, 1)
ImageList_Add himlBmp, hBitmap, 0&
If (X = 0) And (Y = 0) Then
    'not manually specified, so center
    If cy < picturebox.ScaleHeight Then
        Y = ((picturebox.ScaleHeight - cy) / 2) '* Screen.TwipsPerPixelY
    End If
    If cx < picturebox.ScaleWidth Then
        X = ((picturebox.ScaleWidth - cx) / 2) '* Screen.TwipsPerPixelX
    End If
'    Debug.Print "frame=" & fraPreview.Width & "," & fraPreview.Height & " pc4=" & pbPreviewPane.ScaleWidth & "(" & pbPreviewPane.Width / Screen.TwipsPerPixelX & ")," & pbPreviewPane.ScaleHeight & " cx=" & CX & " cy=" & CY & " x=" & X & " y=" & Y
End If
ImageList_Draw himlBmp, 0, picturebox.hDC, X, Y, ILD_NORMAL

ImageList_Destroy himlBmp
End Sub
Private Function LPWSTRtoStr(lPtr As Long, Optional ByVal fFree As Boolean = True) As String
SysReAllocString VarPtr(LPWSTRtoStr), lPtr
If fFree Then
    Call CoTaskMemFree(lPtr)
End If
End Function

Private Sub ucShellTree1_ItemSelect(sName As String, sFullPath As String, bFolder As Boolean, hItem As Long)
'If bStartup = False Then
    If bFolder Then
        'Previously we set by path, but this won't work with Win10 virtual devices like phones and cameras
        'ucShellTree had to be extensively modified to not navigate or set the selected ishellitem by just the path either
'        ucShellBrowse1.BrowserPath = sFullPath
         ucShellBrowse1.BrowserOpenItem ucShellTree1.SelectedShellItem
    End If
'End If
End Sub
