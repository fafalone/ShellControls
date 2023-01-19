VERSION 5.00
Begin VB.UserControl ucShellTree 
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   690
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForwardFocus    =   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   40
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   46
   ToolboxBitmap   =   "ucShellTree.ctx":0000
End
Attribute VB_Name = "ucShellTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const mVersionStr As String = "Shell Tree Control 2.7"
''*********************************************************************************************
'
'ucShellTree.ctl
'Shell Tree Control v2.7
'
'Author: fafalone
'(c) 2018-2021
'
'For questions, comments, and bug reports, stop by the project thread:
'http://www.vbforums.com/showthread.php?862137-ucShellTree-Full-featured-Shell-Tree-UserControl
'
'----------------------------------------ABOUT-------------------------------------------
'This UserControl displays a TreeView of the Windows Shell, similar to the one
'found in Explorer. While there are other controls that are similar, and there's
'also the INamespaceTreeControl that actually hosts an Explorer TreeView, this has
'the advantage that it has modern styling and features that weren't available before
'Windows Vista, but being completely done in VB provides customizations that are
'not possible with the hosted object.
'
'-------------------------------------REQUIREMENTS---------------------------------------
'-Windows Vista or newer
'-oleexp.tlb v4.42 or higher
'
'STRONGLY RECOMMENDED: Common Controls 6.0 manifest. Partial and exclusion checkboxes will
'                      not work, visual styling is effected, and other features may not work
'                      as well.
'-------------------------------------KEY FEATURES---------------------------------------
'-Displays complete Explorer-type tree with either Desktop or Computer as the root folder.
'-Tri-state checkboxes show partial selections when check mode is enabled
'-Supports drag/drop with modern drag images and drag-over-highlighting (including
'  expanding on hover), based on my cDropTarget project. Can drop on all valid drop targets:
'  folders, zip files, programs, shortcuts to them, etc.
'-Right-click shows the standard Explorer context menu for the clicked item
'-Automatically monitors for changes (item created, deleted, renamed) and updates the
'  tree accordingly.
'-Option to show files as well
'-Can automatically expand to a given path. NOTE: OpenToPath/OpenToItem can only be used
'  after loading; to open to a path on your Form_Load or Main() routine, use .InitialPath
'-InfoTips with several lines of details depending on file type are shown as ToolTips
'-Filter option can limit the type of files displayed (or even folders)
'-Can rename in place using LabelEdit
'-Can (optionally) treat .zip/.cab files as a folder
'-Optional additional root entry for 'Favorites' that shows the Links folder.
'-Browses the Network folder too and returns paths as \\Share\etcetc
'-Complete Unicode support
'
'
'---------------------------------------CHANGELOG----------------------------------------
'
'v2.7 (Released 25 Jan 2022)
'
'-Added ShowHiddenItems/ShowSuperHidden options
'
'-Added EnableShellMenu option to control whether the right-click menu pops up.
'
'-Added SetFocusOnTree method.
'
'-Added public event for UserControl_EnterFocus and UserControl_ExitFocus (EnterFocus
' and ExitFocus, respectively).
'
'-(Bug fix) Keyboard focus never went to ucTreeView when on a form with ucShellBrowse.
'
'v2.6 (Released 03 Apr 2021)
'
'-Eliminated the need to use a variable to keep tracking of dir change operations when
' combining this control with a ucShellBrowse control; previously you'd handle a path
' change notification from the browser with
' If bChanging = False Then
'    bChanging = True
'    ucShellTree1.OpenToItem siItem, False
'    bChanging = False
' End If
' Now you no long need the bChanging variable.
'
'-Added SelectNone sub, which will clear all selected items (supports multiselect).
'
'-Added ItemSelectByShellItem event. I wanted to just include an IShellItem member in
' ItemSelect, but the backwards compatibility concerns are too great. This event will
' also include all the information of ItemSelect if you want to fully switch over.
'-Did the same for ItemClick->ItemClickByShellItem
'
'-For MultiSelect, added the MultiSelectChange event, which includes an IShellItemArray
' of selected items as well as a name list and full path list
'
'-(Bug fix) Using On Error Resume Next to handle an uninitialized array in the Terminate
'           event caused the control to freeze on compile or on run after changes if
'           you were using it as an OCX.
'
'
'v2.5 R2 (Released 15 Dec 2020)
'
'-(Bug fix) Navigation in unmapped Network locations wouldn't expand properly. This may
'           also have effected other items, as the bug involves scenarios where two
'           IShellItem.GetDisplayName calls for the same location return a different case
'           at a different time, which resulted in an infinite loop looking for something
'           that should have been there, because case differed unexpectedly.
'
'v2.5 (Released 09 Dec 2020)
'
'-Added BackColor and ForeColor (text) options.
'
'-Added some missing style options: AutoHScroll, NoIndentState, TrackSelect, and
' ShowSelAlways. The latter two are enabled by default.
'
'-Added RefreshTreeView function. There already is 'ResetTreeView', but that restores the
' tree to its state on load. RefreshTreeView resets it then expands all the folders that
' were previously expanded, which will make any updates in the folders that were visible
' as they're loaded from scratch.
'
'-To perform the refresh, a list is generated of not all folders, but just one per end
' node, otherwise there'd be potentially hundreds of useless OpenToPath calls.
' This list is available if you want to view or save it through the new GetExpansionState
' function. If you want to load a saved list, also added LoadExpansionState. You can pass
' any list of full paths in a string delimited by a | if desired.
'
'-(Bug fix) If a portable device (e.g. phone, camera) was connected while the control was
'           running, it wouldn't be added to the tree.
'
'v2.4 (Released 20 Sep 2020)
'
'-Extended Verbs for the Shell Context Menu is now an option. If it's set to False (the
' default), you will need to hold down Shift when bringing up the menu to show the
' additional items, the same way it works in Explorer.
'
'-(Bug fix) Eliminated all non-explicit types that were defined in the type library, so
'           that no further conflicts with Public versions are possible.
'
'v2.3 (Released 19 August 2020)
'
'-(Bug fix) Extended styles were being cleared improperly. Checkboxes, ExclusionChecks,
'           FadingExpandos, and MultiSelect would not turn off once enabled.
'
'v2.2 (Released 23 April 2020)
'
'-Added DropFiles event and DragStart event.
'
'-Added Multiselect option. MSDN says "Not supported, do not use." but this style appears
' to be working without issue. Be advised, this deprecated status means that it could cease
' to work in future versions of Windows without notice.
' This effects dragging items, and a new Public Sub SelectedItems has been added; these are
' the only places multiselect currently effects.
'
'-Added DisableDragDrop option if you want to disable it.
'
'
'v2.17 (Released 15 Mar 2020)
'
'-(Bug fix) The default overlays that should always stay on weren't showing (Link/Share)
'
'v2.16 (Released 05 Mar 2020)
'
'-Made extended overlays, found in programs such as TortoiseSVN and Dropbox, a default-off
' option due to the extreme performance cost (a factor of 10-100).
'
'v2.15 (Released 19 Feb 2020)
'
'-(Bug fix) If a Private Enum defined in a UserControl had the same name as one that
'           is defined in a module containing Sub Main, placing more than one of the
'           UserControl in a project caused the control to be grayed out and then an
'           app crash when any other control was initially added to the form.
'           To fix this, all API enums in this control have been prefixed with ucsb_
'           If you're modifying the code of this control, just keep that in mind.
'           No change is needed for using this control or any other code as they're
'           all Private.
'
'-(Bug fix) In ensuring a USB device is already added after OpenToItem, the GetNodeByPath
'           function did not handle USB paths properly, so additional USB device entries
'           were added to the tree.
'
'v2.14
'
'-(Bug fix) If a USB media device (phone, camera, etc-- no drive letter) was added after
'           the program started, it could not be added to the TreeView-- there was some
'           issue with an infinite loop where it kept finding the parent but not adding
'           the child. I couldn't figure out where it was going wrong, so implemented a
'           workaround where the control checks whether it's being asked to navigate to
'           one of these devices, and if it is, manually ensures the device is added.
'           After that, subfolder adding works fine, you can directly navigate to a deep
'           path and all the subfolders will be added; the issue was just adding the
'           device itself.
'
'-(Bug fix) The control uses the SHChangeNotifyRegister API without its own declare; there
'           is a declare in the typelib, and since it wasn't declared explicitly, any
'           Public version in a project module would take precedence, which sometimes
'           caused Type Mismatch errors if the declare was slightly different.
'
'v2.13
'
'-InfoTips are now cached on first load.
'
'-Now store fully qualified pidls for each item, which enables compatibility with the search
' folders generated by ucShellBrowse's new search method (ISearchFolderItemFactory), as these
' for some reason lack a relative pidl even though other direct children of the desktop do not.
' There may be other types of objects lacking a relative pidl as well; this can only increase
' the number of item types supported.
'
'-(Bug fix) In Design Mode, the control displayed "Shell Tree Control 1.0" instead of the current
'           version number. It now gets the info from mVersionStr on Line #2 of this module.
'
'v2.12 (Released 16 June 2019)
'-(Bug fix) On Windows 10, when browsing some virtual devices, like connected phones or cameras,
'           the SHCreateItemFromParsingName fails when 2 levels or deeper. Changed navigation,
'           selection generation, and other features to have full pidl records to fall back on.
'
'v2.11 (Released 15 Feb 2019 - Critical bug fix only)
'-(Bug fix) Previous fix regarding PathMatchSpecW resulted in always loading the Computer
'           folder as root.
'
'v2.1 (Released 30 Jan 2019)
'-The font for the TreeView is now a standard property. (Borrowed from Krool's
' TreeView. Thanks!)
'
'-Shell context menu tips are now passed in the StatusMessage event.
'
'-Filter now supports multiple patterns separated by semi-colon.
'
'-Replaced Border property with BorderStyle, which has several more options.
'
'-(Bug fix) If the control was on a secondary form, unloading that form did not unload
'           the control, and it stayed loaded in the background until the whole program
'           ended. Thanks to dz32 and Eduardo- for figuring out the solution.
'
'-(Bug fix) Fixed automatic navigation not expanding when a parent of a parent was
'           collapsed; only the immediate parent was checked previously.
'
'-(Bug fix) Custom Roots were checked for validity with PathMatchSpecW, which does
'           not support virtual locations. Now it's checked by going ahead and trying
'           to create the IShellItem for it, allowing roots like ::{GUID}
'v2
'
'-You can now specify a custom folder as root*. Changeable during runtime.

'-Added PathGetCheck and PathSetCheck functions. The Set function also has an option to
' expand to show the given path in the event it's not yet visible.
'
'-The OpenToPath/OpenToItem functions now have an option to just expand to but not select
' the item, primarily for the check set function but available in general; default=select.
'
'-The .InitialPath property is deprecated. The creation sequence has changed, you can
' now just use OpenPath in your Form_Load (or equiv.) event. It remains for compatibility.
'
'-Added RootHasCheckbox option to set whether or not one appears (when checkboxes are enabled).
'
'-Added ExclusionChecks option, which adds an additional checkbox state- a red x. The .ExcludedPaths
' method functions in the same way as the .CheckedPaths method to retrieve paths in this state (they
' are not counted as checked).
'
'-Added ExplorerStyle option to allow control over whether the Explorer visual style is applied.
'
'-Added HorizontalScroll option, to set whether the tree expands without needed HScroll.
'
'-Adjusted the RootHasCheckbox option so that if it's changed at runtime, the correct state is set.
'
'-Changed disable criteria such that valid file drop targets are excepted from normal disabled prop
' list. This was mainly to get the Recycle Bin enabled since it's completely browsable. Other items
' remain disabled. (You can change this in TVExpandFolder if desired).
'
'-Added Autocheck option for control over whether partial checks are shown and whether parent and
' child items are automatically changed according to a new check action.
'
'-The control will now fall back to plain checkboxes if no Common Control 6 manifest is present, but
' all autocheck functionality is unavailable (even checking all children of a checked parent).
' Previously checkboxes were not present at all without a manifest.
'
'(Bug fix) The outer edge of the control on dragover indicated a droptarget on no item but
'          seemingly valid; it now correctly shows no drop is possible.
'
'(Bug fix) Invalid (malformed/corrupt) shortcut files caused an error that wasn't handled, so
'          the item enumeration when expanding a folder stopped when it reached it.
'
'(Bug fix) When you call GetParent on your user folder, you get the desktop instead of \Users\,
'          this caused some issues with auto-navigating.
'
'(Control/code) So I thought you needed to supply your own check imagelist for partial checkboxes,
'               since the extended style wasn't working, but it turned out the problem was just that
'               you can't set the TVS_CHECKBOXES style; just TVS_EX_PARTIALCHECKBOXES. So the images
'               have been removed, which eliminates transparency issues and will keep the appearance
'               the same as the OS.
'
'-------
'* - Technical note: This can be any path string/identifier resolvable by SHCreateItemFromParsingName
'
'*********************************************************************************************

'*1
'USER OPTIONS
'The following are meant to be toggled based on your preferences:
Private Const dbg_PrintToImmediate As Boolean = False 'This control has very extensive debug information, you may not want
                                                     'to see that in your IDE.
Private Const dbg_IncludeDate As Boolean = True 'Prefix all Debug output with the date and time, [yyyy-mm-dd Hh:Mm:Ss]
Private Const dbg_RaiseEvent As Boolean = True 'Raise DebugMessage event
Private Const dbg_MinLevel As Long = 0& 'Only fire debug statement if iLvl >= this value
'-----------------------------------------------------------------------------------------------

Implements oleexp.IDropTarget

Public Event StatusMessage(sMsg As String)
Public Event ItemClick(sName As String, sFullPath As String, bFolder As Boolean, nButton As Long, hItem As Long)
Public Event ItemClickByShellItem(siItem As oleexp.IShellItem, sName As String, sFullPath As String, bFolder As Boolean, nButton As Long, hItem As Long)
Public Event ItemSelect(sName As String, sFullPath As String, bFolder As Boolean, hItem As Long)
Public Event ItemSelectByShellItem(siItem As oleexp.IShellItem, sName As String, sFullPath As String, bFolder As Boolean, hItem As Long)
Public Event MultiSelectChange(siaItems As oleexp.IShellItemArray, sNames() As String, sFullPaths() As String)
Public Event ItemCheck(sName As String, sFullPath As String, bFolder As Boolean, fCheck As Long, hItem As Long)
Public Event ItemExpand(sName As String, sFullPath As String, hItem As Long)
Public Event ItemCollapse(sName As String, sFullPath As String, hItem As Long)
Public Event TreeKeyDown(VKey As Integer)
Public Event ItemRename(sNameOld As String, sNameNew As String, sFullPathOld As String, sFullPathNew As String, bFolder As Boolean, hItem As Long)
Public Event Initialized()
Public Event DropFiles(sFiles() As String, siaFiles As oleexp.IShellItemArray, doDropped As oleexp.IDataObject, sDropParent As String, siDropParent As oleexp.IShellItem, dwDropEffect As oleexp.DROPEFFECTS, dwKeyState As Long, ptPointX As Long, ptPointY As Long)
Public Event DragStart(sFiles() As String, siaDragged As oleexp.IShellItemArray, pdoDragged As oleexp.IDataObject)

Public Event EnterFocus()
Public Event ExitFocus()

Public Event DebugMessage(sMsg As String, iLevel As Integer)

Private bLoadDone As Boolean
Private pDTH As oleexp.IDropTargetHelper
Private Const CLSID_DragDropHelper = "{4657278A-411B-11D2-839A-00C04FD918D0}"
Private m_hWnd As Long
Private bRegDD As Boolean
Private psfCur As oleexp.IShellFolder
Private siaSelected As oleexp.IShellItemArray
Private sSelectedItems() As String
Private siaSIH As oleexp.IShellItemArray
Private lUDTC As Long, sUDsz As String
Private mTmrProc As Long

'todo: finish Refresh
Private mDefExpComp As Boolean
Private mDefExpLib As Boolean

Private mRefreshPaths() As String


Private hFontTV As Long
Private WithEvents PropFont As StdFont
Attribute PropFont.VB_VarHelpID = -1
Private mIFMain As IFont

Private lLoopTrack As Long

Private mBkm As Boolean

Private m_hRoot As Long
Private m_hFav As Long

Private mCustomRoot As String
Private Const mCustomRoot_def As String = ""
Private bCustRt As Boolean

Private sUserFolder As String
Private sUserDesktop As String

Private Type TVEntry
    bFolder As Boolean
    bZip As Boolean
    bLink As Boolean
    sFullPath As String
    sName As String
    sNameFull As String
    sParentFull As String
    sLinkTarget As String
    sInfoTip As String
    LinkPIDL As Long
    bLinkIsFolder As Boolean
    bDropTarget As Boolean
    hNode As Long
    hParentNode As Long
    nIcon As Long
    nOverlay As Long
    dwAttrib As SFGAO_Flags
    Checked As Boolean
    Excluded As Boolean
    bDisabled As Boolean
    bDeleted As Boolean
    pidlFQPar As Long
    pidlRel As Long
    pidlFQ As Long
    bIsDefItem As Boolean
End Type
Private TVEntries() As TVEntry
Private nCur As Long
Private TVVisMap() As TVEntry
Private nTVVis As Long
Private bFlagRecurseOI As Boolean

Private fRefreshing As Long

Private hTVD As Long
Private ICtxMenu2 As oleexp.IContextMenu2
Private ICtxMenu3 As oleexp.IContextMenu3
Private Const wIDSel As Long = 3000&
Private bSetParents As Boolean
Private bFilling As Boolean
Private gPaths() As String
Private nPaths As Long
Private gExPaths() As String
Private nExPaths As Long
Private nItr As Long
Private hSysIL As Long
Private pIML As iImageList
Private g_fDeleting As Boolean
Private fNoExpand As Long
Private sFavPath As String
Private siSelected As oleexp.IShellItem
Private sSelectedItem As String
Private gCurSelIdx As Long
Private m_hSHNotify As Long
Private Const WM_SHNOTIFY = &H477
Private hLEEdit As Long
Private sOldLEText As String
Private bRNf As Boolean
Private Const sComp = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
Private Const sLibRoot = "::{031E4825-7B94-4dc3-B131-E946B44C8DD5}\"
Private Const lnLibRoot As Long = 41& 'length
Private ddRightButton As Boolean
Private bNoDrop As Boolean
Private lHover1 As Long, lHover2 As Long
Private bFlagBlockClkExp As Boolean
Private hItemBlocked As Long
Private bBlockExec As Boolean
Private bNavigating As Boolean
Private fLoad As Long
Private sDesktopPath As String
Private m_cbSort As Long
Private m_sAltDrop As String
Private szNavWav As String
Private xHover As Long, yHover As Long
Private bHoverFired As Boolean
Private sFolder As String
Private bAbort As Boolean
Private bNoTarget As Boolean
Private lItemIndex As Long
Private mDragOver As String
Private mDropTipMsg As String
Private mDropTipIns As String
Private mDropTipImg As oleexp.DROPIMAGETYPE
Private mDefEffect As DROPEFFECTS
Private mAllowedEffects As DROPEFFECTS
Private mDataObj As oleexp.IDataObject
Private lvDragOverIdx As Long
Private IsComCtl6 As Boolean


Private clrBack As stdole.OLE_COLOR
Private clrFore As stdole.OLE_COLOR


Private mCheckboxes As Boolean
Private Const mCheckboxes_def As Boolean = False
Private mExCheckboxes As Boolean
Private Const mExCheckboxes_def As Boolean = False
Private mAutocheck As Boolean
Private Const mAutocheck_def As Boolean = True

Private mExplorerStyle As Boolean
Private Const mExplorerStyle_def As Boolean = True

Private mExpandZip As Boolean
Private Const mExpandZip_def As Boolean = False

Private mShowFiles As Boolean
Private Const mShowFiles_def As Boolean = False

Private mFadeExpandos As Boolean
Private Const mFadeExpandos_def As Boolean = True

Private mShowLines As Boolean
Private Const mShowLines_def As Boolean = False

Private mHasButtons As Boolean
Private Const mHasButtons_def As Boolean = True

Private mSingleExpand As Boolean
Private Const mSingleExpand_def As Boolean = False

Private mMultiSel As Boolean
Private Const mDefMultiSel As Boolean = False

Private mDisableDD As Boolean
Private Const mDefDisableDD As Boolean = False

Private mComputerAsRoot As Boolean
Private Const mComputerAsRoot_def As Boolean = False

Private mFullRowSelect As Boolean
Private Const mFullRowSelect_def As Boolean = True

Private mFilter As String
Private Const mFilter_def As String = "*.*"

Private mFilterFilesOnly As Boolean
Private Const mFilterFilesOnly_def As Boolean = True

Private mInfoTipOnFiles As Boolean
Private Const mInfoTipOnFiles_def As Boolean = True

Private m_TrackSel As Boolean
Private Const m_def_TrackSel As Boolean = True

Private mShowSelAlw As Boolean
Private Const mDefShowSelAlw As Boolean = True

Private mExpandOnLabelClick As Boolean
Private Const mExpandOnLabelClick_def As Boolean = False

Private mNoIndState As Boolean
Private Const mDefNoIndState As Boolean = False

Private mAutoHS As Boolean
Private Const mDefAutoHS As Boolean = True

Private mInfoTipOnFolders As Boolean
Private Const mInfoTipOnFolders_def As Boolean = False

Private mNavSound As Boolean
Private Const mNavSound_def As Boolean = True

Private mFavorites As Boolean
Private Const mFavorites_def As Boolean = True

Private mSHCN As Boolean
Private Const mSHCN_def As Boolean = True

Private mLabelEdit As Boolean
Private Const mLabelEdit_def As Boolean = True

Private mNameColors As Boolean
Private Const mNameColors_def As Boolean = True
Private m_SysClrText As Long

Private mExtOverlay As Boolean
Private Const m_def_ExtOverlay As Boolean = False

Private mAlwaysShowExtVerbs As Boolean
Private Const mDefAlwaysShowExtVerbs As Boolean = False

Private m_EnableShellMenu As Boolean
Private Const m_def_EnableShellMenu As Boolean = True

Public Enum ST_HDN_PREF
    STHP_UseExplorer = 0&
    STHP_AlwaysShow = 1&
    STHP_AlwaysHide = 2&
End Enum
Private m_HiddenPref As ST_HDN_PREF
Private Const m_def_HiddenPref As Long = 0&
Private mHPInExp As Boolean

Public Enum ST_SPRHDN_PREF
    STSHP_UseExplorer = 0&
    STSHP_AlwaysShow = 1&
    STSHP_AlwaysHide = 2&
End Enum
Private m_SuperHiddenPref As ST_SPRHDN_PREF
Private Const m_def_SuperHiddenPref As Long = 0&
Private mSHPInExp As Boolean

'Private mBorder As Boolean
'Private Const mBorder_def As Boolean = True
Public Enum ST_BORDERSTYLE
    STBS_None = 0&
    STBS_Standard = 1&
    STBS_Thick = 2&
    STBS_Thicker = 3&
End Enum
Private mBorder As ST_BORDERSTYLE
Private Const mBorder_def As Long = 1&

Private mInitialPath As String
Private Const mInitialPath_def As String = ""

Private mRootHasCheckbox As Boolean
Private Const mRootHasCheckbox_def As Boolean = False

Private mHScroll As Boolean
Private Const mHScroll_def As Boolean = True

Private lRaiseHover As Long
Private Const m_def_lRaiseHover As Long = 2500&

Private bTopLostFocus As Boolean
Private bHasFocus As Boolean
Private hBmpBack As Long

'------------------------------------------------------------
'BEGIN STANDARD APIs AND SYSTEM CONSTANTS

'I'm going to try to have a bitmap background at some point, but it's not ready yet.
Private Type PAINTSTRUCT
        hDC As Long
        fErase As Long
        rcPaint As RECT
        fRestore As Long
        fIncUpdate As Long
        rgbReserved(32) As Byte
End Type
Private Type BITMAP
    BMType As Long
    BMWidth As Long
    BMHeight As Long
    BMWidthBytes As Long
    BMPlanes As Integer
    BMBitsPixel As Integer
    BMBits As Long
End Type
Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal hSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long
Private Const TRANSPARENT = 1&
Private Const OPAQUE = 2&

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectW" (ByRef lpLogFont As LOGFONT) As Long
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageW" (ByRef lpMsg As Any) As Long
Private Declare Function DllGetVersion Lib "comctl32" (ByRef pdvi As DLLVERSIONINFO) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal VKey As Long) As Integer
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As oleexp.RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As Any) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetObjectW Lib "gdi32" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As SystemMetrics) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function ImageList_ReplaceIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hIcon As Long) As Long
Private Declare Function InsertMenuItemW Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Boolean, lpmii As MENUITEMINFOW) As Boolean
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function lstrcpyA Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long
Private Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long
Private Declare Function lstrlenW Lib "kernel32" (lpString As Any) As Long
Private Declare Function MessageBoxW Lib "user32.dll" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal wType As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function OleGetClipboard Lib "ole32" (ppDataObj As oleexp.IDataObject) As Long
Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal Color As Long, ByVal hpal As Long, ByRef RGBResult As Long) As Long
Private Declare Function PathFileExistsW Lib "shlwapi" (ByVal lpszPath As Long) As Long
Private Declare Function PathIsDirectoryW Lib "shlwapi" (ByVal lpszPath As Long) As Long
Private Declare Function PathMatchSpecW Lib "shlwapi" (ByVal pszFileParam As Long, ByVal pszSpec As Long) As Long
Private Declare Function PathMatchSpecExW Lib "shlwapi" (ByVal pszFile As Long, ByVal pszSpec As Long, ByVal dwFlags As ucst_PMS_Flags) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As MsgType, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundW" (ByVal lpszName As Long, ByVal hModule As Long, ByVal dwFlags As ucst_SND_FLAGS) As Long
Private Declare Function PSFormatPropertyValue Lib "propsys.dll" (ByVal pps As Long, ByVal ppd As Long, ByVal pdff As PROPDESC_FORMAT_FLAGS, ppszDisplay As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function RegisterClipboardFormatW Lib "user32" (ByVal lpszFormat As Long) As Long
Private Declare Function RegisterDragDrop Lib "ole32" (ByVal hWnd As Long, ByVal DropTarget As oleexp.IDropTarget) As Long
Private Declare Function RevokeDragDrop Lib "ole32" (ByVal hWnd As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As Any) As Long  ' lpPoint As oleexp.POINT) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal cy As Long, ByVal wFlags As ucst_SWP_Flags) As Long
Private Declare Function SetWindowTheme Lib "uxtheme" (ByVal hWnd As Long, ByVal pszSubAppName As Long, ByVal pszSubIdList As Long) As Long
Private Declare Function SHDoDragDrop Lib "Shell32" (ByVal hWnd As Long, ByVal pdtobj As Long, ByVal pdsrc As Long, ByVal dwEffect As Long, pdwEffect As Long) As Long
Private Declare Function SHFileOperationW Lib "Shell32" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function SHGetFileInfo Lib "Shell32" Alias "SHGetFileInfoA" (ByVal pszPath As Any, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As ucst_SHGFI_flags) As Long
Private Declare Function SHGetFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwReserved As Long, ppidl As Long) As Long
Private Declare Function SHGetImageList Lib "shell32.dll" (ByVal iImageList As ucst_ShellImageListFlags, ByRef riid As oleexp.UUID, ByRef ppv As Any) As Long
Private Declare Function SHGetKnownFolderIDList Lib "Shell32" (rfid As Any, ByVal dwFlags As Long, ByVal hToken As Long, pidl As Long) As Long
Private Declare Sub SHGetSetSettings Lib "shell32.dll" (ByRef lpss As oleexp.SHELLSTATE, ByVal dwMask As oleexp.SFS_MASK, Optional ByVal bSet As oleexp.BOOL)
Private Declare Function SHGetSettings Lib "Shell32" (lpsfs As Integer, ByVal dwMask As oleexp.SFS_MASK) As Long
Private Declare Function StrCmpLogicalW Lib "shlwapi" (ByVal lpStr1 As Long, ByVal lpStr2 As Long) As Long
Private Declare Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As ucst_TPM_wFlags, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lpRC As Any) As Long
Private Declare Function TranslateMessage Lib "user32" (ByRef lpMsg As Any) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long


Private Type MsgType
    hWnd        As Long
    message     As Long
    wParam      As Long
    lParam      As Long
    Time        As Long
    PT          As oleexp.POINT
End Type
Private Const PM_NOREMOVE           As Long = 0&
Private Const PM_REMOVE             As Long = 1&
'
Private Const CLR_NONE = &HFFFFFFFF
Private Const RDW_UPDATENOW As Long = &H100
Private Const RDW_INVALIDATE As Long = &H1
Private Const RDW_ERASE As Long = &H4
Private Const RDW_ALLCHILDREN As Long = &H80


Private Enum ucst_SWP_Flags
    SWP_NOSIZE = &H1
    SWP_NOMOVE = &H2
    SWP_NOZORDER = &H4
    SWP_NOREDRAW = &H8
    SWP_NOACTIVATE = &H10
    SWP_FRAMECHANGED = &H20
    SWP_DRAWFRAME = &H20
    SWP_SHOWWINDOW = &H40
    SWP_HIDEWINDOW = &H80
    SWP_NOCOPYBITS = &H100
    SWP_NOREPOSITION = &H200
    SWP_NOSENDCHANGING = &H400
    
    SWP_DEFERERASE = &H2000
    SWP_ASYNCWINDOWPOS = &H4000
End Enum

Private Const SM_CYFRAME         As Long = 33
Private Const SM_CYCAPTION = 4

Private Const IDC_ARROW = 32512&
Private Const IDC_WAIT = 32514&

Private Const COLOR_WINDOWTEXT = 8

Private Const UNICODE_NOCHAR = &HFFFF

Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90

Private Const LF_FACESIZE As Long = 32
Private Const FW_NORMAL As Long = 400
Private Const FW_BOLD As Long = 700
Private Const DEFAULT_QUALITY As Long = 0
Private Type LOGFONT
    LFHeight As Long
    LFWidth As Long
    LFEscapement As Long
    LFOrientation As Long
    LFWeight As Long
    LFItalic As Byte
    LFUnderline As Byte
    LFStrikeOut As Byte
    LFCharset As Byte
    LFOutPrecision As Byte
    LFClipPrecision As Byte
    LFQuality As Byte
    LFPitchAndFamily As Byte
    LFFaceName(0 To ((LF_FACESIZE * 2) - 1)) As Byte
End Type

Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformID As Long
End Type

Private Enum ucst_PMS_Flags
    PMSF_NORMAL = &H0
    PMSF_MULTIPLE = &H1
    PMSF_DONT_STRIP_SPACES = &H10000
End Enum

Private Enum ucst_IL_CreateFlags
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
Private Type IconHeader
    ihReserved      As Integer
    ihType          As Integer
    ihCount         As Integer
End Type
Private Type IconEntry
    ieWidth         As Byte
    ieHeight        As Byte
    ieColorCount    As Byte
    ieReserved      As Byte
    iePlanes        As Integer
    ieBitCount      As Integer
    ieBytesInRes    As Long
    ieImageOffset   As Long
End Type
Private Const EM_SETSEL = &HB1
Private CF_SHELLIDLIST As Long
Private CF_DROPDESCRIPTION As Long
Private CF_PREFERREDDROPEFFECT As Long
Private CF_COMPUTEDDRAGIMAGE As Long
Private CF_INDRAGLOOP As Long

Private Enum ucst_SHGFI_flags
  SHGFI_LARGEICON = &H0            ' sfi.hIcon is large icon
  SHGFI_SMALLICON = &H1            ' sfi.hIcon is small icon
  SHGFI_OPENICON = &H2              ' sfi.hIcon is open icon
  SHGFI_SHELLICONSIZE = &H4      ' sfi.hIcon is shell size (not system size), rtns BOOL
  SHGFI_PIDL = &H8                        ' pszPath is pidl, rtns BOOL
  ' Indicates that the function should not attempt to access the file specified by pszPath.
  ' Rather, it should act as if the file specified by pszPath exists with the file attributes
  ' passed in dwFileAttributes. This flag cannot be combined with the SHGFI_ATTRIBUTES,
  ' SHGFI_EXETYPE, or SHGFI_PIDL flags <---- !!!
  SHGFI_USEFILEATTRIBUTES = &H10   ' pretend pszPath exists, rtns BOOL
  SHGFI_ICON = &H100                    ' fills sfi.hIcon, rtns BOOL, use DestroyIcon
  SHGFI_DISPLAYNAME = &H200    ' isf.szDisplayName is filled (SHGDN_NORMAL), rtns BOOL
  SHGFI_TYPENAME = &H400          ' isf.szTypeName is filled, rtns BOOL
  SHGFI_ATTRIBUTES = &H800         ' rtns IShellFolder::GetAttributesOf  SFGAO_* flags
  SHGFI_ICONLOCATION = &H1000   ' fills sfi.szDisplayName with filename
                                                        ' containing the icon, rtns BOOL
  SHGFI_EXETYPE = &H2000            ' rtns two ASCII chars of exe type
  SHGFI_SYSICONINDEX = &H4000   ' sfi.iIcon is sys il icon index, rtns hImagelist
  SHGFI_LINKOVERLAY = &H8000    ' add shortcut overlay to sfi.hIcon
  SHGFI_SELECTED = &H10000        ' sfi.hIcon is selected icon
  SHGFI_ATTR_SPECIFIED = &H20000    ' get only attributes specified in sfi.dwAttributes
End Enum
Private Type SHFILEINFO   ' shfi
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type
Private Const NFR_UNICODE = 2
Private Const NM_FIRST As Long = 0&
Private Const NM_CLICK As Long = NM_FIRST - 2& 'uses NMCLICK struct
Private Const NM_DBLCLK As Long = NM_FIRST - 3&
Private Const NM_RETURN As Long = NM_FIRST - 4&
Private Const NM_RCLICK As Long = NM_FIRST - 5& 'uses NMCLICK struct
Private Const NM_RDBLCLK As Long = NM_FIRST - 6&
Private Const NM_CUSTOMDRAW As Long = NM_FIRST - 12&
Private Const WM_UNICHAR As Long = &H109
Private Const CDDS_PREPAINT As Long = &H1&
Private Const CDRF_NOTIFYITEMDRAW As Long = &H20&
Private Const CDRF_NOTIFYSUBITEMDRAW As Long = &H20&
Private Const CDDS_ITEM As Long = &H10000
Private Const CDDS_ITEMPREPAINT As Long = CDDS_ITEM Or CDDS_PREPAINT
Private Const CDDS_ITEMPOSTPAINT = (&H10000 Or &H2)
Private Const CDDS_SUBITEM = &H20000

Private Const CDRF_NEWFONT As Long = &H2&
Private Const CDRF_DODEFAULT As Long = &H0&

Private Type EDITBALLOONTIP
    cbStruct As Long
    pszTitle As Long
    pszText As Long
    ttiIcon As ucst_BalloonTipIconConstants ' ; // From TTI_*
End Type
Private Const ECM_FIRST As Long = &H1500
Private Const EM_SHOWBALLOONTIP As Long = (ECM_FIRST + 3)
Private Enum ucst_BalloonTipIconConstants
   TTI_NONE = 0
   TTI_INFO = 1
   TTI_WARNING = 2
   TTI_ERROR = 3
End Enum

Private Const MK_CONTROL = 8
Private Const MK_LBUTTON = 1
Private Const MK_MBUTTON = &H10
Private Const MK_RBUTTON = 2
Private Const MK_SHIFT = 4
Private Const MK_XBUTTON1 = &H20
Private Const MK_XBUTTON2 = &H40
Private Enum ucst_TPM_wFlags
  TPM_LEFTBUTTON = &H0
  TPM_RIGHTBUTTON = &H2
  TPM_LEFTALIGN = &H0
  TPM_CENTERALIGN = &H4
  TPM_RIGHTALIGN = &H8
  TPM_TOPALIGN = &H0
  TPM_VCENTERALIGN = &H10
  TPM_BOTTOMALIGN = &H20

  TPM_HORIZONTAL = &H0         ' Horz alignment matters more
  TPM_VERTICAL = &H40            ' Vert alignment matters more
  TPM_NONOTIFY = &H80           ' Don't send any notification msgs
  TPM_RETURNCMD = &H100
  
  TPM_HORPOSANIMATION = &H400
  TPM_HORNEGANIMATION = &H800
  TPM_VERPOSANIMATION = &H1000
  TPM_VERNEGANIMATION = &H2000
  TPM_NOANIMATION = &H4000
End Enum
Private Enum ucst_SND_FLAGS
    SND_ALIAS = &H10000 ' name is a WIN.INI [sounds] entry
    SND_ALIAS_ID = &H110000 ' name is a WIN.INI [sounds]Entry identifier
    SND_ALIAS_START = 0 ' must be > 4096 to keep strings insame section of resource file
    SND_APPLICATION = &H80 ' look for applicationspecific association
    SND_ASYNC = &H1 ' play asynchronously
    SND_FILENAME = &H20000 ' name is a file name
    SND_LOOP = &H8 ' loop the sound until nextsndPlaySound
    SND_MEMORY = &H4 ' lpszSoundName points to a memoryFile
    SND_NODEFAULT = &H2 ' silence not default, if soundnot found
    SND_NOSTOP = &H10 ' don't stop any currently playingsound
    SND_NOWAIT = &H2000 ' don't wait if the driver is busy
    SND_PURGE = &H40 ' purge non-static events forTask
    SND_RESERVED = &HFF000000 ' In particular these flags areReserved
    SND_RESOURCE = &H40004 ' name is a resource name or atom
    SND_SYNC = &H0 ' play synchronously (default)
    SND_TYPE_MASK = &H170007
    SND_VALID = &H1F ' valid flags / ;Internal /
    SND_VALIDFLAGS = &H17201F ' Set of valid flag bits.
End Enum

Private Enum ucst_ShellImageListFlags
    SHIL_LARGE = &H0
    SHIL_SMALL = &H1
    SHIL_EXTRALARGE = &H2
    SHIL_SYSSMALL = &H3
    '6.0
    SHIL_JUMBO = &H4
    SHIL_LAST = &H5 'NOT AN IMAGELIST
End Enum

Private Enum ucst_TVItemCheckStates
    tvcsNoBox = 0
    tvcsEmpty = 1
    tvcsChecked = 2
    tvcsPartial = 3
    tvcsExclude = 4
    'if you wish to add more check states (besides dimmed selected)
End Enum

Private Const MAX_ITEM = 256

' ============================================
' TREEVIEW COMPLETE DEFINITIONS
' ============================================
Private Const CCM_FIRST = &H2000

Private Const CCM_SETBKCOLOR = (CCM_FIRST + 1)   ' lParam is bkColor
Private Const CCM_SETCOLORSCHEME = (CCM_FIRST + 2)     ' lParam is color scheme
Private Const CCM_GETCOLORSCHEME = (CCM_FIRST + 3)     ' fills in COLORSCHEME pointed to by lParam
Private Const CCM_GETDROPTARGET = (CCM_FIRST + 4)
Private Const CCM_SETUNICODEFORMAT = (CCM_FIRST + 5)
Private Const CCM_GETUNICODEFORMAT = (CCM_FIRST + 6)
Private Const CCM_TRANSLATEACCELERATOR = &H461 '(WM_USER + 97)

Private Const IDD_TREEVIEW = 100

Private Enum ucst_TVHandles
  TVI_ROOT = &HFFFF0000
  TVI_FIRST = &HFFFF0001
  TVI_LAST = &HFFFF0002
  TVI_SORT = &HFFFF0003
End Enum

Private Const WC_TREEVIEW As String = "SysTreeView32"
' messages
Private Const I_CHILDRENCALLBACK = (-1)

Private Enum ucst_TVMessages
  TV_FIRST = &H1100
  
'#If UNICODE Then
  TVM_INSERTITEMW = (TV_FIRST + 50)
'#Else
  TVM_INSERTITEM = (TV_FIRST + 0)
'#End If
  
  TVM_DELETEITEM = (TV_FIRST + 1)
  TVM_EXPAND = (TV_FIRST + 2)
  TVM_GETITEMRECT = (TV_FIRST + 4)
  TVM_GETCOUNT = (TV_FIRST + 5)
  TVM_GETINDENT = (TV_FIRST + 6)
  TVM_SETINDENT = (TV_FIRST + 7)
  TVM_GETIMAGELIST = (TV_FIRST + 8)
  TVM_SETIMAGELIST = (TV_FIRST + 9)
  TVM_GETNEXTITEM = (TV_FIRST + 10)
  TVM_SELECTITEM = (TV_FIRST + 11)
  
'#If UNICODE Then
  TVM_GETITEMW = (TV_FIRST + 62)
  TVM_SETITEMW = (TV_FIRST + 63)
  TVM_EDITLABELW = (TV_FIRST + 65)
'#Else
  TVM_GETITEM = (TV_FIRST + 12)
  TVM_SETITEM = (TV_FIRST + 13)
  TVM_EDITLABEL = (TV_FIRST + 14)
'#End If
  
  TVM_GETEDITCONTROL = (TV_FIRST + 15)
  TVM_GETVISIBLECOUNT = (TV_FIRST + 16)
  TVM_HITTEST = (TV_FIRST + 17)
  TVM_CREATEDRAGIMAGE = (TV_FIRST + 18)
  TVM_SORTCHILDREN = (TV_FIRST + 19)
  TVM_ENSUREVISIBLE = (TV_FIRST + 20)
  TVM_SORTCHILDRENCB = (TV_FIRST + 21)
  TVM_ENDEDITLABELNOW = (TV_FIRST + 22)
  
'#If UNICODE Then
  TVM_GETISEARCHSTRINGW = (TV_FIRST + 64)
'#Else
  TVM_GETISEARCHSTRING = (TV_FIRST + 23)
'#End If
  
  TVM_SETTOOLTIPS = (TV_FIRST + 24)
  TVM_GETTOOLTIPS = (TV_FIRST + 25)
  TVM_SETINSERTMARK = (TV_FIRST + 26)
  TVM_SETUNICODEFORMAT = CCM_SETUNICODEFORMAT
  TVM_GETUNICODEFORMAT = CCM_GETUNICODEFORMAT
  TVM_SETITEMHEIGHT = (TV_FIRST + 27)
  TVM_GETITEMHEIGHT = (TV_FIRST + 28)
  TVM_SETBKCOLOR = (TV_FIRST + 29)
  TVM_SETTEXTCOLOR = (TV_FIRST + 30)
  TVM_GETBKCOLOR = (TV_FIRST + 31)
  TVM_GETTEXTCOLOR = (TV_FIRST + 32)
  TVM_SETSCROLLTIME = (TV_FIRST + 33)
  TVM_GETSCROLLTIME = (TV_FIRST + 34)
  TVM_SETBORDER = (TV_FIRST + 35)

  TVM_SETINSERTMARKCOLOR = (TV_FIRST + 37)
  TVM_GETINSERTMARKCOLOR = (TV_FIRST + 38)
  '5.0
  TVM_GETITEMSTATE = (TV_FIRST + 39)
  TVM_SETLINECOLOR = (TV_FIRST + 40)
  TVM_GETLINECOLOR = (TV_FIRST + 41)
  TVM_MAPACCIDTOHTREEITEM = (TV_FIRST + 42)
  TVM_MAPHTREEITEMTOACCID = (TV_FIRST + 43)
  TVM_SETEXTENDEDSTYLE = (TV_FIRST + 44)
  TVM_GETEXTENDEDSTYLE = (TV_FIRST + 45)
  TVM_SETHOT = (TV_FIRST + 58)
  TVM_SETAUTOSCROLLINFO = (TV_FIRST + 59)
  '6.0
  TVM_GETSELECTEDCOUNT = (TV_FIRST + 70)
  TVM_SHOWINFOTIP = (TV_FIRST + 71)
  TVM_GETITEMPARTRECT = (TV_FIRST + 72)
    
End Enum   ' TVMessages

Private Enum ucst_TV_Styles
    TVS_HASBUTTONS = &H1
    TVS_HASLINES = &H2
    TVS_LINESATROOT = &H4
    TVS_EDITLABELS = &H8
    TVS_DISABLEDRAGDROP = &H10
    TVS_SHOWSELALWAYS = &H20
    TVS_RTLREADING = &H40
    TVS_NOTOOLTIPS = &H80
    TVS_CHECKBOXES = &H100
    TVS_TRACKSELECT = &H200
    TVS_SINGLEEXPAND = &H400
    TVS_INFOTIP = &H800
    TVS_FULLROWSELECT = &H1000
    TVS_NOSCROLL = &H2000
    TVS_NONEVENHEIGHT = &H4000
    TVS_NOHSCROLL = &H8000
End Enum
Private Enum ucst_TV_Ex_Styles
    TVS_EX_NOSINGLECOLLAPSE = &H1
    TVS_EX_MULTISELECT = &H2
    TVS_EX_DOUBLEBUFFER = &H4
    TVS_EX_NOINDENTSTATE = &H8
    TVS_EX_RICHTOOLTIP = &H10
    TVS_EX_AUTOHSCROLL = &H20
    TVS_EX_FADEINOUTEXPANDOS = &H40
    TVS_EX_PARTIALCHECKBOXES = &H80
    TVS_EX_EXCLUSIONCHECKBOXES = &H100
    TVS_EX_DIMMEDCHECKBOXES = &H200
    TVS_EX_DRAWIMAGEASYNC = &H400
End Enum

Private Enum ucst_TVSB_Flags
    TVSBF_XBORDER = &H1
    TVSBF_YBORDER = &H2
End Enum
' TVM_GET/SETIMAGELIST wParam
Private Enum ucst_TVImageLists
    TVSIL_NORMAL = 0
    TVSIL_STATE = 2
End Enum

' TVM_GETNEXTITEM wParam
Private Enum ucst_TVM_GETNEXTITEM_wParam
  TVGN_ROOT = &H0
  TVGN_NEXT = &H1
  TVGN_PREVIOUS = &H2
  TVGN_PARENT = &H3
  TVGN_CHILD = &H4
  TVGN_FIRSTVISIBLE = &H5
  TVGN_NEXTVISIBLE = &H6
  TVGN_PREVIOUSVISIBLE = &H7
  TVGN_DROPHILITE = &H8
  TVGN_CARET = &H9
  TVGN_LASTVISIBLE = &HA
  TVGN_NEXTSELECTED = &HB

End Enum
Private Const TVSI_NOSINGLEEXPAND = &H8000

' TVM_GET/SETITEM lParam
Private Type TVITEM   'TVITEMW
  Mask As ucst_TVITEM_Mask
  hItem As Long
  State As ucst_TVITEM_State
  StateMask As ucst_TVITEM_State
  pszText As Long
  cchTextMax As Long
  iImage As Long
  iSelectedImage As Long
  cChildren As Long
  lParam As Long
End Type
Private Type TVITEMA   ' was TV_ITEM
  Mask As ucst_TVITEM_Mask
  hItem As Long
  State As ucst_TVITEM_State
  StateMask As ucst_TVITEM_State
  pszText As String    ' if a string, must be pre-allocated!!
  cchTextMax As Long
  iImage As Long
  iSelectedImage As Long
  cChildren As Long
  lParam As Long
End Type

Private Type TVITEMEX 'TVITEMEXW
    Mask As ucst_TVITEM_Mask
    hItem As Long
    State As ucst_TVITEM_State
    StateMask As ucst_TVITEM_State
    pszText As Long
    cchTextMax As Long
    iImage As Long
    iSelectedImage As Long
    cChildren As Long
    lParam As Long
    iIntegral As Long
    uStateEx As ucst_TVITEM_State_Ex
    hWnd As Long
    iExpandedImage As Long
    iReserved As Long 'Win7
End Type
Private Type TVITEMEXA
    Mask As ucst_TVITEM_Mask
    hItem As Long
    State As ucst_TVITEM_State
    StateMask As ucst_TVITEM_State
    pszText As String
    cchTextMax As Long
    iImage As Long
    iSelectedImage As Long
    cChildren As Long
    lParam As Long
    iIntegral As Long
    uStateEx As ucst_TVITEM_State_Ex
    hWnd As Long
    iExpandedImage As Long
    iReserved As Long 'Win7
End Type

'TVINSERTSTRUCT is supposed to be a union where item is TVITEMA/W or EXA/EXW
'but VB doesn't support unions or 'As Any' in structs. Macros using these
'also had to be duplicated
Private Type TVINSERTSTRUCT
    hParent As Long
    hInsertAfter As Long
    Item As TVITEM
End Type
Private Type TVINSERTSTRUCTEX
    hParent As Long
    hInsertAfter As Long
    Item As TVITEMEX
End Type

Private Enum ucst_TVITEMPART
    TVGIPR_BUTTON = &H1
End Enum

Private Type TVGETITEMPARTRECTINFO
    hti As Long
    prc As oleexp.RECT
    partid As ucst_TVITEMPART
End Type

' TVITEM mask
Private Enum ucst_TVITEM_Mask
    TVIF_TEXT = &H1
    TVIF_IMAGE = &H2
    TVIF_PARAM = &H4
    TVIF_STATE = &H8
    TVIF_HANDLE = &H10
    TVIF_SELECTEDIMAGE = &H20
    TVIF_CHILDREN = &H40
    TVIF_INTEGRAL = &H80
    '6.0
    TVIF_STATEEX = &H100
    TVIF_EXPANDEDIMAGE = &H200
    TVIF_DI_SETITEM = &H1000
End Enum
' TVITEM state, stateMask
Private Enum ucst_TVITEM_State
    TVIS_SELECTED = &H2
    TVIS_CUT = &H4
    TVIS_DROPHILITED = &H8
    TVIS_BOLD = &H10
    TVIS_EXPANDED = &H20
    TVIS_EXPANDEDONCE = &H40
    TVIS_EXPANDPARTIAL = &H80
    TVIS_OVERLAYMASK = &HF00
    TVIS_STATEIMAGEMASK = &HF000
    TVIS_USERMASK = &HF000
End Enum
Private Enum ucst_TVITEM_State_Ex
    '6.0
    TVIS_EX_FLAT = &H1
    TVIS_EX_DISABLED = &H2
    TVIS_EX_ALL = &H2
End Enum
' TVM_HITTEST lParam
Private Type TVHITTESTINFO   ' was TV_HITTESTINFO
  PT As oleexp.POINT
  Flags As ucst_TVHT_flags
  hItem As Long
End Type

Private Enum ucst_TVHT_flags
  TVHT_NOWHERE = &H1   ' In the client area, but below the last item
  TVHT_ONITEMICON = &H2
  TVHT_ONITEMLABEL = &H4
  TVHT_ONITEMINDENT = &H8
  TVHT_ONITEMBUTTON = &H10
  TVHT_ONITEMRIGHT = &H20
  TVHT_ONITEMSTATEICON = &H40
  TVHT_ONITEM = (TVHT_ONITEMICON Or TVHT_ONITEMLABEL Or TVHT_ONITEMSTATEICON)
  
  TVHT_ABOVE = &H100
  TVHT_BELOW = &H200
  TVHT_TORIGHT = &H400
  TVHT_TOLEFT = &H800
  
  ' user-defined
  TVHT_ONITEMLINE = (TVHT_ONITEM Or TVHT_ONITEMINDENT Or TVHT_ONITEMBUTTON Or TVHT_ONITEMRIGHT)
End Enum

' TVM_SORTCHILDRENCB lParam
Private Type TVSORTCB   ' was TV_SORTCB
  hParent As Long
  lpfnCompare As Long
  lParam As Long
End Type
Private Enum ucst_TVM_EXPAND_wParam
  TVE_COLLAPSE = &H1
  TVE_EXPAND = &H2
  TVE_TOGGLE = &H3

  TVE_EXPANDPARTIAL = &H4000
  
  TVE_COLLAPSERESET = &H8000
End Enum
Private Const TVC_UNKNOWN = &H0
Private Const TVC_BYMOUSE = &H1
Private Const TVC_BYKEYBOARD = &H2

' notifications
Private Enum ucst_TVNotifications
  TVN_FIRST = -400&   ' &HFFFFFE70   ' (0U-400U)
  TVN_LAST = -499&    ' &HFFFFFE0D    ' (0U-499U)
                                                      ' lParam points to:
  TVN_SELCHANGINGA = (TVN_FIRST - 1)          ' NMTREEVIEW
  TVN_SELCHANGEDA = (TVN_FIRST - 2)           ' NMTREEVIEW
  TVN_GETDISPINFOA = (TVN_FIRST - 3)            ' NMTVDISPINFO
  TVN_SETDISPINFOA = (TVN_FIRST - 4)            ' NMTVDISPINFO
  TVN_ITEMEXPANDINGA = (TVN_FIRST - 5)       ' NMTREEVIEW
  TVN_ITEMEXPANDEDA = (TVN_FIRST - 6)        ' NMTREEVIEW
  TVN_BEGINDRAGA = (TVN_FIRST - 7)              ' NMTREEVIEW
  TVN_BEGINRDRAGA = (TVN_FIRST - 8)            ' NMTREEVIEW
  TVN_DELETEITEMA = (TVN_FIRST - 9)             ' NMTREEVIEW
  TVN_BEGINLABELEDITA = (TVN_FIRST - 10)    ' NMTVDISPINFO
  TVN_ENDLABELEDITA = (TVN_FIRST - 11)       ' NMTVDISPINFO
  TVN_KEYDOWN = (TVN_FIRST - 12)                ' NMTVKEYDOWN

  TVN_SELCHANGINGW = (TVN_FIRST - 50)
  TVN_SELCHANGEDW = (TVN_FIRST - 51)
  TVN_GETDISPINFOW = (TVN_FIRST - 52)
  TVN_SETDISPINFOW = (TVN_FIRST - 53)
  TVN_ITEMEXPANDINGW = (TVN_FIRST - 54)
  TVN_ITEMEXPANDEDW = (TVN_FIRST - 55)
  TVN_BEGINDRAGW = (TVN_FIRST - 56)
  TVN_BEGINRDRAGW = (TVN_FIRST - 57)
  TVN_DELETEITEMW = (TVN_FIRST - 58)
  TVN_BEGINLABELEDITW = (TVN_FIRST - 59)
  TVN_ENDLABELEDITW = (TVN_FIRST - 60)

  TVN_GETINFOTIPA = (TVN_FIRST - 13)
  TVN_GETINFOTIPW = (TVN_FIRST - 14)
  TVN_SINGLEEXPAND = (TVN_FIRST - 15)
    TVN_ITEMCHANGINGA = (TVN_FIRST - 16)
    TVN_ITEMCHANGINGW = (TVN_FIRST - 17)
    TVN_ITEMCHANGEDA = (TVN_FIRST - 18)
    TVN_ITEMCHANGEDW = (TVN_FIRST - 19)
    TVN_ASYNCDRAW = (TVN_FIRST - 20)

End Enum   ' Notifications

Private Const TVNRET_DEFAULT = &H0
Private Const TVNRET_SKIPOLD = &H1
Private Const TVNRET_SKIPNEW = &H2

' lParam for most treeview notification messages
Private Type NMTREEVIEW   ' was NM_TREEVIEW
  hdr As NMHDR
  ' Specifies a notification-specific action flag.
  ' Is TVC_* for TVN_SELCHANGING, TVN_SELCHANGED, TVN_SETDISPINFO
  ' Is TVE_* for TVN_ITEMEXPANDING, TVN_ITEMEXPANDED
  Action As Long
  itemOld As TVITEM
  itemNew As TVITEM
  ptDrag As oleexp.POINT
End Type
Private Type NMCUSTOMDRAW
    hdr As NMHDR
    dwDrawStage As Long
    hDC As Long
    rc As oleexp.RECT
    dwItemSpec As Long
    uItemState As Long
    lItemlParam As Long
End Type
Private Type NMTVSTATEIMAGECHANGING
    hdr As NMHDR
    hti As Long
    iOldStateImageIndex As Long
    iNewStateImageIndex As Long
End Type
Private Type NMTVDISPINFO
    hdr As NMHDR
    Item As TVITEM
End Type
Private Type NMTVDISPINFOEX
    hdr As NMHDR
    Item As TVITEMEX
End Type
Private Type NMTVKEYDOWN
    hdr As NMHDR
    wVKey As Integer
    Flags As Long
End Type
Private Const TVCDRF_NOIMAGES = &H10000
Private Type NMTVCUSTOMDRAW
    NMCD As NMCUSTOMDRAW
    ClrText As Long
    ClrTextBk As Long
    iLevel As Long
End Type
Private Type NMTVGETINFOTIP
    hdr As NMHDR
    pszText As Long
    cchTextMax As Long
    hItem As Long
    lParam As Long
End Type
Private Type NMTVITEMCHANGE
    hdr As NMHDR
    uChanged As Long
    hItem As Long
    uStateNew As ucst_TVITEM_State
    uStateOld As ucst_TVITEM_State
    lParam As Long
End Type
Private Type NMTVASYNCDRAW
    hdr As NMHDR
    pimldp As IMAGELISTDRAWPARAMS
    hr As Long
    hItem As Long
    lParam As Long
    dwRetFlags As Long
    iRetImageIndex As Long
End Type

Private Type MENUITEMINFOW
  cbSize As Long
  fMask As ucst_MII_Mask
  fType As ucst_MF_Type              ' MIIM_TYPE
  fState As ucst_MF_State             ' MIIM_STATE
  wID As Long                       ' MIIM_ID
  hSubMenu As Long            ' MIIM_SUBMENU
  hbmpChecked As Long      ' MIIM_CHECKMARKS
  hbmpUnchecked As Long  ' MIIM_CHECKMARKS
  dwItemData As Long          ' MIIM_DATA
  dwTypeData As Long        ' MIIM_TYPE
  cch As Long                       ' MIIM_TYPE
  hbmpItem As Long
End Type
Private Enum ucst_MII_Mask
  MIIM_STATE = &H1
  MIIM_ID = &H2
  MIIM_SUBMENU = &H4
  MIIM_CHECKMARKS = &H8
  MIIM_TYPE = &H10
  MIIM_DATA = &H20
  MIIM_BITMAP = &H80
  MIIM_STRING = &H40
End Enum
Private Enum ucst_MenuFlags
  MF_INSERT = &H0
  MF_ENABLED = &H0
  MF_UNCHECKED = &H0
  MF_BYCOMMAND = &H0
  MF_STRING = &H0
  MF_UNHILITE = &H0
  MF_GRAYED = &H1
  MF_DISABLED = &H2
  MF_BITMAP = &H4
  MF_CHECKED = &H8
  MF_POPUP = &H10
  MF_MENUBARBREAK = &H20
  MF_MENUBREAK = &H40
  MF_HILITE = &H80
  MF_CHANGE = &H80
  MF_END = &H80                    ' Obsolete -- only used by old RES files
  MF_APPEND = &H100
  MF_OWNERDRAW = &H100
  MF_DELETE = &H200
  MF_USECHECKBITMAPS = &H200
  MF_BYPOSITION = &H400
  MF_SEPARATOR = &H800
  MF_REMOVE = &H1000
  MF_DEFAULT = &H1000
  MF_SYSMENU = &H2000
  MF_HELP = &H4000
  MF_RIGHTJUSTIFY = &H4000
  MF_MOUSESELECT = &H8000&
End Enum

Private Enum ucst_MF_Type
  MFT_STRING = MF_STRING
  MFT_BITMAP = MF_BITMAP
  MFT_MENUBARBREAK = MF_MENUBARBREAK
  MFT_MENUBREAK = MF_MENUBREAK
  MFT_OWNERDRAW = MF_OWNERDRAW
  MFT_RADIOCHECK = &H200
  MFT_SEPARATOR = MF_SEPARATOR
  MFT_RIGHTORDER = &H2000
  MFT_RIGHTJUSTIFY = MF_RIGHTJUSTIFY
End Enum
Private Enum ucst_MF_State
  MFS_GRAYED = &H3
  MFS_DISABLED = MFS_GRAYED
  MFS_CHECKED = MF_CHECKED
  MFS_HILITE = MF_HILITE
  MFS_ENABLED = MF_ENABLED
  MFS_UNCHECKED = MF_UNCHECKED
  MFS_UNHILITE = MF_UNHILITE
  MFS_DEFAULT = MF_DEFAULT
End Enum

Private Type SHFILEOPSTRUCT
   hWnd        As Long
   wFunc       As FILEOP
   pFrom       As Long
   pTo         As Long
   fFlags      As FILEOP_FLAGS
   fAborted    As Boolean
   hNameMaps   As Long
   sProgress   As Long
 End Type

'========================================================================================
' mIOLEInPlaceActiveObject Implementation
' Author:      Mike Gainer, Matt Curland and Bill Storage
'
' Requires:    OleGuids.tlb (in IDE only)
'
' Description:
' Allows you to replace the standard IOLEInPlaceActiveObject interface for a
' UserControl with a customisable one.  This allows you to take control
' of focus in VB controls.
'
' The code could be adapted to replace other UserControl OLE interfaces.
'
' ---------------------------------------------------------------------------------------
' Visit vbAccelerator, advanced, free source for VB programmers
' http://vbaccelerator.com
'========================================================================================
Private Type IPAOHookStruct
    lpVTable    As Long                    'VTable pointer
    IPAOReal    As Long 'IOleInPlaceActiveObject 'Un-AddRefed pointer for forwarding calls
    ThisPointer As Long
End Type
Private m_uIPAO         As IPAOHookStruct
Private Declare Function IsEqualGUID Lib "ole32" (iid1 As oleexp.UUID, iid2 As oleexp.UUID) As Long

Private Type OLEINPLACEFRAMEINFO
    cb              As Long
    fMDIApp         As Boolean
    hwndFrame       As Long
    haccel          As Long
    cAccelEntries   As Long
End Type

'Private Type POINT
'    x As Long
'    y As Long
'End Type

Private Type Msg
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    Time As Long
    PT As POINT
End Type 'MSG

'Private Const S_FALSE               As Long = 1
'Private Const S_OK                  As Long = 0

Private IID_IOleInPlaceActiveObject As oleexp.UUID
Private m_IPAOVTable(9)             As Long

'*************************************************************************************************
' ==== Used by CallInterface Function =====================================================
'
'Private Type uuid
'  Data1         As Long
'  Data2         As Integer
'  Data3         As Integer
'  Data4(0 To 7) As Byte
'End Type

Private Enum ucst_IUnknown_Exports
    [QueryInterface] = 0
    [AddRef] = 1
    [Release] = 2
End Enum

Private Enum ucst_IPAO_Exports
    [GetWindow] = 3
    [ContextSensitiveHelp] = 4
    [TranslateAccelerator] = 5
    [OnFrameWindowActivate] = 6
    [OnDocWindowActivate] = 7
    [ResizeBorder] = 8
    [EnableModeless] = 9
End Enum

Private Declare Function PutMem2 Lib "msvbvm60" (ByVal pWORDDst As Long, ByVal newValue As Long) As Long
Private Declare Function PutMem4 Lib "msvbvm60" (ByVal pDWORDDst As Long, ByVal newValue As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60" (ByVal pDWORDSrc As Long, ByVal pDWORDDst As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function IIDFromString Lib "ole32.dll" (ByVal lpsz As Long, lpiid As oleexp.UUID) As Long

Private Const szIID_IOleInPlaceActive     As String = "{00000117-0000-0000-C000-000000000046}"
Private Const szIID_IOleObject            As String = "{00000112-0000-0000-C000-000000000046}"
Private Const szIID_IOleInPlaceSite       As String = "{00000119-0000-0000-C000-000000000046}"
Private Const szIID_IOleControlSite       As String = "{B196B289-BAB4-101A-B69C-00AA00341D07}"
Private ptrMe As Long

Private Const GMEM_FIXED As Long = &H0
Private Const asmPUSH_imm32 As Byte = &H68
Private Const asmRET_imm16 As Byte = &HC2
Private Const asmCALL_rel32 As Byte = &HE8

' === Subclassing ========================================================
' Subclasing by Paul Caton
Private z_scFunk            As Collection   'hWnd/thunk-address collection
Private z_hkFunk            As Collection   'hook/thunk-address collection
Private z_cbFunk            As Collection   'callback/thunk-address collection
Private Const IDX_INDEX     As Long = 2     'index of the subclassed hWnd OR hook type
Private Const IDX_PREVPROC  As Long = 9     'Thunk data index of the original WndProc
Private Const IDX_BTABLE    As Long = 11    'Thunk data index of the Before table for messages
Private Const IDX_ATABLE    As Long = 12    'Thunk data index of the After table for messages
Private Const IDX_CALLBACKORDINAL As Long = 36 ' Ubound(callback thunkdata)+1, index of the callback

' Declarations:
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpFN As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Enum ucst_eThunkType
    SubclassThunk = 0
    CallbackThunk = 2
End Enum

Private Enum ucst_eMsgWhen                                                   'When to callback
  MSG_BEFORE = 1                                                        'Callback before the original WndProc
  MSG_AFTER = 2                                                         'Callback after the original WndProc
  MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                            'Callback before and after the original WndProc
End Enum

Private Const IDX_PARM_USER As Long = 13    'Thunk data index of the User-defined callback parameter data index
Private Const IDX_UNICODE   As Long = 107   'Must be UBound(subclass thunkdata)+1; index for unicode support
Private Const MSG_ENTRIES   As Long = 32    'Number of msg table entries. Set to 1 if using ALL_MESSAGES for all subclassed windows

Private Enum ucst_eAllMessages
    ALL_MESSAGES = -1     'All messages will callback
End Enum

Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProcW Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'***************************************************************************
'
'----------------------------END DECLARE SECTION---------------------------

#If False Then
    Dim TV_FIRST, TVM_INSERTITEM, TVM_INSERTITEM, TVM_DELETEITEM, TVM_EXPAND, TVM_GETITEMRECT, TVM_GETCOUNT, TVM_GETINDENT, TVM_SETINDENT, TVM_GETIMAGELIST, TVM_SETIMAGELIST, TVM_GETNEXTITEM, TVM_SELECTITEM, TVM_GETITEMW, TVM_SETITEMW, TVM_EDITLABELW, TVM_GETITEM, TVM_SETITEM, TVM_EDITLABEL, TVM_GETEDITCONTROL, TVM_GETVISIBLECOUNT, TVM_HITTEST, TVM_CREATEDRAGIMAGE, TVM_SORTCHILDREN, TVM_ENSUREVISIBLE, TVM_SORTCHILDRENCB, TVM_ENDEDITLABELNOW, TVM_GETISEARCHSTRINGW, TVM_GETISEARCHSTRING, TVM_SETTOOLTIPS, TVM_GETTOOLTIPS, TVM_SETINSERTMARK, TVM_SETUNICODEFORMAT, TVM_GETUNICODEFORMAT, TVM_SETITEMHEIGHT, TVM_GETITEMHEIGHT, TVM_SETBKCOLOR, TVM_SETTEXTCOLOR, TVM_GETBKCOLOR, TVM_GETTEXTCOLOR, TVM_SETSCROLLTIME, TVM_GETSCROLLTIME, TVM_SETINSERTMARKCOLOR, TVM_GETINSERTMARKCOLOR, TVM_GETITEMSTATE, TVM_SETEXTENDEDSTYLE, TVM_GETEXTENDEDSTYLE
#End If

#If False Then
    Dim TVS_HASBUTTONS, TVS_HASLINES, TVS_LINESATROOT, TVS_EDITLABELS, TVS_DISABLEDRAGDROP, TVS_SHOWSELALWAYS, TVS_RTLREADING, TVS_NOTOOLTIPS, TVS_CHECKBOXES, TVS_TRACKSELECT, TVS_SINGLEEXPAND, TVS_INFOTIP, TVS_FULLROWSELECT, TVS_NOSCROLL, TVS_NONEVENHEIGHT, TVS_NOHSCROLL
#End If

#If False Then
    Dim TVS_EX_MULTISELECT, TVS_EX_DOUBLEBUFFER, TVS_EX_NOINDENTSTATE, TVS_EX_RICHTOOLTIP, TVS_EX_AUTOHSCROLL, TVS_EX_FADEINOUTEXPANDOS, TVS_EX_PARTIALCHECKBOXES, TVS_EX_EXCLUSIONCHECKBOXES, TVS_EX_DIMMEDCHECKBOXES, TVS_EX_DRAWIMAGEASYNC
#End If

#If False Then
    Dim TVI_ROOT, TVI_FIRST, TVI_LAST, TVI_SORT, TVGN_ROOT, TVGN_NEXT, TVGN_PREVIOUS, TVGN_PARENT, TVGN_CHILD, TVGN_FIRSTVISIBLE, TVGN_NEXTVISIBLE, TVGN_PREVIOUSVISIBLE, TVGN_DROPHILITE, TVGN_CARET, TVGN_LASTVISIBLE
#End If

#If False Then
    Dim TVIF_TEXT, TVIF_IMAGE, TVIF_PARAM, TVIF_STATE, TVIF_SELECTEDIMAGE, TVIF_CHILDREN, TVIF_INTEGRAL, TVIF_STATEEX, TVIF_EXPANDEDIMAGE
#End If

#If False Then
    Dim TVIS_SELECTED, TVIS_CUT, TVIS_DROPHILITED, TVIS_BOLD, TVIS_EXPANDED, TVIS_EXPANDEDONCE, TVIS_EXPANDPARTIAL, TVIS_OVERLAYMASK, TVIS_STATEIMAGEMASK, TVIS_USERMASK
#End If

#If False Then
    Dim TVHT_NOWHERE, TVHT_ONITEMICON, TVHT_ONITEMLABEL, TVHT_ONITEMINDENT, TVHT_ONITEMBUTTON, TVHT_ONITEMRIGHT, TVHT_ONITEMSTATEICON, TVHT_ONITEM, TVHT_ONITEMLINE
#End If

#If False Then
    Dim TVN_FIRST, TVN_LAST, TVN_SELCHANGINGA, TVN_SELCHANGEDA, TVN_GETDISPINFOA, TVN_SETDISPINFOA, TVN_ITEMEXPANDINGA, TVN_ITEMEXPANDEDA, TVN_BEGINDRAGA, TVN_BEGINRDRAGA, TVN_DELETEITEMA, TVN_BEGINLABELEDITA, TVN_ENDLABELEDITA, TVN_KEYDOWN, TVN_SELCHANGINGW, TVN_SELCHANGEDW, TVN_GETDISPINFOW, TVN_SETDISPINFOW, TVN_ITEMEXPANDINGW, TVN_ITEMEXPANDEDW, TVN_BEGINDRAGW, TVN_BEGINRDRAGW, TVN_DELETEITEMW, TVN_BEGINLABELEDITW, TVN_ENDLABELEDITW, TVN_GETINFOTIPA, TVN_GETINFOTIPW, TVN_SINGLEEXPAND, TVN_ITEMCHANGINGA, TVN_ITEMCHANGINGW, TVN_ITEMCHANGEDA, TVN_ITEMCHANGEDW, TVN_ASYNCDRAW
#End If
#If False Then
Dim TVSIL_NORMAL, TVSIL_STATE
#End If
#If False Then
Dim TVIS_EX_FLAT, TVIS_EX_DISABLED, TVIS_EX_ALL
#End If
#If False Then
Dim TVE_EXPAND, TVE_COLLAPSE, TVE_TOGGLE, TVE_EXPANDPARTIAL, TVE_COLLAPSERESET
#End If
#If False Then
Dim TV_FIRST, TVM_INSERTITEMW, TVM_INSERTITEM, TVM_DELETEITEM, TVM_EXPAND, TVM_GETITEMRECT, _
TVM_GETCOUNT, TVM_GETINDENT, TVM_SETINDENT, TVM_GETIMAGELIST, TVM_SETIMAGELIST, _
TVM_GETNEXTITEM, TVM_SELECTITEM, TVM_GETITEMW, TVM_SETITEMW, TVM_EDITLABELW, TVM_GETITEM, _
TVM_SETITEM, TVM_EDITLABEL, TVM_GETEDITCONTROL, TVM_GETVISIBLECOUNT, TVM_HITTEST, _
TVM_CREATEDRAGIMAGE, TVM_SORTCHILDREN, TVM_ENSUREVISIBLE, TVM_SORTCHILDRENCB, _
TVM_ENDEDITLABELNOW, TVM_GETISEARCHSTRINGW, TVM_GETISEARCHSTRING, TVM_SETTOOLTIPS, _
TVM_GETTOOLTIPS, TVM_SETINSERTMARK, TVM_SETUNICODEFORMAT, TVM_GETUNICODEFORMAT, _
TVM_SETITEMHEIGHT, TVM_GETITEMHEIGHT, TVM_SETBKCOLOR, TVM_SETTEXTCOLOR, TVM_GETBKCOLOR, _
TVM_GETTEXTCOLOR, TVM_SETSCROLLTIME, TVM_GETSCROLLTIME, TVM_SETINSERTMARKCOLOR, _
TVM_GETINSERTMARKCOLOR, TVM_GETITEMSTATE, TVM_SETLINECOLOR, TVM_GETLINECOLOR, _
TVM_MAPACCIDTOHTREEITEM, TVM_MAPHTREEITEMTOACCID, TVM_SETEXTENDEDSTYLE, TVM_GETEXTENDEDSTYLE, _
TVM_SETAUTOSCROLLINFO, TVM_GETSELECTEDCOUNT, TVM_SHOWINFOTIP, TVM_GETITEMPARTRECT
#End If
#If False Then
Dim STBS_None, STBS_Standard, STBS_Thick, STBS_Thicker
#End If

#If False Then
Dim tvcsEmpty, tvcsChecked, tvcsPartial
#End If

'-SelfSub code------------------------------------------------------------------------------------
'-The following routines are exclusively for the ssc_Subclass routines----------------------------
Private Function ssc_Subclass(ByVal lng_hWnd As Long, _
                    Optional ByVal lParamUser As Long = 0, _
                    Optional ByVal nOrdinal As Long = 1, _
                    Optional ByVal oCallback As Object = Nothing, _
                    Optional ByVal bIdeSafety As Boolean = True, _
                    Optional ByRef bUnicode As Boolean = False, _
                    Optional ByVal bIsAPIwindow As Boolean = False) As Boolean 'Subclass the specified window handle

    '*************************************************************************************************
    '* lng_hWnd   - Handle of the window to subclass
    '* lParamUser - Optional, user-defined callback parameter
    '* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
    '* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
    '* bIdeSafety - Optional, enable/disable IDE safety measures. There is not reason to set this to False
    '* bUnicode - Optional, if True, Unicode API calls should be made to the window vs ANSI calls
    '*            Parameter is byRef and its return value should be checked to know if ANSI to be used or not
    '* bIsAPIwindow - Optional, if True DestroyWindow will be called if IDE ENDs
    '*****************************************************************************************
    '** Subclass.asm - subclassing thunk
    '**
    '** Paul_Caton@hotmail.com
    '** Copyright free, use and abuse as you see fit.
    '**
    '** v2.0 Re-write by LaVolpe, based mostly on Paul Caton's original thunks....... 20070720
    '** .... Reorganized & provided following additional logic
    '** ....... Unsubclassing only occurs after thunk is no longer recursed
    '** ....... Flag used to bypass callbacks until unsubclassing can occur
    '** ....... Timer used as delay mechanism to free thunk memory afer unsubclassing occurs
    '** .............. Prevents crash when one window subclassed multiple times
    '** .............. More END safe, even if END occurs within the subclass procedure
    '** ....... Added ability to destroy API windows when IDE terminates
    '** ....... Added auto-unsubclass when WM_NCDESTROY received
    '** NOTE: zAddressOf has been modified by fafalone to handle UserControls with >512 procedures
    '*****************************************************************************************
    ' Subclassing procedure must be declared identical to the one at the end of this class (Sample at Ordinal #1)

    Dim z_Sc(0 To IDX_UNICODE) As Long                 'Thunk machine-code initialised here
    
    Const SUB_NAME      As String = "ssc_Subclass"     'This routine's name
    Const CODE_LEN      As Long = 4 * IDX_UNICODE + 4  'Thunk length in bytes
    Const PAGE_RWX      As Long = &H40&                'Allocate executable memory
    Const MEM_COMMIT    As Long = &H1000&              'Commit allocated memory
    Const MEM_RELEASE   As Long = &H8000&              'Release allocated memory flag
    Const GWL_WNDPROC   As Long = -4                   'SetWindowsLong WndProc index
    Const WNDPROC_OFF   As Long = &H60                 'Thunk offset to the WndProc execution address
    Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1)) 'Bytes to allocate per thunk, data + code + msg tables
    
  ' This is the complete listing of thunk offset values and what they point/relate to.
  ' Those rem'd out are used elsewhere or are initialized in Declarations section
  
  'Const IDX_RECURSION  As Long = 0     'Thunk data index of callback recursion count
  'Const IDX_SHUTDOWN   As Long = 1     'Thunk data index of the termination flag
  'Const IDX_INDEX      As Long = 2     'Thunk data index of the subclassed hWnd
   Const IDX_EBMODE     As Long = 3     'Thunk data index of the EbMode function address
   Const IDX_CWP        As Long = 4     'Thunk data index of the CallWindowProc function address
   Const IDX_SWL        As Long = 5     'Thunk data index of the SetWindowsLong function address
   Const IDX_FREE       As Long = 6     'Thunk data index of the VirtualFree function address
   Const IDX_BADPTR     As Long = 7     'Thunk data index of the IsBadCodePtr function address
   Const IDX_OWNER      As Long = 8     'Thunk data index of the Owner object's vTable address
  'Const IDX_PREVPROC   As Long = 9     'Thunk data index of the original WndProc
   Const IDX_CALLBACK   As Long = 10    'Thunk data index of the callback method address
  'Const IDX_BTABLE     As Long = 11    'Thunk data index of the Before table
  'Const IDX_ATABLE     As Long = 12    'Thunk data index of the After table
  'Const IDX_PARM_USER  As Long = 13    'Thunk data index of the User-defined callback parameter data index
   Const IDX_DW         As Long = 14    'Thunk data index of the DestroyWinodw function address
   Const IDX_ST         As Long = 15    'Thunk data index of the SetTimer function address
   Const IDX_KT         As Long = 16    'Thunk data index of the KillTimer function address
   Const IDX_EBX_TMR    As Long = 20    'Thunk code patch index of the thunk data for the delay timer
   Const IDX_EBX        As Long = 26    'Thunk code patch index of the thunk data
  'Const IDX_UNICODE    As Long = xx    'Must be UBound(subclass thunkdata)+1; index for unicode support
    
    Dim z_ScMem       As Long           'Thunk base address
    Dim nAddr         As Long
    Dim nid           As Long
    Dim nMyID         As Long
    Dim bIDE          As Boolean

    If IsWindow(lng_hWnd) = 0 Then      'Ensure the window handle is valid
        Call zError(SUB_NAME, "Invalid window handle")
        Exit Function
    End If
    
    nMyID = GetCurrentProcessId                         'Get this process's ID
    GetWindowThreadProcessId lng_hWnd, nid              'Get the process ID associated with the window handle
    If nid <> nMyID Then                                'Ensure that the window handle doesn't belong to another process
        Call zError(SUB_NAME, "Window handle belongs to another process")
        Exit Function
    End If
    
    If oCallback Is Nothing Then Set oCallback = Me     'If the user hasn't specified the callback owner
    
    nAddr = zAddressOf(oCallback, nOrdinal)             'Get the address of the specified ordinal method
    If nAddr = 0 Then                                   'Ensure that we've found the ordinal method
        Call zError(SUB_NAME, "Callback method not found")
        Exit Function
    End If
        
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
    
    If z_ScMem <> 0 Then                                'Ensure the allocation succeeded
    
      If z_scFunk Is Nothing Then Set z_scFunk = New Collection 'If this is the first time through, do the one-time initialization
      On Error GoTo CatchDoubleSub                              'Catch double subclassing
      Call z_scFunk.Add(z_ScMem, "h" & lng_hWnd)                'Add the hWnd/thunk-address to the collection
      On Error GoTo 0
      
   'z_Sc (0) thru z_Sc(17) are used as storage for the thunks & IDX_ constants above relate to these thunk positions which are filled in below
    z_Sc(18) = &HD231C031: z_Sc(19) = &HBBE58960: z_Sc(21) = &H21E8F631: z_Sc(22) = &HE9000001: z_Sc(23) = &H12C&: z_Sc(24) = &HD231C031: z_Sc(25) = &HBBE58960: z_Sc(27) = &H3FFF631: z_Sc(28) = &H75047339: z_Sc(29) = &H2873FF23: z_Sc(30) = &H751C53FF: z_Sc(31) = &HC433913: z_Sc(32) = &H53FF2274: z_Sc(33) = &H13D0C: z_Sc(34) = &H18740000: z_Sc(35) = &H875C085: z_Sc(36) = &H820443C7: z_Sc(37) = &H90000000: z_Sc(38) = &H87E8&: z_Sc(39) = &H22E900: z_Sc(40) = &H90900000: z_Sc(41) = &H2C7B8B4A: z_Sc(42) = &HE81C7589: z_Sc(43) = &H90&: z_Sc(44) = &H75147539: z_Sc(45) = &H6AE80F: z_Sc(46) = &HD2310000: z_Sc(47) = &HE8307B8B: z_Sc(48) = &H7C&: z_Sc(49) = &H7D810BFF: z_Sc(50) = &H8228&: z_Sc(51) = &HC7097500: z_Sc(52) = &H80000443: z_Sc(53) = &H90900000: z_Sc(54) = &H44753339: z_Sc(55) = &H74047339: z_Sc(56) = &H2473FF3F: z_Sc(57) = &HFFFFFC68
    z_Sc(58) = &H2475FFFF: z_Sc(59) = &H811453FF: z_Sc(60) = &H82047B: z_Sc(61) = &HC750000: z_Sc(62) = &H74387339: z_Sc(63) = &H2475FF07: z_Sc(64) = &H903853FF: z_Sc(65) = &H81445B89: z_Sc(66) = &H484443: z_Sc(67) = &H73FF0000: z_Sc(68) = &H646844: z_Sc(69) = &H56560000: z_Sc(70) = &H893C53FF: z_Sc(71) = &H90904443: z_Sc(72) = &H10C261: z_Sc(73) = &H53E8&: z_Sc(74) = &H3075FF00: z_Sc(75) = &HFF2C75FF: z_Sc(76) = &H75FF2875: z_Sc(77) = &H2473FF24: z_Sc(78) = &H891053FF: z_Sc(79) = &H90C31C45: z_Sc(80) = &H34E30F8B: z_Sc(81) = &H1078C985: z_Sc(82) = &H4C781: z_Sc(83) = &H458B0000: z_Sc(84) = &H75AFF228: z_Sc(85) = &H90909023: z_Sc(86) = &H8D144D8D: z_Sc(87) = &H8D503443: z_Sc(88) = &H75FF1C45: z_Sc(89) = &H2C75FF30: z_Sc(90) = &HFF2875FF: z_Sc(91) = &H51502475: z_Sc(92) = &H2073FF52: z_Sc(93) = &H902853FF: z_Sc(94) = &H909090C3: z_Sc(95) = &H74447339: z_Sc(96) = &H4473FFF7
    z_Sc(97) = &H4053FF56: z_Sc(98) = &HC3447389: z_Sc(99) = &H89285D89: z_Sc(100) = &H45C72C75: z_Sc(101) = &H800030: z_Sc(102) = &H20458B00: z_Sc(103) = &H89145D89: z_Sc(104) = &H81612445: z_Sc(105) = &H4C4&: z_Sc(106) = &H1862FF00

    ' cache callback related pointers & offsets
      z_Sc(IDX_EBX) = z_ScMem                                                 'Patch the thunk data address
      z_Sc(IDX_EBX_TMR) = z_ScMem                                             'Patch the thunk data address
      z_Sc(IDX_INDEX) = lng_hWnd                                              'Store the window handle in the thunk data
      z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN                                   'Store the address of the before table in the thunk data
      z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4)         'Store the address of the after table in the thunk data
      z_Sc(IDX_OWNER) = ObjPtr(oCallback)                                     'Store the callback owner's object address in the thunk data
      z_Sc(IDX_CALLBACK) = nAddr                                              'Store the callback address in the thunk data
      z_Sc(IDX_PARM_USER) = lParamUser                                        'Store the lParamUser callback parameter in the thunk data
      
      ' validate unicode request & cache unicode usage
      If bUnicode Then bUnicode = (IsWindowUnicode(lng_hWnd) <> 0&)
      z_Sc(IDX_UNICODE) = bUnicode                                            'Store whether the window is using unicode calls or not
      
      ' get function pointers for the thunk
      If bIdeSafety = True Then                                               'If the user wants IDE protection
          Debug.Assert zInIDE(bIDE)
          If bIDE = True Then z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode", bUnicode) 'Store the EbMode function address in the thunk data
                                                        '^^ vb5 users, change vba6 to vba5
      End If
      If bIsAPIwindow Then                                                    'If user wants DestroyWindow sent should IDE end
          z_Sc(IDX_DW) = zFnAddr("user32", "DestroyWindow", bUnicode)
      End If
      z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree", bUnicode)           'Store the VirtualFree function address in the thunk data
      z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr", bUnicode)        'Store the IsBadCodePtr function address in the thunk data
      z_Sc(IDX_ST) = zFnAddr("user32", "SetTimer", bUnicode)                  'Store the SetTimer function address in the thunk data
      z_Sc(IDX_KT) = zFnAddr("user32", "KillTimer", bUnicode)                 'Store the KillTimer function address in the thunk data
      
      If bUnicode Then
          z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcW", bUnicode)      'Store CallWindowProc function address in the thunk data
          z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongW", bUnicode)       'Store the SetWindowLong function address in the thunk data
          RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                    'Copy the thunk code/data to the allocated memory
          z_Sc(IDX_PREVPROC) = SetWindowLongW(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF) 'Set the new WndProc, return the address of the original WndProc
      Else
          z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA", bUnicode)      'Store CallWindowProc function address in the thunk data
          z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA", bUnicode)       'Store the SetWindowLong function address in the thunk data
          RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                    'Copy the thunk code/data to the allocated memory
          z_Sc(IDX_PREVPROC) = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF) 'Set the new WndProc, return the address of the original WndProc
      End If
      If z_Sc(IDX_PREVPROC) = 0 Then                                          'Ensure the new WndProc was set correctly
          zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
          GoTo ReleaseMemory
      End If
      'Store the original WndProc address in the thunk data
      Call RtlMoveMemory(z_ScMem + IDX_PREVPROC * 4, VarPtr(z_Sc(IDX_PREVPROC)), 4&)
      ssc_Subclass = True                                                     'Indicate success
    Else
        Call zError(SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError)
    End If
 Exit Function                                                                'Exit ssc_Subclass
    
CatchDoubleSub:
 Call zError(SUB_NAME, "Window handle is already subclassed")
      
ReleaseMemory:
      Call VirtualFree(z_ScMem, 0, MEM_RELEASE)                               'ssc_Subclass has failed after memory allocation, so release the memory
End Function

'Terminate all subclassing
Private Sub ssc_Terminate()
    ' can be made public, can be removed & zTerminateThunks can be called instead
    Call zTerminateThunks(SubclassThunk)
End Sub

'UnSubclass the specified window handle
Private Sub ssc_UnSubclass(ByVal lng_hWnd As Long)
    ' can be made public, can be removed & zUnthunk can be called instead
    Call zUnThunk(lng_hWnd, SubclassThunk)
End Sub

'Add the message value to the window handle's specified callback table
Private Sub ssc_AddMsg(ByVal lng_hWnd As Long, ByVal When As ucst_eMsgWhen, ParamArray Messages() As Variant)
    Dim z_ScMem       As Long                                   'Thunk base address
    
    z_ScMem = zMap_VFunction(lng_hWnd, SubclassThunk)           'Ensure that the thunk hasn't already released its memory
    If z_ScMem Then
      Dim m As Long
      For m = LBound(Messages) To UBound(Messages)
        Select Case VarType(Messages(m))                        ' ensure no strings, arrays, doubles, objects, etc are passed
        Case vbByte, vbInteger, vbLong
            If When And MSG_BEFORE Then                         'If the message is to be added to the before original WndProc table...
              If zAddMsg(Messages(m), IDX_BTABLE, z_ScMem) = False Then 'Add the message to the before table
                When = (When And Not MSG_BEFORE)
              End If
            End If
            If When And MSG_AFTER Then                          'If message is to be added to the after original WndProc table...
              If zAddMsg(Messages(m), IDX_ATABLE, z_ScMem) = False Then 'Add the message to the after table
                When = (When And Not MSG_AFTER)
              End If
            End If
        End Select
      Next
    End If
End Sub

'Delete the message value from the window handle's specified callback table
Private Sub ssc_DelMsg(ByVal lng_hWnd As Long, ByVal When As ucst_eMsgWhen, ParamArray Messages() As Variant)
    Dim z_ScMem       As Long                           'Thunk base address
    z_ScMem = zMap_VFunction(lng_hWnd, SubclassThunk)   'Ensure that the thunk hasn't already released its memory
    If z_ScMem Then
      Dim m As Long
      For m = LBound(Messages) To UBound(Messages) ' ensure no strings, arrays, doubles, objects, etc are passed
        Select Case VarType(Messages(m))
        Case vbByte, vbInteger, vbLong
            If When And MSG_BEFORE Then            'If the message is to be removed from the before original WndProc table...
              Call zDelMsg(Messages(m), IDX_BTABLE, z_ScMem) 'Remove the message to the before table
            End If
            If When And MSG_AFTER Then                       'If message is to be removed from the after original WndProc table...
              zDelMsg Messages(m), IDX_ATABLE, z_ScMem       'Remove the message to the after table
            End If
        End Select
      Next
    End If
End Sub

'Call the original WndProc
Private Function ssc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ' can be made public, can be removed if you will not use this in your window procedure
    Dim z_ScMem       As Long                           'Thunk base address
    z_ScMem = zMap_VFunction(lng_hWnd, SubclassThunk)
    If z_ScMem Then                                     'Ensure that the thunk hasn't already released its memory
        If zData(IDX_UNICODE, z_ScMem) Then
            ssc_CallOrigWndProc = CallWindowProcW(zData(IDX_PREVPROC, z_ScMem), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
        Else
            ssc_CallOrigWndProc = CallWindowProcA(zData(IDX_PREVPROC, z_ScMem), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
        End If
    End If
End Function

'Get the subclasser lParamUser callback parameter
Private Function zGet_lParamUser(ByVal hWnd_Hook_ID As Long, ByVal vType As ucst_eThunkType) As Long
    ' can be removed if you never will retrieve or replace the user-defined parameter
    If vType <> CallbackThunk Then
        Dim z_ScMem       As Long                                       'Thunk base address
        z_ScMem = zMap_VFunction(hWnd_Hook_ID, vType)
        If z_ScMem Then                                                 'Ensure that the thunk hasn't already released its memory
          zGet_lParamUser = zData(IDX_PARM_USER, z_ScMem)               'Get the lParamUser callback parameter
        End If
    End If
End Function

'Let the subclasser lParamUser callback parameter
Private Sub zSet_lParamUser(ByVal hWnd_Hook_ID As Long, ByVal vType As ucst_eThunkType, ByVal newValue As Long)
    ' can be removed if you never will retrieve or replace the user-defined parameter
    If vType <> CallbackThunk Then
        Dim z_ScMem       As Long                                       'Thunk base address
        z_ScMem = zMap_VFunction(hWnd_Hook_ID, vType)
        If z_ScMem Then                                                 'Ensure that the thunk hasn't already released its memory
          zData(IDX_PARM_USER, z_ScMem) = newValue                      'Set the lParamUser callback parameter
        End If
    End If
End Sub

'Add the message to the specified table of the window handle
Private Function zAddMsg(ByVal uMsg As Long, ByVal nTable As Long, ByVal z_ScMem As Long) As Boolean
      Dim nCount As Long                            'Table entry count
      Dim nBase  As Long
      Dim i      As Long                            'Loop index
    
      zAddMsg = True
      nBase = zData(nTable, z_ScMem)                'Map zData() to the specified table
      
      If uMsg = ALL_MESSAGES Then                   'If ALL_MESSAGES are being added to the table...
        nCount = ALL_MESSAGES                       'Set the table entry count to ALL_MESSAGES
      Else
        
        nCount = zData(0, nBase)                    'Get the current table entry count
        For i = 1 To nCount                         'Loop through the table entries
          If zData(i, nBase) = 0 Then               'If the element is free...
            zData(i, nBase) = uMsg                  'Use this element
            GoTo Bail                               'Bail
          ElseIf zData(i, nBase) = uMsg Then        'If the message is already in the table...
            GoTo Bail                               'Bail
          End If
        Next i                                      'Next message table entry
    
        nCount = i                                  'On drop through: i = nCount + 1, the new table entry count
        If nCount > MSG_ENTRIES Then                'Check for message table overflow
          Call zError("zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values")
          zAddMsg = False
          GoTo Bail
        End If
        
        zData(nCount, nBase) = uMsg                                            'Store the message in the appended table entry
      End If
    
      zData(0, nBase) = nCount                                                 'Store the new table entry count
Bail:
End Function

'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long, ByVal z_ScMem As Long)
      Dim nCount As Long                                                        'Table entry count
      Dim nBase  As Long
      Dim i      As Long                                                        'Loop index
    
      nBase = zData(nTable, z_ScMem)                                            'Map zData() to the specified table
    
      If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being deleted from the table...
        zData(0, nBase) = 0                                                     'Zero the table entry count
      Else
        nCount = zData(0, nBase)                                                'Get the table entry count
        
        For i = 1 To nCount                                                     'Loop through the table entries
          If zData(i, nBase) = uMsg Then                                        'If the message is found...
            zData(i, nBase) = 0                                                 'Null the msg value -- also frees the element for re-use
            GoTo Bail                                                           'Bail
          End If
        Next i                                                                  'Next message table entry
       ' zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
      End If
Bail:
End Sub

'-SelfCallback code------------------------------------------------------------------------------------
'-The following routines are exclusively for the scb_SetCallbackAddr routines----------------------------
Private Function scb_SetCallbackAddr(ByVal nParamCount As Long, _
                     Optional ByVal nOrdinal As Long = 1, _
                     Optional ByVal oCallback As Object = Nothing, _
                     Optional ByVal bIdeSafety As Boolean = True, _
                     Optional ByVal bIsTimerCallback As Boolean) As Long   'Return the address of the specified callback thunk
    '*************************************************************************************************
    '* nParamCount  - The number of parameters that will callback
    '* nOrdinal     - Callback ordinal number, the final private method is ordinal 1, the second last is ordinal 2, etc...
    '* oCallback    - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
    '* bIdeSafety   - Optional, set to false to disable IDE protection.
    '* bIsTimerCallback - optional, set to true for extra protection when used as a SetTimer callback
    '       If True, timer will be destroyed when IDE/app terminates. See scb_ReleaseCallback.
    '*************************************************************************************************
    ' Callback procedure must return a Long even if, per MSDN, the callback procedure is a Sub vs Function
    ' The number of parameters and their types are dependent on the individual callback procedures
    
    Const MEM_LEN     As Long = IDX_CALLBACKORDINAL * 4 + 4     'Memory bytes required for the callback thunk
    Const PAGE_RWX    As Long = &H40&                           'Allocate executable memory
    Const MEM_COMMIT  As Long = &H1000&                         'Commit allocated memory
    Const SUB_NAME      As String = "scb_SetCallbackAddr"       'This routine's name
    Const INDX_OWNER    As Long = 0                             'Thunk data index of the Owner object's vTable address
    Const INDX_CALLBACK As Long = 1                             'Thunk data index of the EbMode function address
    Const INDX_EBMODE   As Long = 2                             'Thunk data index of the IsBadCodePtr function address
    Const INDX_BADPTR   As Long = 3                             'Thunk data index of the IsBadCodePtr function address
    Const INDX_KT       As Long = 4                             'Thunk data index of the KillTimer function address
    Const INDX_EBX      As Long = 6                             'Thunk code patch index of the thunk data
    Const INDX_PARAMS   As Long = 18                            'Thunk code patch index of the number of parameters expected in callback
    Const INDX_PARAMLEN As Long = 24                            'Thunk code patch index of the bytes to be released after callback
    Const PROC_OFF      As Long = &H14                          'Thunk offset to the callback execution address

    Dim z_ScMem       As Long                                   'Thunk base address
    Dim z_Cb()    As Long                                       'Callback thunk array
    Dim nValue    As Long
    Dim nCallback As Long
    Dim bIDE      As Boolean
      
    If oCallback Is Nothing Then Set oCallback = Me     'If the user hasn't specified the callback owner
    If z_cbFunk Is Nothing Then
        Set z_cbFunk = New Collection           'If this is the first time through, do the one-time initialization
    Else
        On Error Resume Next                    'Catch already initialized?
        z_ScMem = z_cbFunk.Item("h" & ObjPtr(oCallback) & "." & nOrdinal) 'Test it
        If Err = 0 Then
            scb_SetCallbackAddr = z_ScMem + PROC_OFF  'we had this one, just reference it
            Exit Function
        End If
        On Error GoTo 0
    End If
    
    If nParamCount < 0 Then                     ' validate parameters
        Call zError(SUB_NAME, "Invalid Parameter count")
        Exit Function
    End If
    If oCallback Is Nothing Then
        Set oCallback = Me
    End If
    nCallback = zAddressOf(oCallback, nOrdinal)         'Get the callback address of the specified ordinal
    If nCallback = 0 Then
        Call zError(SUB_NAME, "Callback address not found.")
        Exit Function
    End If
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
        
    If z_ScMem = 0& Then
        Call zError(SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError)  ' oops
        Exit Function
    End If
    Call z_cbFunk.Add(z_ScMem, "h" & ObjPtr(oCallback) & "." & nOrdinal) 'Add the callback/thunk-address to the collection
        
    ReDim z_Cb(0 To IDX_CALLBACKORDINAL) As Long          'Allocate for the machine-code array
    
    ' Create machine-code array
    z_Cb(5) = &HBB60E089: z_Cb(7) = &H73FFC589: z_Cb(8) = &HC53FF04: z_Cb(9) = &H59E80A74: z_Cb(10) = &HE9000000
    z_Cb(11) = &H30&: z_Cb(12) = &H87B81: z_Cb(13) = &H75000000: z_Cb(14) = &H9090902B: z_Cb(15) = &H42DE889: z_Cb(16) = &H50000000: z_Cb(17) = &HB9909090: z_Cb(19) = &H90900AE3
    z_Cb(20) = &H8D74FF: z_Cb(21) = &H9090FAE2: z_Cb(22) = &H53FF33FF: z_Cb(23) = &H90909004: z_Cb(24) = &H2BADC261: z_Cb(25) = &H3D0853FF: z_Cb(26) = &H1&: z_Cb(27) = &H23DCE74: z_Cb(28) = &H74000000: z_Cb(29) = &HAE807
    z_Cb(30) = &H90900000: z_Cb(31) = &H4589C031: z_Cb(32) = &H90DDEBFC: z_Cb(33) = &HFF0C75FF: z_Cb(34) = &H53FF0475: z_Cb(35) = &HC310&

    z_Cb(INDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr", False)
    z_Cb(INDX_OWNER) = ObjPtr(oCallback)                    'Set the Owner
    z_Cb(INDX_CALLBACK) = nCallback                         'Set the callback address
    z_Cb(IDX_CALLBACKORDINAL) = nOrdinal                    'Cache ordinal used for zTerminateThunks
      
    If bIdeSafety = True Then                               'If the user wants IDE protection
        Debug.Assert zInIDE(bIDE)
        If bIDE = True Then z_Cb(INDX_EBMODE) = zFnAddr("vba6", "EbMode", False) 'Store the EbMode function address in the thunk data
    End If
    If bIsTimerCallback Then
        z_Cb(INDX_KT) = zFnAddr("user32", "KillTimer", False)
    End If
        
    z_Cb(INDX_PARAMS) = nParamCount                         'Set the parameter count
    Call RtlMoveMemory(VarPtr(z_Cb(INDX_PARAMLEN)) + 2, VarPtr(nParamCount * 4), 2&)

    z_Cb(INDX_EBX) = z_ScMem                                'Set the data address relative to virtual memory pointer

    Call RtlMoveMemory(z_ScMem, VarPtr(z_Cb(INDX_OWNER)), MEM_LEN) 'Copy thunk code to executable memory
    scb_SetCallbackAddr = z_ScMem + PROC_OFF                       'Thunk code start address
End Function

Private Sub scb_ReleaseCallback(ByVal nOrdinal As Long, Optional ByVal oCallback As Object)
    ' can be made public, can be removed & zUnThunk can be called instead
    ' NEVER call this from within the callback routine itself
    
    ' oCallBack is the object containing nOrdinal to be released
    ' if oCallback was already closed (say it was a class or form), then you won't be
    '   able to release it here, but it will be released when zTerminateThunks is
    '   eventually called
    
    ' Special Warning. If the callback thunk is used for a recurring callback (i.e., Timer),
    ' then ensure you terminate what is using the callback before releasing the thunk,
    ' otherwise you are subject to a crash when that item tries to callback to zeroed memory
    Call zUnThunk(nOrdinal, CallbackThunk, oCallback)
End Sub

Private Sub scb_TerminateCallbacks()
    ' can be made public, can be removed & zTerminateThunks can be called instead
    Call zTerminateThunks(CallbackThunk)
End Sub

'-The following routines are used for each of the three types of thunks ----------------------------

'Maps zData() to the memory address for the specified thunk type
Private Function zMap_VFunction(vFuncTarget As Long, _
                                vType As ucst_eThunkType, _
                                Optional oCallback As Object, _
                                Optional bIgnoreErrors As Boolean) As Long
    
    Dim thunkCol As Collection
    Dim colID As String
    Dim z_ScMem       As Long         'Thunk base address
    
    If vType = CallbackThunk Then
        Set thunkCol = z_cbFunk
        If oCallback Is Nothing Then Set oCallback = Me
        colID = "h" & ObjPtr(oCallback) & "." & vFuncTarget
    ElseIf vType = SubclassThunk Then
        Set thunkCol = z_scFunk
        colID = "h" & vFuncTarget
    Else
        Call zError("zMap_Vfunction", "Invalid thunk type passed")
        Exit Function
    End If
    
    If thunkCol Is Nothing Then
        Call zError("zMap_VFunction", "Thunk hasn't been initialized")
    Else
        If thunkCol.Count Then
            On Error GoTo Catch
            z_ScMem = thunkCol(colID)               'Get the thunk address
            If IsBadCodePtr(z_ScMem) Then z_ScMem = 0&
            zMap_VFunction = z_ScMem
        End If
    End If
    Exit Function                                   'Exit returning the thunk address
Catch:
    ' error ignored when zUnThunk is called, error handled there
    If Not bIgnoreErrors Then Call zError("zMap_VFunction", "Thunk type for " & vType & " does not exist")
End Function

' sets/retrieves data at the specified offset for the specified memory address
Private Property Get zData(ByVal nIndex As Long, ByVal z_ScMem As Long) As Long
  Call RtlMoveMemory(VarPtr(zData), z_ScMem + (nIndex * 4), 4)
End Property

Private Property Let zData(ByVal nIndex As Long, ByVal z_ScMem As Long, ByVal nValue As Long)
  Call RtlMoveMemory(z_ScMem + (nIndex * 4), VarPtr(nValue), 4)
End Property

'Error handler
Private Sub zError(ByRef sRoutine As String, ByVal sMsg As String)
  ' Note. These two lines can be rem'd out if you so desire. But don't remove the routine
  ' App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
  Call MsgBox(sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine)
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String, ByVal asUnicode As Boolean) As Long
  If asUnicode Then
    zFnAddr = GetProcAddress(GetModuleHandleW(StrPtr(sDLL)), sProc)         'Get the specified procedure address
  Else
    zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                 'Get the specified procedure address
  End If
  Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
  ' ^^ FYI VB5 users. Search for zFnAddr("vba6", "EbMode") and replace with zFnAddr("vba5", "EbMode")
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
    ' Note: used both in subclassing and hooking routines
  Dim bSub  As Byte                                     'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                     'Address of the vTable
  Dim i     As Long                                     'Loop index
  Dim j     As Long                                     'Loop limit
  
  Call RtlMoveMemory(VarPtr(nAddr), ObjPtr(oCallback), 4) 'Get the address of the callback object's instance
  If Not zProbe(nAddr + &H1C, i, bSub) Then             'Probe for a Class method
    If Not zProbe(nAddr + &H6F8, i, bSub) Then          'Probe for a Form method
      If Not zProbe(nAddr + &H710, i, bSub) Then        'Probe for a PropertyPage method
        If Not zProbe(nAddr + &H7A4, i, bSub) Then      'Probe for a UserControl method
            Exit Function                               'Bail...
        End If
      End If
    End If
  End If
  
  i = i + 4                                             'Bump to the next entry
'  J = i + 2048                                          'Set a reasonable limit, scan 512 vTable entries
  j = i + 4096                                          'NOTE: Modified by fafalone for ultra-large UCs (>512 methods)
  Do While i < j
    Call RtlMoveMemory(VarPtr(nAddr), i, 4)             'Get the address stored in this vTable entry
    
    If IsBadCodePtr(nAddr) Then                                     'Is the entry an invalid code address?
      Call RtlMoveMemory(VarPtr(zAddressOf), i - (nOrdinal * 4), 4) 'Return the specified vTable entry address
      Exit Do                                                       'Bad method signature, quit loop
    End If

    Call RtlMoveMemory(VarPtr(bVal), nAddr, 1)                      'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                            'If the byte doesn't match the expected value...
      Call RtlMoveMemory(VarPtr(zAddressOf), i - (nOrdinal * 4), 4) 'Return the specified vTable entry address
      Exit Do                                                       'Bad method signature, quit loop
    End If
    
    i = i + 4                                                       'Next vTable entry
  Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                    'Start address
  nLimit = nAddr + 32                                               'Probe eight entries
  Do While nAddr < nLimit                                           'While we've not reached our probe depth
    Call RtlMoveMemory(VarPtr(nEntry), nAddr, 4)                    'Get the vTable entry
    
    If nEntry <> 0 Then                                             'If not an implemented interface
      Call RtlMoveMemory(VarPtr(bVal), nEntry, 1)                   'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                            'Check for a native or pcode method signature
        nMethod = nAddr                                             'Store the vTable entry
        bSub = bVal                                                 'Store the found method signature
        zProbe = True                                               'Indicate success
        Exit Do                                                     'Return
      End If
    End If
    nAddr = nAddr + 4                                               'Next vTable entry
  Loop
End Function

Private Function zInIDE(ByRef bIDE As Boolean) As Boolean
    ' only called in IDE, never called when compiled
    bIDE = True
    zInIDE = bIDE
End Function

Private Sub zUnThunk(ByVal thunkID As Long, ByVal vType As ucst_eThunkType, Optional ByVal oCallback As Object)
    ' thunkID, depends on vType:
    '   - Subclassing:  the hWnd of the window subclassed
    '   - Callbacks:    the ordinal of the callback
    '       ensure KillTimer is already called, if any callback used for SetTimer
    ' oCallback only used when vType is CallbackThunk
    Const IDX_SHUTDOWN  As Long = 1
    Const MEM_RELEASE As Long = &H8000&             'Release allocated memory flag
    
    Dim z_ScMem       As Long                       'Thunk base address
    
    z_ScMem = zMap_VFunction(thunkID, vType, oCallback, True)
    Select Case vType
    Case SubclassThunk
        If z_ScMem Then                         'Ensure that the thunk hasn't already released its memory
            zData(IDX_SHUTDOWN, z_ScMem) = 1                  'Set the shutdown indicator
            Call zDelMsg(ALL_MESSAGES, IDX_BTABLE, z_ScMem)   'Delete all before messages
            Call zDelMsg(ALL_MESSAGES, IDX_ATABLE, z_ScMem)   'Delete all after messages
        End If
        Call z_scFunk.Remove("h" & thunkID)                   'Remove the specified thunk from the collection
    Case CallbackThunk
        If z_ScMem Then                         'Ensure that the thunk hasn't already released its memory
            Call VirtualFree(z_ScMem, 0, MEM_RELEASE)   'Release allocated memory
        End If
        Call z_cbFunk.Remove("h" & ObjPtr(oCallback) & "." & thunkID) 'Remove the specified thunk from the collection
    End Select
End Sub

Private Sub zTerminateThunks(ByVal vType As ucst_eThunkType)
    ' Terminates all thunks of a specific type
    ' Any subclassing, recurring callbacks should have already been canceled
    Dim i As Long
    Dim oCallback As Object
    Dim thunkCol As Collection
    Dim z_ScMem       As Long                           'Thunk base address
    Const INDX_OWNER As Long = 0
    
    Select Case vType
    Case SubclassThunk
        Set thunkCol = z_scFunk
    Case CallbackThunk
        Set thunkCol = z_cbFunk
    Case Else
        Exit Sub
    End Select
    
    If Not (thunkCol Is Nothing) Then                 'Ensure that hooking has been started
      With thunkCol
        For i = .Count To 1 Step -1                   'Loop through the collection of hook types in reverse order
          z_ScMem = .Item(i)                          'Get the thunk address
          If IsBadCodePtr(z_ScMem) = 0 Then           'Ensure that the thunk hasn't already released its memory
            Select Case vType
                Case SubclassThunk
                    zUnThunk zData(IDX_INDEX, z_ScMem), SubclassThunk    'Unsubclass
                Case CallbackThunk
                    ' zUnThunk expects object not pointer, convert pointer to object
                    Call RtlMoveMemory(VarPtr(oCallback), VarPtr(zData(INDX_OWNER, z_ScMem)), 4&)
                    Call zUnThunk(zData(IDX_CALLBACKORDINAL, z_ScMem), CallbackThunk, oCallback) ' release callback
                    ' remove the object pointer reference
                    Call RtlMoveMemory(VarPtr(oCallback), VarPtr(INDX_OWNER), 4&)
            End Select
          End If
        Next i                                        'Next member of the collection
      End With
      Set thunkCol = Nothing                         'Destroy the hook/thunk-address collection
    End If
End Sub
'==============================================================================
'End of Self-Subclass procedures
'==============================================================================

' ===================================================================
' treeview macros

' Sets the normal or state image list for a tree-view control and redraws the control using the new images.
' Returns the handle to the previous image list, if any, or 0 otherwise.

Private Function TreeView_SetImageList(hWnd As Long, himl As Long, iImage As ucst_TVImageLists) As Long
  TreeView_SetImageList = SendMessage(hWnd, TVM_SETIMAGELIST, ByVal iImage, ByVal himl)
End Function

Private Function TreeView_GetImageList(hWnd As Long) As Long
TreeView_GetImageList = SendMessage(hWnd, TVM_GETIMAGELIST, 0, ByVal 0&)
End Function

Private Function TreeView_GetIndent(hWnd As Long) As Long
TreeView_GetIndent = SendMessage(hWnd, TVM_GETINDENT, 0, ByVal 0&)
End Function

Private Function TreeView_SetIndent(hWnd As Long, indent As Long) As Long
TreeView_SetIndent = SendMessage(hWnd, TVM_SETINDENT, indent, ByVal 0&)
End Function

Private Function TreeView_GetISearchString(hWnd As Long, lpsz As Long) As Long
TreeView_GetISearchString = SendMessage(hWnd, TVM_GETISEARCHSTRING, 0, ByVal lpsz)
End Function

Private Function TreeView_SetToolTips(hWnd As Long, hwndTT As Long) As Long
TreeView_SetToolTips = SendMessage(hWnd, TVM_SETTOOLTIPS, hwndTT, ByVal 0&)
End Function

Private Function TreeView_GetToolTips(hWnd As Long) As Long
TreeView_GetToolTips = SendMessage(hWnd, TVM_GETTOOLTIPS, 0, ByVal 0&)
End Function

Private Function TreeView_GetItemPartRect(hWnd As Long, hItem As Long, prc As oleexp.RECT, partid As ucst_TVITEMPART) As Long
Dim tInfo As TVGETITEMPARTRECTINFO
tInfo.hti = hItem
tInfo.prc = prc
tInfo.partid = partid
TreeView_GetItemPartRect = SendMessage(hWnd, TVM_GETITEMPARTRECT, 0, tInfo)

End Function

Private Function TreeView_GetItemRect(hWnd As Long, hItem As Long, prc As oleexp.RECT, Code As Long) As Long
TreeView_GetItemRect = SendMessage(hWnd, TVM_GETITEMRECT, Code, prc)
End Function

Private Function TreeView_GetLineColor(hWnd As Long) As Long
    TreeView_GetLineColor = SendMessage(hWnd, TVM_GETLINECOLOR, 0, ByVal 0&)
End Function

Private Function TreeView_SetLineColor(hWnd As Long, clr As Long) As Long
    TreeView_SetLineColor = SendMessage(hWnd, TVM_SETLINECOLOR, 0, ByVal clr)
End Function

Private Function TreeView_SetStyle(hWnd As Long, dwStyle As ucst_TV_Styles) As Long
    Dim lStyle As Long
    lStyle = GetWindowLong(hWnd, GWL_STYLE)
    lStyle = lStyle Or dwStyle
    TreeView_SetStyle = SetWindowLong(hWnd, GWL_STYLE, lStyle)

End Function

Private Function TreeView_SetExtendedStyle(hWnd As Long, lST As ucst_TV_Ex_Styles) As Long
    Dim lStyle As Long

    lStyle = SendMessage(hWnd, TVM_GETEXTENDEDSTYLE, 0, 0)
    lStyle = lStyle Or lST
    Call SendMessage(hWnd, TVM_SETEXTENDEDSTYLE, 0, ByVal lStyle)

End Function

' TreeView_GetNextItem

' Retrieves the tree-view item that bears the specified relationship to a specified item.
' Returns the handle to the item if successful or 0 otherwise.

Private Function TreeView_GetNextItem(hWnd As Long, _
                                     hItem As Long, _
                                     flag As ucst_TVM_GETNEXTITEM_wParam) As Long
    TreeView_GetNextItem = SendMessage(hWnd, TVM_GETNEXTITEM, ByVal flag, ByVal hItem)
End Function

Private Function TreeView_GetNextSelected(hWnd As Long, hItem As Long)
    TreeView_GetNextSelected = TreeView_GetNextItem(hWnd, hItem, TVGN_NEXTSELECTED)
End Function
' Retrieves the first child item. The hitem parameter must be NULL.
' Returns the handle to the item if successful or 0 otherwise.

Private Function TreeView_GetChild(hWnd As Long, hItem As Long) As Long
    TreeView_GetChild = TreeView_GetNextItem(hWnd, hItem, TVGN_CHILD)
End Function

' Retrieves the next sibling item.
' Returns the handle to the item if successful or 0 otherwise.

Private Function TreeView_GetNextSibling(hWnd As Long, hItem As Long) As Long
    TreeView_GetNextSibling = TreeView_GetNextItem(hWnd, hItem, TVGN_NEXT)
End Function

Private Function TreeView_GetPrevSibling(hWnd As Long, hItem As Long)
  TreeView_GetPrevSibling = TreeView_GetNextItem(hWnd, hItem, TVGN_PREVIOUS)
End Function

Private Function TreeView_GetNextVisible(hWnd As Long, hItem As Long)
  TreeView_GetNextVisible = TreeView_GetNextItem(hWnd, hItem, TVGN_NEXTVISIBLE)
End Function

Private Function TreeView_GetPrevVisible(hWnd As Long, hItem As Long)
  TreeView_GetPrevVisible = TreeView_GetNextItem(hWnd, hItem, TVGN_PREVIOUSVISIBLE)
End Function

' Retrieves the parent of the specified item.
' Returns the handle to the item if successful or 0 otherwise.
Private Function TreeView_SelectDropTarget(hWnd As Long, hItem As Long)
  TreeView_SelectDropTarget = TreeView_Select(hWnd, hItem, TVGN_DROPHILITE)
End Function

Private Function TreeView_SelectSetFirstVisible(hWnd As Long, hItem As Long)
  TreeView_SelectSetFirstVisible = TreeView_Select(hWnd, hItem, TVGN_FIRSTVISIBLE)
End Function

Private Function TreeView_SetAutoScrollInfo(hWnd As Long, uPixPerSec As Long, uUpdateTime As Long) As Long
TreeView_SetAutoScrollInfo = SendMessage(hWnd, TVM_SETAUTOSCROLLINFO, uPixPerSec, ByVal uUpdateTime)
End Function

'#define TreeView_SetBorder(hwnd, dwFlags, xBorder, yBorder) (int)SNDMSG ((hwnd), TVM_SETBORDER,(WPARAM) (dwFlags),
'MAKELPARAM (xBorder, yBorder))
Private Function TreeView_SetBorder(hWnd As Long, dwFlags As ucst_TVSB_Flags, xBorder As Long, yBorder As Long) As Long
TreeView_SetBorder = SendMessage(hWnd, TVM_SETBORDER, dwFlags, ByVal MAKELPARAM(xBorder, yBorder))
End Function

Private Function TreeView_SetHot(hWnd As Long, hItem As Long) As Long
TreeView_SetHot = SendMessage(hWnd, TVM_SETHOT, 0, ByVal hItem)
End Function

Private Function TreeView_SetTextColor(hWnd As Long, clr As Long) As Long
TreeView_SetTextColor = SendMessage(hWnd, TVM_SETTEXTCOLOR, 0, ByVal clr)
End Function

Private Function TreeView_ShowInfoTip(hWnd As Long, hItem As Long) As Long
TreeView_ShowInfoTip = SendMessage(hWnd, TVM_SHOWINFOTIP, 0, ByVal hItem)
End Function

Private Function TreeView_GetParent(hWnd As Long, hItem As Long) As Long
    TreeView_GetParent = TreeView_GetNextItem(hWnd, hItem, TVGN_PARENT)
End Function

Private Function TreeView_MapAccIDToHTREEITEM(hWnd As Long, id As Long) As Long
TreeView_MapAccIDToHTREEITEM = SendMessage(hWnd, TVM_MAPACCIDTOHTREEITEM, id, ByVal 0&)
End Function

Private Function TreeView_MapHTREEITEMToAccID(hWnd As Long, htreeitem As Long) As Long
TreeView_MapHTREEITEMToAccID = SendMessage(hWnd, TVM_MAPHTREEITEMTOACCID, htreeitem, ByVal 0&)
End Function

' Retrieves the currently selected item.
' Returns the handle to the item if successful or 0 otherwise.

Private Function TreeView_GetSelection(hWnd As Long) As Long
    TreeView_GetSelection = TreeView_GetNextItem(hWnd, 0, TVGN_CARET)
End Function

Private Function TreeView_InsertItem(hWnd As Long, lpis As TVINSERTSTRUCT) As Long
TreeView_InsertItem = SendMessage(hWnd, TVM_INSERTITEM, 0, lpis)
End Function

Private Function TreeView_InsertItemEx(hWnd As Long, lpis As TVINSERTSTRUCTEX) As Long
TreeView_InsertItemEx = SendMessage(hWnd, TVM_INSERTITEM, 0, lpis)
End Function
' Retrieves the topmost or very first item of the tree-view control.
' Returns the handle to the item if successful or 0 otherwise.

Private Function TreeView_GetRoot(hWnd As Long) As Long
    TreeView_GetRoot = TreeView_GetNextItem(hWnd, 0, TVGN_ROOT)
End Function

' Retrieves some or all of a tree-view item's attributes.
' Returns TRUE if successful or FALSE otherwise.

Private Function TreeView_GetItem(hWnd As Long, pItem As TVITEM) As Boolean
    TreeView_GetItem = SendMessage(hWnd, TVM_GETITEM, 0, pItem)
End Function

Private Function TreeView_GetItemState(hwndTV As Long, hti As Long, Mask As ucst_TVITEM_State) As Long
    TreeView_GetItemState = SendMessage(hwndTV, TVM_GETITEMSTATE, hti, ByVal Mask)
End Function
' Sets some or all of a tree-view item's attributes.
' Old docs say returns zero if successful or - 1 otherwise.
' New docs say returns TRUE if successful, or FALSE otherwise

Private Function TreeView_SetItem(hWnd As Long, pItem As TVITEM) As Boolean
    TreeView_SetItem = SendMessage(hWnd, TVM_SETITEM, 0, pItem)
End Function

Private Function TreeView_SetItemState(hwndTV As Long, _
                                      hti As Long, _
                                      data As ucst_TVITEM_State, _
                                      Mask As ucst_TVITEM_State) As Long
    Dim tVI As TVITEMEX
    tVI.Mask = TVIF_STATE Or TVIF_HANDLE
    tVI.hItem = hti
    tVI.StateMask = Mask
    tVI.State = data
    TreeView_SetItemState = SendMessage(hwndTV, TVM_SETITEM, 0, tVI)
End Function

Private Function TreeView_SetItemStateEx(hwndTV As Long, _
                                        hti As Long, _
                                        Mask As ucst_TVITEM_State_Ex) As Long
    Dim tVI As TVITEMEX
    tVI.Mask = TVIF_STATEEX
    tVI.hItem = hti
    tVI.uStateEx = Mask
    TreeView_SetItemStateEx = SendMessage(hwndTV, TVM_SETITEM, 0, tVI)
End Function

Private Function TreeView_SetCheckState(hwndTV As Long, _
                                       hti As Long, _
                                       fCheck As Long) As Long
    TreeView_SetCheckState = TreeView_SetItemState(hwndTV, hti, IndexToStateImageMask(IIf(fCheck, 2, 1)), TVIS_STATEIMAGEMASK)

End Function

Private Function TreeView_SetCheckStateEx(hwndTV As Long, _
                                       hti As Long, _
                                       fCheck As Long) As Long
    TreeView_SetCheckStateEx = TreeView_SetItemState(hwndTV, hti, IndexToStateImageMask(fCheck), TVIS_STATEIMAGEMASK)

End Function

Private Function TreeView_GetCheckState(hwndTV As Long, hti As Long) As Long
    TreeView_GetCheckState = StateImageMaskToIndex(SendMessage(hwndTV, TVM_GETITEMSTATE, hti, ByVal TVIS_STATEIMAGEMASK))

End Function
' Determines the location of the specified point relative to the client area of a tree-view control.
' Returns the handle to the tree-view item that occupies the specified point or NULL if no item
' occupies the point.

Private Function TreeView_HitTest(hWnd As Long, lpht As TVHITTESTINFO) As Long
    TreeView_HitTest = SendMessage(hWnd, TVM_HITTEST, 0, lpht)
End Function
' Removes an item from a tree-view control.
' Returns TRUE if successful or FALSE otherwise.

Private Function TreeView_DeleteItem(hWnd As Long, hItem As Long) As Boolean
    TreeView_DeleteItem = SendMessage(hWnd, TVM_DELETEITEM, 0, ByVal hItem)
End Function

' Removes all items from a tree-view control.
' Returns TRUE if successful or FALSE otherwise.

Private Function TreeView_DeleteAllItems(hWnd As Long) As Boolean
    TreeView_DeleteAllItems = SendMessage(hWnd, TVM_DELETEITEM, 0, ByVal TVI_ROOT)
End Function
' Creates a dragging bitmap for the specified item in a tree-view control, creates an image list
' for the bitmap, and adds the bitmap to the image list. An application can display the image
' when dragging the item by using the image list functions.
' Returns the handle of the image list to which the dragging bitmap was added if successful or
' NULL otherwise.

Private Function TreeView_CreateDragImage(hWnd As Long, hItem As Long) As Long
    TreeView_CreateDragImage = SendMessage(hWnd, TVM_CREATEDRAGIMAGE, 0, ByVal hItem)
End Function

' Sorts the child items of the specified parent item in a tree-view control.
' Returns TRUE if successful or FALSE otherwise.
' fRecurse is reserved for future use and must be zero.
Private Function TreeView_EditLabel(hWnd As Long, hItem As Long) As Long
    TreeView_EditLabel = SendMessage(hWnd, TVM_EDITLABEL, ByVal 0&, ByVal hItem)
End Function

Private Function TreeView_EndEditLabelNow(hWnd As Long, fCancel As Long) As Long
    TreeView_EndEditLabelNow = SendMessage(hWnd, TVM_ENDEDITLABELNOW, fCancel, ByVal 0&)
End Function

Private Function TreeView_GetFirstVisible(hWnd As Long) As Long
    TreeView_GetFirstVisible = TreeView_GetNextItem(hWnd, 0, TVGN_FIRSTVISIBLE)
End Function

Private Function TreeView_GetLastVisible(hWnd As Long) As Long   ' IE4
    TreeView_GetLastVisible = TreeView_GetNextItem(hWnd, 0, TVGN_LASTVISIBLE)
End Function

Private Function TreeView_GetBkColor(hWnd As Long) As Long
    TreeView_GetBkColor = SendMessage(hWnd, TVM_GETBKCOLOR, 0, ByVal 0&)
End Function

Private Function TreeView_SetInsertMark(hWnd As Long, _
                                       hItem As Long, _
                                       fAfter As BOOL) As Boolean   ' IE4
    TreeView_SetInsertMark = SendMessage(hWnd, TVM_SETINSERTMARK, ByVal fAfter, ByVal hItem)
End Function

Private Function TreeView_GetCount(hWnd As Long) As Long
    TreeView_GetCount = SendMessage(hWnd, TVM_GETCOUNT, 0, ByVal 0&)
End Function

Private Function TreeView_GetDropHilight(hWnd As Long) As Long
    TreeView_GetDropHilight = TreeView_GetNextItem(hWnd, 0, TVGN_DROPHILITE)
End Function

Private Function TreeView_GetExtendedStyle(hWnd As Long) As Long
    TreeView_GetExtendedStyle = SendMessage(hWnd, TVM_GETEXTENDEDSTYLE, 0, ByVal 0&)
End Function

Private Function TreeView_SetUnicodeFormat(hWnd As Long, _
                                          fUnicode As BOOL) As Boolean   ' IE4
    TreeView_SetUnicodeFormat = SendMessage(hWnd, TVM_SETUNICODEFORMAT, ByVal fUnicode, 0)
End Function

Private Function TreeView_GetUnicodeFormat(hWnd As Long) As Boolean   ' IE4
    TreeView_GetUnicodeFormat = SendMessage(hWnd, TVM_GETUNICODEFORMAT, 0, 0)
End Function

' returns (int), old?
Private Function TreeView_SetItemHeight(hWnd As Long, iHeight As Long) As Long   ' IE4
    TreeView_SetItemHeight = SendMessage(hWnd, TVM_SETITEMHEIGHT, ByVal iHeight, 0)
End Function

Private Function TreeView_GetTextColor(hWnd As Long) As Long   ' IE4
    TreeView_GetTextColor = SendMessage(hWnd, TVM_GETTEXTCOLOR, 0, 0)
End Function

' returns (UINT), old?
Private Function TreeView_SetScrollTime(hWnd As Long, uTime As Long) As Long   ' IE4
    TreeView_SetScrollTime = SendMessage(hWnd, TVM_SETSCROLLTIME, ByVal uTime, 0)
End Function

' returns (UINT)
Private Function TreeView_GetScrollTime(hWnd As Long) As Long   ' IE4
    TreeView_GetScrollTime = SendMessage(hWnd, TVM_GETSCROLLTIME, 0, 0)
End Function

' returns (COLORREF), old?
Private Function TreeView_SetInsertMarkColor(hWnd As Long, clr As Long) As Long   ' IE4
    TreeView_SetInsertMarkColor = SendMessage(hWnd, TVM_SETINSERTMARKCOLOR, 0, ByVal clr)
End Function

' returns (COLORREF)
Private Function TreeView_GetInsertMarkColor(hWnd As Long) As Long   ' IE4
    TreeView_GetInsertMarkColor = SendMessage(hWnd, TVM_GETINSERTMARKCOLOR, 0, 0)
End Function
'
' ================================================================

' returns (int)
Private Function TreeView_GetItemHeight(hWnd As Long) As Long   ' IE4
    TreeView_GetItemHeight = SendMessage(hWnd, TVM_GETITEMHEIGHT, 0, 0)
End Function

' returns (COLORREF), old?
Private Function TreeView_SetBkColor(hWnd As Long, clr As Long) As Long   ' IE4
    TreeView_SetBkColor = SendMessage(hWnd, TVM_SETBKCOLOR, 0, ByVal clr)
End Function

' Retrieves the handle to the edit control being used to edit a tree-view item's text.
' Returns the handle to the edit control if successful or NULL otherwise.

Private Function TreeView_GetEditControl(hWnd As Long) As Long
    TreeView_GetEditControl = SendMessage(hWnd, TVM_GETEDITCONTROL, 0, 0)
End Function

' Returns the number of items that are fully visible in the client window of the tree-view control.

Private Function TreeView_GetVisibleCount(hWnd As Long) As Long
    TreeView_GetVisibleCount = SendMessage(hWnd, TVM_GETVISIBLECOUNT, 0, 0)
End Function

Private Function TreeView_SortChildren(hWnd As Long, _
                                      hItem As Long, _
                                      fRecurse As Boolean) As Boolean
    TreeView_SortChildren = SendMessage(hWnd, TVM_SORTCHILDREN, ByVal fRecurse, ByVal hItem)
End Function

' Ensures that a tree-view item is visible, expanding the parent item or scrolling the tree-view
' control, if necessary.
' Returns TRUE if the system scrolled the items in the tree-view control to ensure that the
' specified item is visible. Otherwise, the macro returns FALSE.

Private Function TreeView_EnsureVisible(hWnd As Long, hItem As Long) As Boolean
    TreeView_EnsureVisible = SendMessage(hWnd, TVM_ENSUREVISIBLE, 0, ByVal hItem)
End Function
' Expands or collapses the list of child items, if any, associated with the specified parent item.
' Returns TRUE if successful or FALSE otherwise.
' (docs say TVM_EXPAND does not send the TVN_ITEMEXPANDING and
' TVN_ITEMEXPANDED notification messages to the parent window...?)

Private Function TreeView_Expand(hWnd As Long, hItem As Long, flag As ucst_TVM_EXPAND_wParam) As Boolean
    TreeView_Expand = SendMessage(hWnd, TVM_EXPAND, ByVal flag, ByVal hItem)
End Function

' Selects the specified tree-view item, scrolls the item into view, or redraws the item
' in the style used to indicate the target of a drag-and-drop operation.
' If hitem is NULL, the selection is removed from the currently selected item, if any.
' Returns TRUE if successful or FALSE otherwise.

Private Function TreeView_Select(hWnd As Long, hItem As Long, Code As Long) As Boolean
    TreeView_Select = SendMessage(hWnd, TVM_SELECTITEM, ByVal Code, ByVal hItem)
End Function

' Sets the selection to the specified item.
' Returns TRUE if successful or FALSE otherwise.

' If the specified item is already selected, a TVN_SELCHANGING *will not* be generated !!

' If the specified item is 0 (indicating to remove selection from any currrently selected item)
' and an item is selected, a TVN_SELCHANGING *will* be generated and the itemNew
' member of NMTREEVIEW will be 0 !!!

Private Function TreeView_SelectItem(hWnd As Long, hItem As Long) As Boolean
    TreeView_SelectItem = TreeView_Select(hWnd, hItem, TVGN_CARET)
End Function

' Sorts tree-view items using an application-defined callback function that compares the items.
' Returns TRUE if successful or FALSE otherwise.
' fRecurse is reserved for future use and must be zero.

Private Function TreeView_SortChildrenCB(hWnd As Long, _
                                        psort As TVSORTCB, _
                                        fRecurse As Boolean) As Boolean
    TreeView_SortChildrenCB = SendMessage(hWnd, TVM_SORTCHILDRENCB, ByVal fRecurse, psort)
End Function

Private Function MAKELPARAM(wLow As Long, wHigh As Long) As Long
  MAKELPARAM = MakeLong(wLow, wHigh)
End Function

Private Function MakeLong(wLow As Long, wHigh As Long) As Long
  MakeLong = LoWord(wLow) Or (&H10000 * LoWord(wHigh))
End Function

' Returns the lParam of the specified treeview item.
Private Function GetTVItemlParam(hwndTV As Long, hItem As Long) As Long
  Dim tVI As TVITEM
  
  tVI.hItem = hItem
  tVI.Mask = TVIF_PARAM
  
  If TreeView_GetItem(hwndTV, tVI) Then
    GetTVItemlParam = tVI.lParam
  End If

End Function
'==============================================
'END GENERIC TREEVIEW DEFS
'==============================================

Private Sub DEFINE_UUID(Name As oleexp.UUID, l As Long, w1 As Integer, w2 As Integer, B0 As Byte, b1 As Byte, b2 As Byte, B3 As Byte, b4 As Byte, b5 As Byte, b6 As Byte, b7 As Byte)
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
Private Sub DEFINE_OLEGUID(Name As oleexp.UUID, l As Long, w1 As Integer, w2 As Integer)
  DEFINE_UUID Name, l, w1, w2, &HC0, 0, 0, 0, 0, 0, 0, &H46
End Sub

Private Function IID_IShellItem() As oleexp.UUID
Static iid As oleexp.UUID
If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H43826D1E, CInt(&HE718), CInt(&H42EE), &HBC, &H55, &HA1, &HE2, &H61, &HC3, &H7B, &HFE)
IID_IShellItem = iid
End Function
Private Function IID_IShellItem2() As oleexp.UUID
'7e9fb0d3-919f-4307-ab2e-9b1860310c93
Static iid As oleexp.UUID
If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7E9FB0D3, CInt(&H919F), CInt(&H4307), &HAB, &H2E, &H9B, &H18, &H60, &H31, &HC, &H93)
IID_IShellItem2 = iid
End Function
Private Function IID_IEnumShellItems() As oleexp.UUID
'{70629033-e363-4a28-a567-0db78006e6d7}
Static iid As oleexp.UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H70629033, CInt(&HE363), CInt(&H4A28), &HA5, &H67, &HD, &HB7, &H80, &H6, &HE6, &HD7)
 IID_IEnumShellItems = iid
End Function
Private Function IID_IShellLinkW() As oleexp.UUID
'{000214F9-0000-0000-C000-000000000046}
Static iid As oleexp.UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214F9, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IShellLinkW = iid
End Function
Private Function IID_IShellIcon() As oleexp.UUID
'{000214E5-0000-0000-C000-000000000046}
Static iid As oleexp.UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H214E5, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IShellIcon = iid
End Function
Private Function IID_IImageList() As oleexp.UUID
'{46EB5926-582E-4017-9FDF-E8998DAA0950}
Static iid As oleexp.UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H46EB5926, CInt(&H582E), CInt(&H4017), &H9F, &HDF, &HE8, &H99, &H8D, &HAA, &H9, &H50)
 IID_IImageList = iid
End Function
Private Function IID_IContextMenu() As oleexp.UUID
Static iid As oleexp.UUID
 If (iid.Data1 = 0) Then Call DEFINE_OLEGUID(iid, &H214E4, 0, 0)
 IID_IContextMenu = iid
End Function
Private Function IID_IContextMenu2() As oleexp.UUID
Static iid As oleexp.UUID
 If (iid.Data1 = 0) Then Call DEFINE_OLEGUID(iid, &H214F4, 0, 0)
 IID_IContextMenu2 = iid
End Function
Private Function IID_IContextMenu3() As oleexp.UUID
'{BCFCE0A0-EC17-11d0-8D10-00A0C90F2719}
Static iid As oleexp.UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBCFCE0A0, CInt(&HEC17), CInt(&H11D0), &H8D, &H10, &H0, &HA0, &HC9, &HF, &H27, &H19)
 IID_IContextMenu3 = iid
End Function
Private Function IID_IDataObject() As oleexp.UUID
'0000010e-0000-0000-C000-000000000046
Static iid As oleexp.UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H10E, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
  IID_IDataObject = iid
End Function
Private Function IID_IDropTarget() As oleexp.UUID
'{00000122-0000-0000-C000-000000000046}
Static iid As oleexp.UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H122, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
 IID_IDropTarget = iid
End Function
Private Function IID_IShellItemArray() As oleexp.UUID
'{b63ea76d-1f85-456f-a19c-48159efa858b}
Static iid As oleexp.UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB63EA76D, CInt(&H1F85), CInt(&H456F), &HA1, &H9C, &H48, &H15, &H9E, &HFA, &H85, &H8B)
  IID_IShellItemArray = iid
End Function
Private Function IID_IPropertyDescriptionList() As oleexp.UUID
'IID_IPropertyDescriptionList, 0x1f9fc1d0, 0xc39b, 0x4b26, 0x81,0x7f, 0x01,0x19,0x67,0xd3,0x44,0x0e
Static iid As oleexp.UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1F9FC1D0, CInt(&HC39B), CInt(&H4B26), &H81, &H7F, &H1, &H19, &H67, &HD3, &H44, &HE)
  IID_IPropertyDescriptionList = iid
End Function
Private Function IID_IPropertyStore() As oleexp.UUID
'DEFINE_GUID(IID_IPropertyStore,0x886d8eeb, 0x8cf2, 0x4446, 0x8d,0x02,0xcd,0xba,0x1d,0xbd,0xcf,0x99);
Static iid As oleexp.UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H886D8EEB, CInt(&H8CF2), CInt(&H4446), &H8D, &H2, &HCD, &HBA, &H1D, &HBD, &HCF, &H99)
  IID_IPropertyStore = iid
End Function
Private Function IID_IPropertyDescription() As oleexp.UUID
'(IID_IPropertyDescription, 0x6f79d558, 0x3e96, 0x4549, 0xa1,0xd1, 0x7d,0x75,0xd2,0x28,0x88,0x14
Static iid As oleexp.UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H6F79D558, CInt(&H3E96), CInt(&H4549), &HA1, &HD1, &H7D, &H75, &HD2, &H28, &H88, &H14)
  IID_IPropertyDescription = iid
End Function
Private Function IID_IQueryInfo() As oleexp.UUID
'{00021500-0000-0000-C000-000000000046}
Static iid As oleexp.UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H21500, CInt(&H0), CInt(&H0), &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46)
IID_IQueryInfo = iid
End Function

Private Function BHID_EnumItems() As oleexp.UUID
'{0x94F60519, 0x2850, 0x4924, 0xAA,0x5A, 0xD1,0x5E,0x84,0x86,0x80,0x39}
Static iid As oleexp.UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H94F60519, &H2850, &H4924, &HAA, &H5A, &HD1, &H5E, &H84, &H86, &H80, &H39)
 BHID_EnumItems = iid
End Function
Private Function BHID_DataObject() As oleexp.UUID
'{0xB8C0BD9F, 0xED24, 0x455C, 0x83,0xE6, 0xD5,0x39,0x0C,0x4F,0xE8,0xC4}
Static iid As oleexp.UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB8C0BD9F, &HED24, &H455C, &H83, &HE6, &HD5, &H39, &HC, &H4F, &HE8, &HC4)
 BHID_DataObject = iid
End Function
Private Function BHID_SFUIObject() As oleexp.UUID
'DEFINE_GUID(BHID_SFUIObject,  0x3981E225, 0xF559, 0x11D3, 0x8E,0x3A, 0x00,0xC0,0x4F,0x68,0x37,0xD5);
'{3981e225-f559-11d3-8e3a-00c04f6837d5}
Static iid As oleexp.UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H3981E225, &HF559, &H11D3, &H8E, &H3A, &H0, &HC0, &H4F, &H68, &H37, &HD5)
  BHID_SFUIObject = iid
End Function

Private Function FOLDERID_ComputerFolder() As oleexp.UUID
Static iid As oleexp.UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HAC0837C, CInt(&HBBF8), CInt(&H452A), &H85, &HD, &H79, &HD0, &H8E, &H66, &H7C, &HA7)
 FOLDERID_ComputerFolder = iid
End Function
Private Function FOLDERID_Windows() As oleexp.UUID
Static iid As oleexp.UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HF38BF404, CInt(&H1D43), CInt(&H42F2), &H93, &H5, &H67, &HDE, &HB, &H28, &HFC, &H23)
 FOLDERID_Windows = iid
End Function
Private Function FOLDERID_Favorites() As oleexp.UUID
Static iid As oleexp.UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H1777F761, CInt(&H68AD), CInt(&H4D8A), &H87, &HBD, &H30, &HB7, &H59, &HFA, &H33, &HDD)
 FOLDERID_Favorites = iid
End Function
Private Function FOLDERID_Links() As oleexp.UUID
Static iid As oleexp.UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBFB9D5E0, CInt(&HC6A9), CInt(&H404C), &HB2, &HB2, &HAE, &H6D, &HB6, &HAF, &H49, &H68)
 FOLDERID_Links = iid
End Function
Private Function FOLDERID_Desktop() As oleexp.UUID
Static iid As oleexp.UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB4BFCC3A, CInt(&HDB2C), CInt(&H424C), &HB0, &H29, &H7F, &HE9, &H9A, &H87, &HC6, &H41)
 FOLDERID_Desktop = iid
End Function
Private Function FOLDERID_UserProfiles() As oleexp.UUID
Static iid As oleexp.UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H762D272, CInt(&HC50A), CInt(&H4BB0), &HA3, &H82, &H69, &H7D, &HCD, &H72, &H9B, &H80)
 FOLDERID_UserProfiles = iid
End Function
Private Function FOLDERID_Profile() As oleexp.UUID
Static iid As oleexp.UUID
 If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H5E6C858F, CInt(&HE22), CInt(&H4760), &H9A, &HFE, &HEA, &H33, &H17, &HB6, &H71, &H73)
 FOLDERID_Profile = iid
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

Private Function PKEY_PropList_InfoTip() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC9944A21, &HA406, &H48FE, &H82, &H25, &HAE, &HC7, &HE2, &H4C, &H21, &H1B, 4)
PKEY_PropList_InfoTip = pkk
End Function

Private Sub SetNavWav()
Dim lp As Long, sz As String
Dim kfm As New oleexp.KnownFolderManager
Dim kfwin As oleexp.IKnownFolder
kfm.GetFolder FOLDERID_Windows, kfwin
If (kfwin Is Nothing) = False Then
    kfwin.GetPath KF_FLAG_DEFAULT, lp
    sz = LPWSTRtoStr(lp)
    
    szNavWav = AddBackslash(sz) & "Media\Windows Navigation Start.wav"
End If

End Sub

Private Sub PlayNavSound()
If mNavSound = False Then Exit Sub
If szNavWav = "" Then SetNavWav
PlaySound StrPtr(szNavWav), 0&, SND_ASYNC Or SND_NODEFAULT Or SND_FILENAME
DebugAppend "PlayNavSound " & szNavWav

End Sub


Public Sub dbg_checkstyle()
DebugAppend "dbg_checkstyle"
Dim dwStyle As ucst_TV_Styles
dwStyle = GetWindowLong(hTVD, GWL_STYLE)
If (dwStyle Or TVS_DISABLEDRAGDROP) = TVS_DISABLEDRAGDROP Then
    DebugAppend "DragDrop disabled"
Else
    DebugAppend "No dd flag found, " & (dwStyle Or TVS_DISABLEDRAGDROP)
End If
End Sub
Private Function TVItemGetDispColor(lp As Long) As Long
'Compressed = blue, encrypted = green
If (mNameColors = False) Then
    OleTranslateColor clrFore, 0&, TVItemGetDispColor
 ' m_SysClrText
    Exit Function
End If
'DebugAppend "QueryColor " & TVEntries(lp).sName

If (TVEntries(lp).dwAttrib And SFGAO_COMPRESSED) = SFGAO_COMPRESSED Then
    TVItemGetDispColor = vbBlue
    Exit Function
End If
If (TVEntries(lp).dwAttrib And SFGAO_ENCRYPTED) = SFGAO_ENCRYPTED Then
    TVItemGetDispColor = &HD9213
    Exit Function
End If
OleTranslateColor clrFore, 0&, TVItemGetDispColor
End Function

Private Sub DoPaths(hItem As Long)
'add checked folders to the path list
Dim tVI As TVITEM
Dim uState As Long
Dim hItemChild As Long
Dim sPath As String
Dim nIdx As Long
nItr = nItr + 1
If nItr > 10000 Then Exit Sub
ReDim sR(0)

If hItem = 0 Then hItem = TreeView_GetRoot(hTVD)

Do While hItem
    tVI.hItem = hItem
    tVI.Mask = TVIF_CHILDREN Or TVIF_STATE Or TVIF_PARAM
    tVI.StateMask = TVIS_STATEIMAGEMASK
    TreeView_GetItem hTVD, tVI
    
    sPath = TVEntries(tVI.lParam).sFullPath
    DebugAppend "hitem=" & hItem & ",path=" & sPath
    If (tVI.Mask) And TVIS_STATEIMAGEMASK = &H2000 Then 'check item; add it
        ReDim Preserve gPaths(nPaths)
        gPaths(nPaths) = sPath 'GetFolderDisplayName(tvid.isfParent, tvid.pidlRel, SHGDN_INFOLDER) '
        nPaths = nPaths + 1
    Else
        If tVI.cChildren > 0 Then
            hItemChild = TreeView_GetNextItem(hTVD, hItem, TVGN_CHILD)
            DoPaths hItemChild
        End If
    End If
    
    hItem = TreeView_GetNextItem(hTVD, hItem, TVGN_NEXT)
Loop
End Sub

Private Sub EnumPaths(hWnd As Long, ByVal hItem As Long)
'DebugAppend "EnumPaths.Entry"
Dim tVI As TVITEM
Dim hItemChild As Long
Dim aTmp() As String
Dim i As Long
Dim nIdx As Long
Dim sPath As String
    Do While hItem
        tVI.hItem = hItem
        tVI.Mask = TVIF_CHILDREN Or TVIF_STATE Or TVIF_HANDLE
        tVI.StateMask = TVIS_STATEIMAGEMASK 'TVIS_BOLD
        TreeView_GetItem hTVD, tVI

        If (tVI.State And TVIS_STATEIMAGEMASK) = &H2000 Then 'TVIS_BOLD Then
            DebugAppend "checked item found, " & GetTVItemText(hWnd, hItem) & "-" & hItem
            
            If TVEntries(tVI.lParam).sFullPath <> "" Then 'merge into main results
                    DebugAppend "Add checked path " & TVEntries(tVI.lParam).sFullPath
                    ReDim Preserve gPaths(nPaths)
                    gPaths(nPaths) = TVEntries(tVI.lParam).sFullPath
                    nPaths = nPaths + 1
            End If
        Else
        hItemChild = TreeView_GetChild(hTVD, tVI.hItem)
        Call EnumPaths(hWnd, hItemChild)
        End If
        
        hItem = TreeView_GetNextSibling(hTVD, hItem)
    Loop
End Sub

Private Function GetNodeByPath(sFullPath As String, Optional bStripComp As Boolean = False) As Long
Dim i As Long
Dim sNoCmp As String
For i = 0 To UBound(TVEntries)
    sNoCmp = TVEntries(i).sFullPath
    If bStripComp Then
        If InStr(sNoCmp, sComp & "\") Then
            sNoCmp = Mid$(sNoCmp, Len(sComp & "\") + 1)
        End If
    End If
    If sNoCmp = sFullPath Then
        GetNodeByPath = TVEntries(i).hNode
        Exit Function
    End If
Next i
GetNodeByPath = -1
End Function

Private Sub CreateCurrentMap()
Dim hRoot As Long
hRoot = TreeView_GetRoot(hTVD)

ReDim TVVisMap(0)
TVVisMap(0) = TVEntries(0) 'copy desktop
nTVVis = 1
MapVisibleItems hTVD, hRoot
End Sub

Private Sub MapVisibleItems(hWnd As Long, hItem As Long)
Dim tVI As TVITEM
Dim tviPar As TVITEM
Dim hItemPar As Long
Dim hItemChild As Long
Dim hParLast As Long
Dim bRunCheck As Boolean
Dim lCR As Long
Dim aTmp() As String
Dim i As Long
Dim nIdx As Long
Dim sPath As String
    Do While hItem
        tVI.hItem = hItem
        tVI.Mask = TVIF_CHILDREN Or TVIF_STATE Or TVIF_HANDLE Or TVIF_PARAM
        tVI.StateMask = TVIS_STATEIMAGEMASK Or TVIS_EXPANDED 'TVIS_BOLD
        TreeView_GetItem hTVD, tVI
'        If (tvi.State And TVIS_EXPANDED) = TVIS_EXPANDED Then 'TVIS_BOLD Then
        If True Then
            hItemPar = TreeView_GetParent(hTVD, tVI.hItem)
            If hItemPar Then
                tviPar.hItem = hItemPar
                tviPar.Mask = TVIF_STATE Or TVIF_HANDLE Or TVIF_PARAM
                tviPar.StateMask = TVIS_EXPANDED
                TreeView_GetItem hTVD, tviPar
                If (tviPar.State And TVIS_EXPANDED) = TVIS_EXPANDED Then
                    If hItemPar <> hParLast Then
                        lCR = TVCheckExpand(hItemPar)
                    End If
                    If lCR = S_OK Then 'Check all higher level parents to make sure they're expanded too
'                    DebugAppend "Found visible item, " & GetTVItemText(hwnd, hItem) & "/" & TVEntries(tvi.lParam).sFullPath
                    ReDim Preserve TVVisMap(nTVVis)
                    TVVisMap(nTVVis) = TVEntries(tVI.lParam)
                    nTVVis = nTVVis + 1
                    End If
                End If
            End If
        hItemChild = TreeView_GetChild(hTVD, tVI.hItem)
        Call MapVisibleItems(hWnd, hItemChild)
        Else
        End If
                    hParLast = hItemPar
        
        hItem = TreeView_GetNextSibling(hTVD, hItem)
    Loop
End Sub

Private Function TVCheckExpand(hNode As Long) As Long
Dim tVI As TVITEM
Dim tviPar As TVITEM
Dim hItemPar As Long
Dim hItemChild As Long
'If hNode = m_hRoot Then
    hItemPar = hNode
'Else
'    hItemPar = TreeView_GetParent(hTVD, tVI.hItem)
'End If
    Do While hItemPar
'        DebugAppend "TVCheckExpand::Checking " & GetTVItemText(hTVD, hItemPar)
        tviPar.hItem = hItemPar
        tviPar.Mask = TVIF_STATE Or TVIF_HANDLE Or TVIF_PARAM
        tviPar.StateMask = TVIS_EXPANDED
        TreeView_GetItem hTVD, tviPar
        If (tviPar.State And TVIS_EXPANDED) = 0& Then
''            DebugAppend "TVCheckExpand::Unexpanded Parent->" & GetTVItemText(hTVD, hItemPar)
            TVCheckExpand = 1&
            Exit Function
        Else
'            DebugAppend "TVCheckExpand::Expanded Parent->" & GetTVItemText(hTVD, hItemPar)
        
        End If
        hItemPar = TreeView_GetParent(hTVD, hItemPar)
    Loop
        
End Function

Private Function IsUSBDevice(sParse As String) As Boolean
If Left$(sParse, 49) = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}\\\?\usb#" Then
    IsUSBDevice = True
End If
End Function

Private Function ParseUSBDevice(sFull As String) As String
Dim s1 As String
Dim s2 As String
Dim n1 As Long
s1 = sFull
s1 = Mid$(s1, InStr(s1, "\\\?\usb#") + 1)

n1 = InStr(6, s1, "\")
If n1 = 0 Then
    ParseUSBDevice = s1
Else
    s2 = Left$(s1, n1 - 1)
    ParseUSBDevice = s2
End If
End Function

Private Function USBDeviceEnsureAdded(sDev As String) As Long
Dim hComp As Long
Dim sUSB As String
sUSB = ParseUSBDevice(sDev)
Dim hDev As Long
hDev = GetNodeByPath(sUSB, True)
DebugAppend "hDev=" & hDev
If hDev <> -1 Then
    DebugAppend "USBDeviceEnsureAdded::Device already added."
    USBDeviceEnsureAdded = hDev
    Exit Function
End If

hComp = GetNodeByPath(sComp)
If hComp Then
    Dim si As oleexp.IShellItem
    oleexp.SHCreateItemFromParsingName StrPtr(sComp & "\" & sUSB), Nothing, IID_IShellItem, si
    If (si Is Nothing) = False Then
        DebugAppend "USBDeviceEnsureAdded::Got device item, adding..."
        TVAddItem si, hComp
        If InStr(TVEntries(UBound(TVEntries)).sFullPath, "\\?\usb#") Then
            USBDeviceEnsureAdded = TVEntries(UBound(TVEntries)).hNode
            Exit Function
        End If
    End If
End If
USBDeviceEnsureAdded = -1
End Function

Private Sub pSetRedrawMode(ByVal Enable As Boolean)
   If (hTVD) Then
      Call SendMessage(hTVD, WM_SETREDRAW, -Enable, ByVal 0&)
   End If
End Sub

Public Sub OpenToPath(sFullPath As String, bExpandTarget As Boolean, Optional bSelectTarget As Boolean = True)
'OpenToItem by path instead of shell item
If sFullPath = sSelectedItem Then Exit Sub
Dim psi As oleexp.IShellItem
oleexp.SHCreateItemFromParsingName StrPtr(sFullPath), Nothing, IID_IShellItem, psi
If (psi Is Nothing) = False Then
     OpenToItem psi, bExpandTarget, bSelectTarget
 End If
End Sub

Public Sub OpenToItem(si As oleexp.IShellItem, bExpandTarget As Boolean, Optional bSelectTarget As Boolean = True)
If bLoadDone = False Then Exit Sub
Dim lpFull As Long, sFull As String
si.GetDisplayName SIGDN_DESKTOPABSOLUTEPARSING, lpFull
sFull = LPWSTRtoStr(lpFull)
If (sFull <> "") Then
    If sFull = sSelectedItem Then Exit Sub
End If
bNavigating = True
bFlagBlockClkExp = False
pSetRedrawMode False
Hourglass True
lLoopTrack = 0
DebugAppend "OpenToItem " & sFull
If IsUSBDevice(sFull) Then
    DebugAppend "OpenToItem->Detected USB device, ensuring add..."
    Dim hrua As Long
    hrua = USBDeviceEnsureAdded(sFull)
    DebugAppend "OpenToItem->EnsureAdd result=" & hrua
End If
DebugAppend bExpandTarget
TVNavigate si, bExpandTarget, bSelectTarget
Hourglass False
pSetRedrawMode True
bNavigating = False
End Sub

Private Sub TVNavigate(si As oleexp.IShellItem, bExpandTarget As Boolean, bSelectTarget As Boolean)
Dim i As Long
Dim bFlag As Boolean
Dim lp As Long
Dim lpFull As Long, sFull As String
Dim lAtr As SFGAO_Flags
Dim hTop As Long
Dim hPar As Long
Dim siPar As oleexp.IShellItem, siPar2 As oleexp.IShellItem
Dim idx As Long
Dim bPaged As Boolean
Dim hAdd As Long

If lLoopTrack > 100 Then
    DebugAppend "Error: Infinite loop detected."
    Exit Sub
End If
lLoopTrack = lLoopTrack + 1
si.GetAttributes SFGAO_FOLDER Or SFGAO_STREAM, lAtr
If (mShowFiles = False) And ((lAtr And SFGAO_FOLDER) = 0&) Then Exit Sub 'it's not a folder, so exit if we're not showing files too
If (mExpandZip = False) And ((lAtr And SFGAO_STREAM) = SFGAO_STREAM) Then Exit Sub 'Zip files only shown if mExpandZip=True

DebugAppend "TVNavigate->CreateCurrentMap"
CreateCurrentMap   'Map the visible items (the main entry list contains items in collapsed folders)

si.GetDisplayName SIGDN_DESKTOPABSOLUTEPARSING, lpFull
sFull = LPWSTRtoStr(lpFull)
DebugAppend "TVNavigate " & sFull & ", " & bExpandTarget
'First, check to see if it's already present
DebugAppend "TVNavigate->GetIdxByPath"
idx = GetMapIdxByPath(sFull)
If idx > -1 Then
        DebugAppend "OpenToItem::Item already visible"
        If bFlagRecurseOI = False Then
            fNoExpand = 1
            If bSelectTarget Then
                TreeView_EnsureVisible hTVD, TVVisMap(idx).hNode
                TreeView_SelectItem hTVD, TVVisMap(idx).hNode
            End If
'            TreeView_Select hTVD, TVVisMap(idx).hNode, TVGN_DROPHILITE
            fNoExpand = 0
        End If
        If (TVVisMap(idx).bFolder) And (bExpandTarget = True) Then
            DebugAppend "OpenToItem::AlreadyVis->Expand"
            TreeView_Expand hTVD, TVVisMap(idx).hNode, TVE_EXPAND
        End If
        SetFocus UserControl.ContainerHwnd
        SetFocus hTVD
        Exit Sub
End If

If sFull = sUserDesktop Then
    SHGetKnownFolderItem FOLDERID_Profile, KF_FLAG_DEFAULT, 0&, IID_IShellItem, siPar
Else
    si.GetParent siPar
End If
If (siPar Is Nothing) = False Then
    lpFull = 0: sFull = ""
    siPar.GetDisplayName SIGDN_DESKTOPABSOLUTEPARSING, lpFull
    sFull = LPWSTRtoStr(lpFull)
    DebugAppend "TVNavigate::FirstParent=" & sFull
    bPaged = False
    bPaged = IsPagedVirtualPath(sFull)
    If bPaged Then
        sFull = Left$(sFull, Len(sFull) - 2)
        DebugAppend "TVNavigate::Adjusted paged virtual item parent to " & sFull
    End If
    idx = GetMapIdxByPath(sFull)
    If idx > -1 Then
        DebugAppend "Found at first parent: " & TVVisMap(idx).sName & "," & TVVisMap(idx).sFullPath
        TreeView_Expand hTVD, TVVisMap(idx).hNode, TVE_EXPAND
        If Not bPaged Then
            TVNavigate si, bExpandTarget, bSelectTarget
        Else
            TVAddItem si, TVVisMap(idx).hNode, True
            hAdd = TVEntries(UBound(TVEntries)).hNode
            TreeView_EnsureVisible hTVD, hAdd
            TreeView_SelectItem hTVD, hAdd
        
        End If
        Exit Sub
    End If
Else
    UpdateStatus "Path not found."
    Exit Sub
End If

Dim siRm() As oleexp.IShellItem
Dim rmct As Long
              
Dim nStartFrom As Long
Dim j As Long, k As Long
Dim sPrev As String
ReDim siRm(0)
Set siRm(0) = siPar
rmct = rmct + 1
Do
    ReDim Preserve siRm(rmct)
    If siRm(rmct - 1).GetParent(siRm(rmct)) = S_OK Then
        sPrev = sFull
        lpFull = 0: sFull = ""
        siRm(rmct).GetDisplayName SIGDN_DESKTOPABSOLUTEPARSING, lpFull
        sFull = LPWSTRtoStr(lpFull)
        If (sFull = sUserDesktop) And (sPrev = sUserFolder) Then
            'Parent is \Users\ not desktop wtf!
            sFull = "C:\Users"
        End If
        DebugAppend "AddParent(" & rmct & ")=" & sFull
        idx = GetMapIdxByPath(sFull)
        If idx > -1 Then nStartFrom = rmct
        rmct = rmct + 1
    Else
        DebugAppend "Ran out of parents? Not found"
        UpdateStatus "Path not found."
        TVAddItem si, m_hRoot, True
        hAdd = TVEntries(UBound(TVEntries)).hNode
        TreeView_EnsureVisible hTVD, hAdd
        TreeView_SelectItem hTVD, hAdd

        Exit Do
    End If
Loop Until nStartFrom <> 0 'if it was 0 our first parent block would have succeeded

If nStartFrom Then
    DebugAppend "ParentStartIndex=" & nStartFrom
    TreeView_Expand hTVD, TVVisMap(idx).hNode, TVE_EXPAND
    bFlagRecurseOI = True
    For j = (nStartFrom - 1) To 0 Step -1
        TVNavigate siRm(j), True, bSelectTarget
    Next j
    bFlagRecurseOI = False
    TVNavigate si, bExpandTarget, bSelectTarget
End If

DebugAppend "TVNavigate.End"

End Sub

Private Function IsPagedVirtualPath(sPath As String) As Boolean
'Automatic navigation has problems with paths like ::GUID\n control panel items that can't actually be expanded to find the child
Dim i As Long, j As Long

If Left$(sPath, 3) = "::{" Then
    For i = 0 To 9
        If Right$(sPath, 2) = "\" & CStr(i) Then
            IsPagedVirtualPath = True
            Exit Function
        End If
   Next i
End If
   
End Function

Private Function GetMapIdxByPath(sFullPath As String) As Long
Dim i As Long
For i = 0 To UBound(TVVisMap)
    If LCase$(TVVisMap(i).sFullPath) = LCase$(sFullPath) Then
        GetMapIdxByPath = i
        Exit Function
    End If
Next i
GetMapIdxByPath = -1
End Function

Private Function GetTVSelectedItemPath() As String
'For single-select mode
Dim tVI As TVITEM
Dim hItem As Long
Dim lp As Long
Dim aTmp() As String
Dim lpsz As Long
Dim lpAtr As SFGAO_Flags
Dim nIdx As Long
On Error GoTo e0

hItem = TreeView_GetSelection(hTVD)
If hItem Then
    lp = GetTVItemlParam(hTVD, hItem)
    If lp Then
        GetTVSelectedItemPath = TVEntries(lp).sFullPath
    Else
        DebugAppend "mTreeView.GetTVSelectedItemPath.Error->Failed to get item lParam", 2
    End If
Else
    DebugAppend "mTreeView.GetTVSelectedItemPath.Warning->GetSelection did not return an hItem", 3
End If

Exit Function

e0:
DebugAppend "mTreeView.GetTVSelectedItemPath.Error->" & Err.Description & " (" & Err.Number & ")"

End Function

Private Sub TVUncheckAllChildren(hWnd As Long, hItem As Long)
'pass the first child item of the node
Dim hItemChild As Long
Dim i As Long
    Do While hItem
'        SetTVItemStateImage hwnd, hItem, tvcsEmpty
        TreeView_SetCheckState hTVD, hItem, 0
        hItemChild = TreeView_GetChild(hTVD, hItem)
        Call TVUncheckAllChildren(hWnd, hItemChild)
        
        hItem = TreeView_GetNextSibling(hTVD, hItem)
    Loop
End Sub

Private Function TVItemIsZip(hWnd As Long, lParam As Long) As Boolean
Dim lAtr As SFGAO_Flags

lAtr = SFGAO_STREAM 'it already has SFGAO_FOLDER by virtue of being in the list
Dim nIdx As Long
nIdx = GetIndexFromNode(lParam)
lAtr = TVEntries(lParam).dwAttrib
If (lAtr And SFGAO_STREAM) Then TVItemIsZip = True

End Function

Private Function TVNodeHasUncheckedChildren(hItem As Long) As Boolean
Dim hChild As Long
Dim tVI As TVITEM
Dim cst As Long

hChild = TreeView_GetChild(hTVD, hItem)
Do While hChild
    cst = TreeView_GetCheckState(hTVD, hChild)
    If (cst = 1) Or (cst >= 3) Then
        TVNodeHasUncheckedChildren = True
        Exit Function
    End If
    hChild = TreeView_GetNextSibling(hTVD, hChild)
Loop

End Function

Private Function TVNodeHasCheckedChildren(hItem As Long) As Boolean
Dim hChild As Long
Dim tVI As TVITEM
Dim cst As Long

hChild = TreeView_GetChild(hTVD, hItem)
Do While hChild
    cst = TreeView_GetCheckState(hTVD, hChild)
    If (cst = 2) Or (cst = 3) Then
        TVNodeHasCheckedChildren = True
        DebugAppend "Found checked child: " & GetTVItemText(hTVD, hChild)
        Exit Function
    End If
    hChild = TreeView_GetNextSibling(hTVD, hChild)
Loop

End Function

Private Function TVSetParentAfterCheck(hPar As Long) As Long
DebugAppend "TVSetParentAfterCheck " & GetTVItemText(hTVD, hPar)
Dim hPar2 As Long
Dim tVI As TVITEM
    If ((hPar > 0)) Then 'And (hPar <> m_hRoot)

hPar2 = hPar
'Do While ((hPar2 > 0) And (hPar2 <> m_hRoot))
    If TVNodeHasUncheckedChildren(hPar2) Then
        TreeView_SetCheckStateEx hTVD, hPar2, 3
    Else
        TreeView_SetCheckStateEx hTVD, hPar2, 2
    End If
    hPar2 = TreeView_GetParent(hTVD, hPar2)
    If ((hPar2 > 0)) Then ' And (hPar2 <> m_hRoot)) Then
        TVSetParentAfterCheck hPar2
    End If
End If
End Function

Private Function TVSetParentAfterUncheck(hPar As Long) As Long
DebugAppend "TVSetParentAfterUncheck " & GetTVItemText(hTVD, hPar)
Dim hPar2 As Long
Dim tVI As TVITEM
    If ((hPar > 0)) Then ' And (hPar <> m_hRoot)) Then

hPar2 = hPar
'Do While ((hPar2 > 0) And (hPar2 <> m_hRoot))
    If TVNodeHasCheckedChildren(hPar2) Then
        TreeView_SetCheckStateEx hTVD, hPar2, 3
    Else
        TreeView_SetCheckStateEx hTVD, hPar2, 1
    End If
    hPar2 = TreeView_GetParent(hTVD, hPar2)
    If ((hPar2 > 0)) Then ' And (hPar2 <> m_hRoot)) Then
        TVSetParentAfterUncheck hPar2
    End If
End If
End Function

Private Function TVUncheckChildren(hItem As Long) As Long
DebugAppend "TVUnCheckChildren.Entry=" & GetTVItemText(hTVD, hItem)
Dim hChild As Long
Dim hSib As Long
Dim tVI As TVITEM

hSib = hItem
Do While hSib
    tVI.Mask = TVIF_CHILDREN Or TVIF_HANDLE Or TVIF_STATE
    tVI.hItem = hSib
    TreeView_GetItem hTVD, tVI
    
    TreeView_SetCheckState hTVD, hSib, 0
    If (tVI.cChildren) And ((tVI.State And TVIS_EXPANDEDONCE) = TVIS_EXPANDEDONCE) Then
        hChild = TreeView_GetChild(hTVD, hSib)
        TVCheckChildren hChild
    End If
    hSib = TreeView_GetNextSibling(hTVD, hSib)
Loop
End Function

Private Function TVCheckChildren(hItem As Long) As Long
DebugAppend "TVCheckChildren.Entry=" & GetTVItemText(hTVD, hItem)
Dim hChild As Long
Dim hSib As Long
Dim tVI As TVITEM

hSib = hItem
Do While hSib
    tVI.Mask = TVIF_CHILDREN Or TVIF_HANDLE Or TVIF_STATE
    tVI.hItem = hSib
    TreeView_GetItem hTVD, tVI
    
    TreeView_SetCheckState hTVD, hSib, 1
    If (tVI.cChildren) And ((tVI.State And TVIS_EXPANDEDONCE) = TVIS_EXPANDEDONCE) Then
        hChild = TreeView_GetChild(hTVD, hSib)
        TVCheckChildren hChild
    End If
    hSib = TreeView_GetNextSibling(hTVD, hSib)
Loop

End Function

Private Function TVExcludeChildren(hItem As Long) As Long
DebugAppend "TVCheckChildren.Entry=" & GetTVItemText(hTVD, hItem)
Dim hChild As Long
Dim hSib As Long
Dim tVI As TVITEM

hSib = hItem
Do While hSib
    tVI.Mask = TVIF_CHILDREN Or TVIF_HANDLE Or TVIF_STATE
    tVI.hItem = hSib
    TreeView_GetItem hTVD, tVI
    
    TreeView_SetCheckStateEx hTVD, hSib, 4
    If (tVI.cChildren) And ((tVI.State And TVIS_EXPANDEDONCE) = TVIS_EXPANDEDONCE) Then
        hChild = TreeView_GetChild(hTVD, hSib)
        TVCheckChildren hChild
    End If
    hSib = TreeView_GetNextSibling(hTVD, hSib)
Loop

End Function

Public Sub SelectNone()
Dim nSel As Long
Dim apidl() As Long, cpidl As Long
Dim lpSel As Long
Dim i As Long
ReDim sSelectedItems(0): ReDim apidl(0)
DebugAppend "SelectNone"
Dim hItemSel As Long
Do
    hItemSel = 0
    hItemSel = TreeView_GetNextSelected(hTVD, hItemSel)
    If hItemSel Then
        TreeView_SetItemState hTVD, hItemSel, 0&, TVIS_SELECTED
    End If
Loop While hItemSel

End Sub

Private Sub SetMultiSel()
Dim nSel As Long
Dim apidl() As Long, cpidl As Long
Dim aNm() As String
Dim lpSel As Long
Dim i As Long
ReDim sSelectedItems(0): ReDim apidl(0)

DebugAppend "TVN_BEGINDRAG"
Dim hItemSel As Long
Do
    hItemSel = TreeView_GetNextSelected(hTVD, hItemSel)
    If hItemSel Then
        lpSel = GetTVItemlParam(hTVD, hItemSel)
        ReDim Preserve sSelectedItems(nSel)
        ReDim Preserve aNm(nSel)
        sSelectedItems(nSel) = TVEntries(lpSel).sFullPath
        aNm(nSel) = TVEntries(lpSel).sName
        nSel = nSel + 1
        ReDim Preserve apidl(cpidl)
        apidl(cpidl) = ILCombine(TVEntries(lpSel).pidlFQPar, TVEntries(lpSel).pidlRel)
        cpidl = cpidl + 1
    End If
Loop While hItemSel
If nSel > 0& Then
    Dim psia As oleexp.IShellItemArray
    Dim iData As oleexp.IDataObject
    DebugAppend "Dragging " & nSel & " items"
    oleexp.SHCreateShellItemArrayFromIDLists cpidl, VarPtr(apidl(0)), siaSelected
End If
For i = 0 To UBound(apidl)
    If apidl(i) Then CoTaskMemFree apidl(i)
Next i

If (fLoad = 0&) Then RaiseEvent MultiSelectChange(siaSelected, aNm, sSelectedItems)

End Sub

Public Sub SelectedItems(sFullPaths() As String, siaItems As oleexp.IShellItemArray)
'SetMultiSel
'sFullPaths = sSelectedItems
Set siaItems = siaSelected
End Sub

Private Function GetStrFromPtrA(lpszA As Long) As String
  Dim sRtn As String
  sRtn = String$(lstrlenA(ByVal lpszA), 0)
  Call lstrcpyA(ByVal sRtn, ByVal lpszA)
  GetStrFromPtrA = sRtn
End Function

Private Sub ShowBalloonTipEx(hWnd As Long, sTitle As String, sText As String, btIcon As ucst_BalloonTipIconConstants)
Dim lr As Long
Dim tEBT As EDITBALLOONTIP
tEBT.cbStruct = LenB(tEBT)
tEBT.pszText = StrPtr(sText)
tEBT.pszTitle = StrPtr(sTitle)
tEBT.ttiIcon = btIcon
lr = SendMessageW(hWnd, EM_SHOWBALLOONTIP, 0, tEBT)
'DebugAppend "ShowBalloonTipEx=" & lR
End Sub

Private Function IsClipboardValidFileName() As Integer
Dim i As Long
Dim sz As String
Dim sChr As String

'sz = Clipboard.GetText
sz = GetClipboardTextW()
'dbg_stringbytes sz
IsClipboardValidFileName = 1

If sz = vbNullString Then
    DebugAppend "Clip is null"
    IsClipboardValidFileName = -1
End If
If Len(sz) > MAX_PATH Then IsClipboardValidFileName = -2

If InStr(sz, "*") Or InStr(sz, "<") Or InStr(sz, ">") Or InStr(sz, "|") Or InStr(sz, Chr$(34)) Or InStr(sz, Chr(&H3F)) Then
    IsClipboardValidFileName = -1
End If

DebugAppend "ClipCheck=" & IsClipboardValidFileName & "," & InStr(sz, "?")

End Function

Private Function GetClipboardTextW() As String
    '<EhHeader>
    On Error GoTo e0
    '</EhHeader>
Dim pdo As oleexp.IDataObject
Dim tSTG As oleexp.STGMEDIUM
Dim tFMT As oleexp.FORMATETC
Dim lpGlobal As Long
Dim hGlobal As Long
Dim lpText As Long
Dim stBuf As String

OleGetClipboard pdo
If (pdo Is Nothing) Then Exit Function

If DataObjSupportsFormat(pdo, CF_UNICODETEXT) Then
    'This should return True and successfully retrieve the text even if technically
    'the clipboard object actually only contains CF_TEXT.
    DebugAppend "DataObj Get CF_UNICODETEXT"
    
    tFMT.cfFormat = CF_UNICODETEXT
    tFMT.dwAspect = DVASPECT_CONTENT
    tFMT.lIndex = -1
    tFMT.TYMED = TYMED_HGLOBAL

    pdo.GetData tFMT, tSTG
    lpText = GlobalLock(tSTG.data)
    GetClipboardTextW = LPWSTRtoStr(lpText, False)
    Call GlobalUnlock(tSTG.data)
    ReleaseStgMedium tSTG
ElseIf DataObjSupportsFormat(pdo, CF_TEXT) Then 'if there's no unicode chars there might be this only
    tFMT.cfFormat = CF_TEXT
    tFMT.dwAspect = DVASPECT_CONTENT
    tFMT.lIndex = -1
    tFMT.TYMED = TYMED_HGLOBAL

    pdo.GetData tFMT, tSTG
    lpText = GlobalLock(tSTG.data)
    GetClipboardTextW = GetStrFromPtrA(lpText)
    DebugAppend "GetClipboardText ANSI fallback=" & GetClipboardTextW
End If
'<EhFooter>
Exit Function

e0:
    DebugAppend "ucShellBrowse.GetClipboardTextW->Error: " & Err.Description & ", 0x" & Hex$(Err.Number)
'</EhFooter>
End Function

Private Function VerifyDropTarget(lIdx As Long) As Long
Dim dt As oleexp.IDropTarget
Dim si As oleexp.IShellItem
Dim tVI As TVITEM
Dim de As DROPEFFECTS
Dim ppidlp As Long
On Error GoTo e0

tVI.hItem = lIdx
tVI.Mask = TVIF_HANDLE Or TVIF_PARAM
TreeView_GetItem hTVD, tVI

If tVI.lParam Then
    If TVEntries(tVI.lParam).sFullPath = "1" Then
        SHGetKnownFolderItem FOLDERID_Links, KF_FLAG_DEFAULT, 0&, IID_IShellItem, si
    Else
        oleexp.SHCreateItemFromParsingName StrPtr(TVEntries(tVI.lParam).sFullPath), Nothing, IID_IShellItem, si
        If (si Is Nothing) Then
            DebugAppend "VerifyDropTarget->ShellItem failed, trying to load from pidls...", 3
            If (TVEntries(tVI.lParam).pidlFQPar <> 0&) And (TVEntries(tVI.lParam).pidlRel <> 0&) Then
                ppidlp = ILCombine(TVEntries(tVI.lParam).pidlFQPar, TVEntries(tVI.lParam).pidlRel)
                oleexp.SHCreateItemFromIDList ppidlp, IID_IShellItem, si
                If (si Is Nothing) = False Then
                    DebugAppend "VerifyDropTarget->LoadFromPidl->Success!", 3
                Else
                    DebugAppend "VerifyDropTarget->LoadFromPidl->Failed.", 3
                End If
            Else
                DebugAppend "VerifyDropTarget->LoadFromPidl->Parent or child pidl not set.", 3
            End If
        End If
    End If
    If (si Is Nothing) = False Then
        si.BindToHandler 0&, BHID_SFUIObject, IID_IDropTarget, dt
        If (dt Is Nothing) = False Then
            VerifyDropTarget = 1
            de = DROPEFFECT_MOVE Or DROPEFFECT_COPY Or DROPEFFECT_LINK
            dt.DragEnter mDataObj, 0&, 0&, 0&, de
            dt.DragOver MK_LBUTTON, 0&, 0&, de
            If TVEntries(tVI.lParam).bLinkIsFolder Then
                VerifyDropTarget = 2
                GoTo out
            End If
            If de = DROPEFFECT_LINK Then
                VerifyDropTarget = 3
            End If
        End If
    End If
End If
out:
If ppidlp Then CoTaskMemFree ppidlp
Exit Function
e0:
    VerifyDropTarget = 0
End Function

Public Function DateToSystemTime(dt As Date) As oleexp.SYSTEMTIME
DateToSystemTime.wDay = CInt(Day(dt))
DateToSystemTime.wMonth = CInt(Month(dt))
DateToSystemTime.wYear = CInt(Year(dt))
DateToSystemTime.wHour = CInt(Hour(dt))
DateToSystemTime.wMinute = CInt(Minute(dt))
DateToSystemTime.wSecond = CInt(Second(dt))

End Function
Private Function PKEY_Size() As oleexp.PROPERTYKEY
Static pkk As oleexp.PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB725F130, &H47EF, &H101A, &HA5, &HF1, &H2, &H60, &H8C, &H9E, &HEB, &HAC, 12)
PKEY_Size = pkk
End Function
Private Function PKEY_DateModified() As oleexp.PROPERTYKEY
Static pkk As oleexp.PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB725F130, &H47EF, &H101A, &HA5, &HF1, &H2, &H60, &H8C, &H9E, &HEB, &HAC, 14)
PKEY_DateModified = pkk
End Function

Private Function TVDragOver(hWnd As Long, ppt As oleexp.POINT, lPrevItem As Long, fValid As Long, fFolder As Long) As Long
Dim tvhti As TVHITTESTINFO

Call ScreenToClient(hWnd, ppt)
tvhti.PT.X = ppt.X
tvhti.PT.Y = ppt.Y

TreeView_HitTest hWnd, tvhti

If (tvhti.Flags And TVHT_ONITEMLINE) Then
    Dim lp As Long
    lp = GetTVItemlParam(hWnd, tvhti.hItem)
    If tvhti.hItem <> lPrevItem Then
            DebugAppend "sel new item, prev=" & lPrevItem & ",cur=" & tvhti.hItem & " item=" & TVEntries(lp).sName & ",dt=" & TVEntries(lp).bDropTarget
            If TVEntries(lp).bDropTarget Then
'                TreeView_SelectItem hwnd, tvhti.hItem
                    If TVEntries(lp).bFolder Then
                        fFolder = 1
                    Else
                        fFolder = 0
                    End If
                    fValid = 1
                    TreeView_SelectDropTarget hTVD, tvhti.hItem
            End If
                TVDragOver = tvhti.hItem
        If (lPrevItem <> -1) And (lPrevItem <> tvhti.hItem) Then
                If TreeView_GetItemState(hWnd, lPrevItem, TVIS_DROPHILITED) Then
                    TreeView_SetItemState hWnd, lPrevItem, 0&, TVIS_DROPHILITED
                    
                End If
        End If
    Else
        TVDragOver = tvhti.hItem
    End If
Else
    If (lPrevItem <> -1) Then
            If TreeView_GetItemState(hWnd, lPrevItem, TVIS_DROPHILITED) Then
                TreeView_SetItemState hWnd, lPrevItem, 0&, TVIS_DROPHILITED
'                TreeView_SetItemState hwnd, lPrevItem, 0&, TVIS_SELECTED
            End If
    End If
    TVDragOver = -1
End If
If (tvhti.Flags And TVHT_ONITEMBUTTON) Then
    DebugAppend "onbutton"
End If
        
End Function

Private Function GetTVItemText(hTVD As Long, _
                                                  hItem As Long, _
                                                  Optional cbItem As Long = MAX_ITEM) As String
  Dim tVI As TVITEM
  
  ' Initialize the struct to retrieve the item's text.
  tVI.Mask = TVIF_TEXT
  tVI.hItem = hItem
  tVI.pszText = StrPtr(String$(cbItem, 0))
  tVI.cchTextMax = cbItem
  
  If TreeView_GetItem(hTVD, tVI) Then
    GetTVItemText = GetStrFromPtrA(tVI.pszText)
  End If

End Function

Private Function GetTVItemPath(hWnd As Long, hItem As Long) As String
Dim lp As Long

lp = GetTVItemlParam(hWnd, hItem)
GetTVItemPath = TVEntries(lp).sFullPath

End Function

Private Function SetTVItemStateImage(hItem As Long, iState As ucst_TVItemCheckStates) As Boolean
  Dim tVI As TVITEM
  
  tVI.Mask = TVIF_HANDLE Or TVIF_STATE
  tVI.hItem = hItem
  tVI.StateMask = TVIS_STATEIMAGEMASK
  tVI.State = IndexToStateImageMask(iState)
  
  SetTVItemStateImage = TreeView_SetItem(hTVD, tVI)
  
End Function

Private Function IndexToStateImageMask(ByVal ImgIndex As Long) As Long
    IndexToStateImageMask = ImgIndex * (2 ^ 12)
End Function

Private Function StateImageMaskToIndex(ByVal ImgState As Long) As Long
    StateImageMaskToIndex = ImgState / (2 ^ 12)
End Function

Private Function INDEXTOOVERLAYMASK(iOverlay As Long) As Long
  INDEXTOOVERLAYMASK = iOverlay * (2 ^ 8)
End Function

Private Function GetFileIconIndexPIDL(pidl As Long, uType As Long) As Long
  Dim sfi As SHFILEINFO
  If SHGetFileInfo(pidl, 0, sfi, Len(sfi), SHGFI_PIDL Or SHGFI_SYSICONINDEX Or uType) Then
    GetFileIconIndexPIDL = sfi.iIcon
  End If
End Function

Public Function GetExpansionState() As String
'Returns a set of folders that if expanded to, will replicate the current appearance of the tree.
'This is *not* a list of all folders, just enough to put the treeview back to the state it's in when this is called.
Dim sData() As String
If CalcRefreshData(sData) Then
    GetExpansionState = Join(sData, "|")
End If

End Function

Public Sub LoadExpansionState(sList As String, Optional bResetFirst As Boolean = True, Optional bExpandItems As Boolean = False)
'Takes a list of directories to expand to, delimited by |
'Generally for the list generated by GetExpansionState
Dim i As Long
Dim sData() As String
If sList = "" Then Exit Sub

sData = Split(sList, "|")

If bResetFirst Then ResetTreeView

For i = 0 To UBound(sData)
    OpenToPath sData(i), bExpandItems
Next i

End Sub
Private Function EnumRoot()
    '<EhHeader>
    On Error GoTo e0
    '</EhHeader>
Dim tvins As TVINSERTSTRUCTEX
Dim tVI As TVITEMEX
Dim siDesk As oleexp.IShellItem
Dim siChild As oleexp.IShellItem
Dim pEnum As oleexp.IEnumShellItems, penum2 As oleexp.IEnumShellItems
Dim upi As oleexp.IParentAndItem
Dim psf As oleexp.IShellFolder
Dim pIcon As oleexp.IShellIcon
Dim pIconOvr As oleexp.IShellIconOverlay
Dim lpIconOvr As Long, lpIconOvr2 As Long
Dim pidlDesktop As Long
Dim pidlPar As Long, pidlRel As Long
Dim lpIcon As Long
Dim pcl As Long, pcl2 As Long
Dim nCur As Long
Dim lAtr As SFGAO_Flags
Dim lpName As Long, sName As String
Dim lpFull As Long, sFull As String
Dim lpNameFull As Long
Dim hr As Long

On Error GoTo e0
If mFavorites Then
nCur = 1&
End If
ReDim TVEntries(0)
fLoad = 1
Dim pKFM As oleexp.KnownFolderManager
Dim pkf As oleexp.IKnownFolder
Dim siFav As oleexp.IShellItem
Dim hFav As Long
Set pKFM = New oleexp.KnownFolderManager
If mFavorites Then
oleexp.SHGetKnownFolderItem FOLDERID_Favorites, KF_FLAG_DEFAULT, 0&, IID_IShellItem, siFav


If (siFav Is Nothing) = False Then
    ReDim Preserve TVEntries(1)
    nCur = 1
    siFav.GetDisplayName SIGDN_NORMALDISPLAY, lpName
    sName = LPWSTRtoStr(lpName)
    DebugAppend "FavName=" & sName
    siFav.GetDisplayName SIGDN_PARENTRELATIVEPARSING, lpNameFull
    TVEntries(1).sNameFull = LPWSTRtoStr(lpNameFull)
    siFav.GetDisplayName SIGDN_DESKTOPABSOLUTEPARSING, lpFull
    sFavPath = LPWSTRtoStr(lpFull)
    TVEntries(1).sFullPath = "1"
    Call SHGetKnownFolderIDList(FOLDERID_Favorites, KF_FLAG_DEFAULT, 0&, pidlDesktop)
    lpIcon = GetFileIconIndexPIDL(pidlDesktop, SHGFI_SMALLICON)
    CoTaskMemFree pidlDesktop
    TVEntries(1).nIcon = lpIcon
    tVI.Mask = TVIF_CHILDREN Or TVIF_TEXT Or TVIF_IMAGE Or TVIF_SELECTEDIMAGE Or TVIF_PARAM Or TVIF_STATE
    tVI.cChildren = 1
    TVEntries(1).sName = sName
    TVEntries(1).bIsDefItem = True
    tVI.cchTextMax = Len(sName)
    tVI.pszText = StrPtr(sName)
    tVI.iImage = lpIcon
    tVI.iSelectedImage = lpIcon
    tVI.StateMask = TVIS_STATEIMAGEMASK
    tVI.State = IndexToStateImageMask(0)
    tVI.lParam = 1
    DebugAppend "FavName=" & TVEntries(1).sName
    tvins.hInsertAfter = 0
    tvins.hParent = TVI_ROOT
    tvins.Item = tVI
    
    m_hFav = SendMessage(hTVD, TVM_INSERTITEMW, 0&, tvins)
    TreeView_SetCheckStateEx hTVD, m_hFav, 0
    TVEntries(1).hNode = m_hFav
    TVEntries(1).hParentNode = -1
    
    siFav.GetAttributes SFGAO_FOLDER Or SFGAO_FILESYSANCESTOR Or SFGAO_STREAM, lAtr
    TVEntries(1).dwAttrib = lAtr
'    TVEntries(1).bFolder = True
    TVEntries(1).bDropTarget = True
    pKFM.GetFolder FOLDERID_Favorites, pkf
    pkf.GetIDList KF_FLAG_DEFAULT, TVEntries(1).pidlFQ
    DebugAppend "hFav=" & m_hFav & ",hTVD=" & hTVD
    TreeView_Expand hTVD, m_hFav, TVE_EXPAND
    UpdateWindow hTVD
        DebugAppend "FavName=" & TVEntries(1).sName

End If
End If
Dim pkfr As oleexp.IKnownFolder
Dim pidlFQR As Long
If mComputerAsRoot Then
    oleexp.SHGetKnownFolderItem FOLDERID_ComputerFolder, KF_FLAG_DEFAULT, 0&, IID_IShellItem, siDesk
    pKFM.GetFolder FOLDERID_ComputerFolder, pkfr
    pkfr.GetIDList KF_FLAG_DEFAULT, pidlFQR
Else
    Set siDesk = Nothing
    If mCustomRoot <> "" Then
        oleexp.SHCreateItemFromParsingName StrPtr(mCustomRoot), Nothing, IID_IShellItem, siDesk
        pidlFQR = ILCreateFromPathW(StrPtr(mCustomRoot))
    End If
    If (siDesk Is Nothing) = False Then
        bCustRt = True
    Else
        Call oleexp.SHCreateItemFromIDList(VarPtr(0&), IID_IShellItem, siDesk)
        pKFM.GetFolder FOLDERID_Desktop, pkfr
        pkfr.GetIDList KF_FLAG_DEFAULT, pidlFQR
    End If
    TVEntries(0).bDropTarget = True
End If
If (siDesk Is Nothing) = False Then
    siDesk.GetDisplayName SIGDN_NORMALDISPLAY, lpName
    sName = LPWSTRtoStr(lpName)
    siDesk.GetDisplayName SIGDN_PARENTRELATIVEPARSING, lpNameFull
    TVEntries(0).sNameFull = LPWSTRtoStr(lpNameFull)
    siDesk.GetDisplayName SIGDN_DESKTOPABSOLUTEPARSING, lpFull
    sFull = LPWSTRtoStr(lpFull)
    If mComputerAsRoot Then
        Call SHGetKnownFolderIDList(FOLDERID_ComputerFolder, KF_FLAG_DEFAULT, 0&, pidlDesktop)
        lpIcon = GetFileIconIndexPIDL(pidlDesktop, SHGFI_SMALLICON)
    Else
        If bCustRt = False Then
            pidlDesktop = GetPIDLFromFolderID(0&, CSIDL_DESKTOP)
            lpIcon = GetFileIconIndexPIDL(pidlDesktop, SHGFI_SMALLICON)
        Else
            pidlDesktop = ILCreateFromPathW(StrPtr(mCustomRoot))
            lpIcon = GetFileIconIndexPIDL(pidlDesktop, SHGFI_SMALLICON)
            Dim pUnk As oleexp.IUnknown
            Set upi = siDesk
            upi.GetParentAndItem pidlPar, psf, pidlRel
            On Error Resume Next
            If (psf Is Nothing) = False Then
                Set pUnk = psf
                hr = pUnk.QueryInterface(IID_IShellIcon, pIcon)
                If hr = S_OK Then
                    pIcon.GetIconOf pidlRel, GIL_FORSHELL, lpIcon
                Else
                  Dim pidlcb As Long
                  pidlcb = ILCombine(pidlPar, pidlRel)
                  lpIcon = GetFileIconIndexPIDL(pidlcb, SHGFI_SMALLICON)
                  DebugAppend "lpIcon on fallback=" & lpIcon
                End If
                If lpIcon = -1 Then lpIcon = 0
            End If
        End If
        sDesktopPath = sFull
    End If
    CoTaskMemFree pidlDesktop
    TVEntries(0).nIcon = lpIcon
    tVI.Mask = TVIF_CHILDREN Or TVIF_TEXT Or TVIF_IMAGE Or TVIF_SELECTEDIMAGE Or TVIF_PARAM Or TVIF_STATE
    tVI.cChildren = 1
    TVEntries(0).sName = sName
    tVI.cchTextMax = Len(sName)
    tVI.pszText = StrPtr(sName)
    tVI.iImage = lpIcon
    tVI.iSelectedImage = lpIcon
    tVI.StateMask = TVIS_STATEIMAGEMASK
    tVI.State = IndexToStateImageMask(0)
    tVI.lParam = 0&
    
    tvins.hInsertAfter = 0
    tvins.hParent = TVI_ROOT
    tvins.Item = tVI
    
    m_hRoot = SendMessage(hTVD, TVM_INSERTITEMW, 0&, tvins)
    If mRootHasCheckbox = False Then
        TreeView_SetCheckStateEx hTVD, m_hRoot, 0
    End If
    TVEntries(0).hNode = m_hRoot
    TVEntries(0).hParentNode = -1
    siDesk.GetAttributes SFGAO_FOLDER Or SFGAO_FILESYSANCESTOR Or SFGAO_STREAM, lAtr
    TVEntries(0).dwAttrib = lAtr
    TVEntries(0).sFullPath = "0"
    TVEntries(0).bFolder = True
    TVEntries(0).pidlFQ = pidlFQR
    TVEntries(0).bIsDefItem = True
    DebugAppend "m_hRoot=" & m_hRoot & ",hTVD=" & hTVD
    TreeView_Expand hTVD, m_hRoot, TVE_EXPAND
    UpdateWindow hTVD
End If


'<EhFooter>
fLoad = 0
Exit Function

e0:
    DebugAppend "ucShellTree.EnumRoot->Error: " & Err.Description & ", 0x" & Hex$(Err.Number)
'</EhFooter>
End Function

Private Function RenameFile(sOld As String, sNewName As String) As Long
DebugAppend "Rename::Enter,old=" & sOld & ",new=" & sNewName
Dim cFO2 As oleexp.FileOperation
Set cFO2 = New oleexp.FileOperation

Dim siItem As oleexp.IShellItem

oleexp.SHCreateItemFromParsingName StrPtr(sOld), Nothing, IID_IShellItem, siItem
If (siItem Is Nothing) = False Then
    cFO2.RenameItem siItem, StrPtr(sNewName), ByVal 0&
    DebugAppend "Rename::SetOps"
    cFO2.SetOperationFlags IFO_ALLOWUNDO
    cFO2.SetOwnerWindow hTVD
    DebugAppend "Rename::Exec"
    cFO2.PerformOperations
    DebugAppend "Rename::CheckAbort"
    cFO2.GetAnyOperationsAborted RenameFile
End If
Set cFO2 = Nothing
Set siItem = Nothing

End Function

Private Function DeleteFile(siFile() As String) As Long
Dim shfos As SHFILEOPSTRUCT
Dim sList As String
Dim i As Long

For i = 0 To UBound(siFile)
    If siFile(i) <> "" Then
        sList = sList & siFile(i) & vbNullChar
    End If
Next i
sList = sList & vbNullChar
DebugAppend "DeleteToBinW->" & sList
With shfos
    .hWnd = hTVD
    .wFunc = FO_DELETE
    .pFrom = StrPtr(sList)
'    .fFlags = FOF_ALLOWUNDO
'    If bDeletePrompt Then
'        .fFlags = .fFlags Or FOF_NOCONFIRMATION
'    End If
End With

DeleteFile = SHFileOperationW(shfos)
End Function

Private Function EmptyTreeView(hTVD As Long, fRedraw As Boolean) As Boolean
    
    ' Prevents any TVN_SELCHANGING/ED processing
    g_fDeleting = True
  
    ' (generates a TVN_DELETEITEM for each added item)
    Call SendMessage(hTVD, WM_SETREDRAW, ByVal 0, 0)
    ' TreeView_DeleteAllItems invokes TVN_SELCHANGING/ED, this doesn't
    EmptyTreeView = TreeView_DeleteItem(hTVD, TreeView_GetRoot(hTVD))
    If m_hFav Then
        TreeView_DeleteItem hTVD, m_hFav
    End If
    If m_hRoot Then
        TreeView_DeleteItem hTVD, m_hRoot
    End If
    
    Call SendMessage(hTVD, WM_SETREDRAW, ByVal 1, 0)
  
    g_fDeleting = False
  
    If fRedraw Then Call UpdateWindow(hTVD)

End Function

Private Sub Hourglass(fOn As Boolean)
If fOn Then
    Call SetCursor(LoadCursor(0, ByVal IDC_WAIT))
Else
    Call SetCursor(LoadCursor(0, ByVal IDC_ARROW))
End If
End Sub

Private Function ExplorerSettingEnabled(lSetting As oleexp.SFS_MASK) As Boolean
Dim lintg As Integer
Call SHGetSettings(lintg, lSetting)
Select Case lSetting
    Case SSF_SHOWALLOBJECTS
        ExplorerSettingEnabled = lintg And 2 ^ 0 'fShowAllObjects
    Case SSF_SHOWEXTENSIONS
        ExplorerSettingEnabled = lintg And 2 ^ 1 'fShowExtensions
    Case SSF_NOCONFIRMRECYCLE
        ExplorerSettingEnabled = lintg And 2 ^ 2 'fNoConfirmRecycly
    Case SSF_SHOWSYSFILES
        ExplorerSettingEnabled = lintg And 2 ^ 3 'fShowSysFiles
    Case SSF_SHOWCOMPCOLOR
        ExplorerSettingEnabled = lintg And 2 ^ 4 'fShowCompColor
    Case SSF_DOUBLECLICKINWEBVIEW
        ExplorerSettingEnabled = lintg And 2 ^ 5 'fDoubleClickInWebView
    Case SSF_DESKTOPHTML
        ExplorerSettingEnabled = lintg And 2 ^ 6 'fDesktopHTML
    Case SSF_WIN95CLASSIC
        ExplorerSettingEnabled = lintg And 2 ^ 7 'fWin95Classic
    Case SSF_DONTPRETTYPATH
        ExplorerSettingEnabled = lintg And 2 ^ 8 'fDontPrettyPath
    Case SSF_SHOWATTRIBCOL
        ExplorerSettingEnabled = lintg And 2 ^ 9 'fShowAttribCol
    Case SSF_MAPNETDRVBUTTON
        ExplorerSettingEnabled = lintg And 2 ^ 10 'fMapNetDrvButton
    Case SSF_SHOWINFOTIP
        ExplorerSettingEnabled = lintg And 2 ^ 11 'fShowInfoTip
    Case SSF_HIDEICONS
        ExplorerSettingEnabled = lintg And 2 ^ 12 'fHideIcons
    Case SSF_SHOWSUPERHIDDEN
        Dim SS As oleexp.SHELLSTATE
        SHGetSetSettings SS, SSF_SHOWSUPERHIDDEN
        ExplorerSettingEnabled = ((SS.fFlags1 And fShowSuperHidden) = fShowSuperHidden)
        
End Select
End Function

Private Function DataObjSupportsFormat(pIDO As oleexp.IDataObject, lFmt As Long, Optional ty As TYMED = TYMED_HGLOBAL, Optional lIndex As Long = -1, Optional dva As DVASPECT = DVASPECT_CONTENT, Optional lDev As Long = 0) As Boolean
Dim tFMT As oleexp.FORMATETC
With tFMT
    .cfFormat = lFmt
    .TYMED = ty
    .dwAspect = dva
    .lIndex = lIndex
    .pDVTARGETDEVICE = lDev
End With
If pIDO.QueryGetData(tFMT) = S_OK Then
    DataObjSupportsFormat = True
End If
End Function

Private Sub TVAddItem(siChild As oleexp.IShellItem, hitemParent As Long, Optional bForceEnable As Boolean = False)
On Error GoTo e0
Dim tvins As TVINSERTSTRUCTEX
Dim tVI As TVITEMEX
Dim siParent As oleexp.IShellItem
'Dim siChild As oleexp.IShellItem
Dim pEnum As oleexp.IEnumShellItems, penum2 As oleexp.IEnumShellItems
Dim upi As oleexp.IParentAndItem
Dim psf As oleexp.IShellFolder
Dim pIcon As oleexp.IShellIcon
Dim pIconOvr As oleexp.IShellIconOverlay
Dim lpIconOvr As Long, lpIconOvr2 As Long
Dim pidlDesktop As Long
Dim pidlPar As Long, pidlRel As Long
Dim lpIcon As Long
Dim pcl As Long, pcl2 As Long
Dim lAtr As SFGAO_Flags
Dim lpName As Long, sName As String
Dim lpFull As Long, sFull As String
Dim lpNameFull As Long, sNameFull As String
Dim idx As Long
Dim fSubFolder As Long
Dim hitemPrev As Long
Dim pUnk As oleexp.IUnknown
Dim bFolder As Boolean, bZip As Boolean
Dim hr As Long
Dim bFav As Boolean
Dim fDisable As Long
Dim pPer As oleexp.IPersistIDList
Dim sFPP As String
Dim i As Long
siChild.GetDisplayName SIGDN_DESKTOPABSOLUTEPARSING, lpFull
sFPP = LPWSTRtoStr(lpFull)
DebugAppend "TVAddItem " & sFPP
'For i = 0 To UBound(TVEntries)
'    If TVEntries(i).sFullPath = sFPP Then Exit Sub
'End If

lpName = 0&: sName = "": lpFull = 0&: sFull = "": lpIcon = 0&: lpIconOvr = 0&: lAtr = 0&: bFolder = False: bZip = False: fDisable = 0&
siChild.GetAttributes SFGAO_CAPABILITYMASK Or SFGAO_CONTENTSMASK Or SFGAO_PKEYSFGAOMASK Or SFGAO_STORAGECAPMASK Or SFGAO_COMPRESSED Or SFGAO_ENCRYPTED Or SFGAO_SHARE, lAtr
If ((lAtr And SFGAO_FOLDER) = SFGAO_FOLDER) Or (mShowFiles = True) Or (bFav = True) Then
    If (((mShowFiles = False) And ((mExpandZip = False) And (lAtr And SFGAO_STREAM) = 0&) Or (mExpandZip = True))) Or (bFav = True) Or (mShowFiles = True) Then 'exlude .zip/.cab
        siChild.GetDisplayName SIGDN_PARENTRELATIVEPARSING, lpNameFull
        sNameFull = LPWSTRtoStr(lpNameFull)
        If (lAtr And SFGAO_FOLDER) = SFGAO_FOLDER Then
            bFolder = True
            If (lAtr And SFGAO_STREAM) = SFGAO_STREAM Then
                bZip = True
            End If
        End If
        If (mFilter <> "*.*") And (mFilter <> "") Then
            If bFolder Then
                If mFilterFilesOnly = False Then
                    If InStr(mFilter, ";") Then
                        If PathMatchSpecExW(StrPtr(sNameFull), StrPtr(mFilter), PMSF_MULTIPLE) = 0& Then Exit Sub
                    Else
                        If PathMatchSpecW(StrPtr(sNameFull), StrPtr(mFilter)) = 0& Then Exit Sub
                    End If
                End If
            Else
                If InStr(mFilter, ";") Then
                    If PathMatchSpecExW(StrPtr(sNameFull), StrPtr(mFilter), PMSF_MULTIPLE) = 0& Then Exit Sub
                Else
                    If PathMatchSpecW(StrPtr(sNameFull), StrPtr(mFilter)) = 0& Then Exit Sub
                End If
            End If
        End If
        nCur = nCur + 1
        ReDim Preserve TVEntries(nCur)
        fSubFolder = 1
        tVI.Mask = TVIF_TEXT Or TVIF_IMAGE Or TVIF_SELECTEDIMAGE Or TVIF_PARAM Or TVIF_STATEEX 'Or TVIF_STATE
        If (((lAtr And SFGAO_FILESYSANCESTOR) = False) And ((lAtr And SFGAO_FILESYSTEM) = False)) Or ((lAtr And SFGAO_NONENUMERATED) = SFGAO_NONENUMERATED) Then
'            If bForceEnable = False Then
'                fDisable = 1&
'            Else
'                fDisable = 0&
'            End If
        End If
        If fDisable = 0& Then
        If (lAtr And SFGAO_HASSUBFOLDER) = SFGAO_HASSUBFOLDER Then
            If (TVEntries(nCur).bZip = False) Or ((TVEntries(nCur).bZip = True) And (mExpandZip = True)) Then
                tVI.cChildren = 1
                tVI.Mask = tVI.Mask Or TVIF_CHILDREN
            End If
        Else
            If (mShowFiles = True) Then
                If bFolder = True Then
                    If FolderIsEmpty(siChild) = False Then
                        tVI.cChildren = 1
                        tVI.Mask = tVI.Mask Or TVIF_CHILDREN
                    End If
                End If
            End If
        End If
        End If
        siChild.GetDisplayName SIGDN_NORMALDISPLAY, lpName
        sName = LPWSTRtoStr(lpName)
        TVEntries(nCur).sFullPath = sFPP
        TVEntries(nCur).sName = sName
        TVEntries(nCur).bFolder = bFolder
        TVEntries(nCur).sParentFull = StrGetPath(TVEntries(nCur).sFullPath)
        TVEntries(nCur).sNameFull = sNameFull
        TVEntries(nCur).bZip = bZip
        TVEntries(nCur).dwAttrib = lAtr
        If (lAtr And SFGAO_DROPTARGET) = SFGAO_DROPTARGET Then
            TVEntries(nCur).bDropTarget = True
        End If
        Set upi = siChild
        upi.GetParentAndItem pidlPar, psf, pidlRel
'                    Set si2Child = siChild
        Set pPer = siChild
        pPer.GetIDList TVEntries(nCur).pidlFQ
        TVEntries(nCur).pidlFQPar = pidlPar
        On Error Resume Next
        If (psf Is Nothing) = False Then
''            DebugAppend "psf is valid, set pIcon with pidlRel=" & pidlRel & ",pidl=" & pidlPar
'            Set pIcon = psf
            Set pUnk = psf
            hr = pUnk.QueryInterface(IID_IShellIcon, pIcon)
'                        DebugAppend "QueryInterface=0x" & Hex$(hr)
            If hr = S_OK Then
                pIcon.GetIconOf pidlRel, GIL_FORSHELL, lpIcon
            Else
              Dim pidlcb As Long
              pidlcb = ILCombine(pidlPar, pidlRel)
'              DebugAppend "pidlcb=" & pidlcb
              lpIcon = GetFileIconIndexPIDL(pidlcb, SHGFI_SMALLICON)
              DebugAppend "lpIcon on fallback=" & lpIcon
            End If
            If lpIcon = -1 Then lpIcon = 0
        End If
        lpIconOvr = -1
        If mExtOverlay Then
            Set pIconOvr = psf
            pIconOvr.GetOverlayIndex pidlRel, VarPtr(lpIconOvr)
        End If
        If (lpIconOvr > 15) Or (lpIconOvr < 0) Then
            'Overlay icons are a mess. On Win7 there's a bunch in root that return 16, which is invalid
            'and will cause a crash later one, and doesn't show anything. Shares never get shown so I'm
            'going to manually set those
            lpIconOvr = -1
            If (lAtr And SFGAO_SHARE) = SFGAO_SHARE Then
                lpIconOvr = 1
            End If
            If (lAtr And SFGAO_LINK) = SFGAO_LINK Then
                lpIconOvr = 2
            End If
        End If
        On Error GoTo e0
        TVEntries(nCur).nIcon = lpIcon
        TVEntries(nCur).nOverlay = lpIconOvr
        TVEntries(nCur).pidlRel = pidlRel
        TVEntries(nCur).hParentNode = hitemParent
        If (lAtr And SFGAO_LINK) = SFGAO_LINK Then
            TVEntries(nCur).bLink = True
            TVEntries(nCur).sLinkTarget = GetLinkTarget(siChild)
            If TVEntries(nCur).sLinkTarget = "" Then TVEntries(nCur).LinkPIDL = GetLinkTargetPIDL(siChild)
            If PathIsDirectoryW(StrPtr(TVEntries(nCur).sLinkTarget)) Then
                TVEntries(nCur).bLinkIsFolder = True
            End If
        End If
        tVI.cchTextMax = Len(TVEntries(nCur).sName)
        tVI.pszText = StrPtr(TVEntries(nCur).sName)
        tVI.iImage = lpIcon
        tVI.iSelectedImage = lpIcon
        tVI.lParam = nCur
        If lpIconOvr > -1 Then
            tVI.Mask = tVI.Mask Or TVIF_STATE
            tVI.StateMask = TVIS_OVERLAYMASK
            tVI.State = INDEXTOOVERLAYMASK(lpIconOvr)
        End If
        
        tvins.hParent = hitemParent
        tvins.hInsertAfter = hitemPrev
        tvins.Item = tVI
        hitemPrev = SendMessage(hTVD, TVM_INSERTITEMW, 0&, tvins)
        DebugAppend "Add@" & nCur & " " & sFPP
        TVEntries(nCur).hNode = hitemPrev
        If TreeView_GetCheckState(hTVD, hitemParent) = 2 Then
            TreeView_SetCheckState hTVD, hitemPrev, 1
        End If

        If fDisable Then
            DebugAppend "TVExpandFolder->Disable " & TVEntries(nCur).sName
            TreeView_SetItemStateEx hTVD, hitemPrev, TVIS_EX_DISABLED
        End If
        
        'Re-sort the new item
        Set psfCur = psf
        Dim tvscb As TVSORTCB
        tvscb.hParent = hitemParent
        If m_cbSort = 0& Then m_cbSort = scb_SetCallbackAddr(3, 2)
        tvscb.lpfnCompare = m_cbSort
        tvscb.lParam = 0&
        Call TreeView_SortChildrenCB(hTVD, tvscb, 0&)
    End If
End If
SetFocus UserControl.ContainerHwnd
SetFocus hTVD
Exit Sub

e0:
    DebugAppend "ucShellTree.TVAddItem->Error: " & Err.Description & ", 0x" & Hex$(Err.Number)
'</EhFooter>
End Sub

Private Sub TVExpandFolder(lParam As Long, hitemParent As Long)
    '<EhHeader>
    On Error GoTo e0
    '</EhHeader>
 DebugAppend "TVExpandFolder()"
Dim tvins As TVINSERTSTRUCTEX
Dim tVI As TVITEMEX
Dim siParent As oleexp.IShellItem
Dim siChild As oleexp.IShellItem
Dim pEnum As oleexp.IEnumShellItems, penum2 As oleexp.IEnumShellItems
Dim upi As oleexp.IParentAndItem
Dim psf As oleexp.IShellFolder
Dim pIcon As oleexp.IShellIcon
Dim pIconOvr As oleexp.IShellIconOverlay
Dim lpIconOvr As Long, lpIconOvr2 As Long
Dim pidlDesktop As Long
Dim pidlPar As Long, pidlRel As Long
Dim lpIcon As Long
Dim pcl As Long, pcl2 As Long
Dim lAtr As SFGAO_Flags
Dim lpName As Long, sName As String
Dim lpFull As Long, sFull As String
Dim lpNameFull As Long, sNameFull As String
Dim idx As Long
Dim fSubFolder As Long
Dim hitemPrev As Long
Dim pUnk As oleexp.IUnknown
Dim bFolder As Boolean, bZip As Boolean
Dim hr As Long
Dim bFav As Boolean
Dim ppidlp As Long
Dim nCount As Long
Dim tc1 As Long, tc2 As Long
On Error GoTo e0
Call Hourglass(True)
idx = lParam
DebugAppend "TVExpandFolder::nCur=" & nCur & "ParentIndex=" & idx & ",lp=" & lParam & ",hItem=" & hitemParent & ",lookup=" & GetIndexFromNode(hitemParent)

If idx = -1 Then Exit Sub

If TVEntries(idx).sFullPath = "0" Then
    If mComputerAsRoot Then
        oleexp.SHGetKnownFolderItem FOLDERID_ComputerFolder, KF_FLAG_DEFAULT, 0&, IID_IShellItem, siParent
    Else
        If mCustomRoot <> "" Then
            oleexp.SHCreateItemFromParsingName StrPtr(mCustomRoot), Nothing, IID_IShellItem, siParent
        End If
        If siParent Is Nothing Then
            Call oleexp.SHCreateItemFromIDList(VarPtr(0&), IID_IShellItem, siParent)
        End If
    End If
ElseIf TVEntries(idx).sFullPath = "1" Then
    oleexp.SHGetKnownFolderItem FOLDERID_Links, KF_FLAG_DEFAULT, 0&, IID_IShellItem, siParent
    bFav = True
Else
    Call oleexp.SHCreateItemFromParsingName(StrPtr(TVEntries(idx).sFullPath), Nothing, IID_IShellItem, siParent)
    If (siParent Is Nothing) Then
        DebugAppend "TVExpandFolder::ParsingPath->ShellItem failed, trying to load from pidls...", 3
        If (TVEntries(idx).pidlFQPar <> 0&) And (TVEntries(idx).pidlRel <> 0&) Then
            ppidlp = ILCombine(TVEntries(idx).pidlFQPar, TVEntries(idx).pidlRel)
            oleexp.SHCreateItemFromIDList ppidlp, IID_IShellItem, siParent
            If (siParent Is Nothing) = False Then
                DebugAppend "TVExpandFolder::LoadFromPidl->Success!", 3
            Else
                DebugAppend "TVExpandFolder::LoadFromPidl->Failed.", 3
            End If
        Else
            DebugAppend "TVExpandFolder::LoadFromPidl->Parent or child pidl not set.", 3
        End If
    End If
End If
Dim fDisable As Long

If (siParent Is Nothing) = False Then
    siParent.BindToHandler 0&, BHID_EnumItems, IID_IEnumShellItems, pEnum
    If (pEnum Is Nothing) = False Then
        bFilling = True
        tc1 = GetTickCount()
        Do While (pEnum.Next(1&, siChild, pcl) = S_OK)
            lpName = 0&: sName = "": lpFull = 0&: sFull = "": lpIcon = 0&: lpIconOvr = 0&: lAtr = 0&: bFolder = False: bZip = False: fDisable = 0&
            siChild.GetAttributes SFGAO_CAPABILITYMASK Or SFGAO_CONTENTSMASK Or SFGAO_PKEYSFGAOMASK Or SFGAO_STORAGECAPMASK Or SFGAO_COMPRESSED Or SFGAO_ENCRYPTED Or SFGAO_SHARE Or SFGAO_HIDDEN Or SFGAO_SYSTEM, lAtr
            If ((lAtr And SFGAO_FOLDER) = SFGAO_FOLDER) Or (mShowFiles = True) Or (bFav = True) Then
                If (((mShowFiles = False) And ((mExpandZip = False) And (lAtr And SFGAO_STREAM) = 0&) Or (mExpandZip = True))) Or (bFav = True) Or (mShowFiles = True) Then 'exlude .zip/.cab
                    If bFav = True Then
                        If (lAtr And SFGAO_LINK) = 0& Then GoTo nxt
                    End If
                    siChild.GetDisplayName SIGDN_PARENTRELATIVEPARSING, lpNameFull
                    sNameFull = LPWSTRtoStr(lpNameFull)
                    If (m_HiddenPref = STHP_UseExplorer) Then
                        If (mHPInExp = False) And ((lAtr And SFGAO_HIDDEN) = SFGAO_HIDDEN) Then GoTo nxt
                    Else
                        If (m_HiddenPref = STHP_AlwaysHide) And ((lAtr And SFGAO_HIDDEN) = SFGAO_HIDDEN) Then GoTo nxt
                    End If
                    If (m_SuperHiddenPref = STSHP_UseExplorer) Then
                        If (mSHPInExp = False) And (((lAtr And SFGAO_HIDDEN) = SFGAO_HIDDEN) And ((lAtr And SFGAO_SYSTEM) = SFGAO_SYSTEM)) Then GoTo nxt
                    Else
                        If (m_SuperHiddenPref = STSHP_AlwaysHide) And (((lAtr And SFGAO_HIDDEN) = SFGAO_HIDDEN) And ((lAtr And SFGAO_SYSTEM) = SFGAO_SYSTEM)) Then GoTo nxt
                    End If

                    If (lAtr And SFGAO_FOLDER) = SFGAO_FOLDER Then
                        bFolder = True
                        If bFav Then GoTo nxt
                        If (lAtr And SFGAO_STREAM) = SFGAO_STREAM Then
                            bZip = True
                        End If
                    End If
                    If (mFilter <> "*.*") And (mFilter <> "") Then
                        If bFav = False Then
                            If bFolder Then
                                If mFilterFilesOnly = False Then
                                    If InStr(mFilter, ";") Then
                                        If PathMatchSpecExW(StrPtr(sNameFull), StrPtr(mFilter), PMSF_MULTIPLE) = 0& Then GoTo nxt
                                    Else
                                        If PathMatchSpecW(StrPtr(sNameFull), StrPtr(mFilter)) = 0& Then GoTo nxt
                                    End If
                                End If
                            Else
                                If InStr(mFilter, ";") Then
                                    If PathMatchSpecExW(StrPtr(sNameFull), StrPtr(mFilter), PMSF_MULTIPLE) = 0& Then GoTo nxt
                                Else
                                    If PathMatchSpecW(StrPtr(sNameFull), StrPtr(mFilter)) = 0& Then GoTo nxt
                                End If
                            End If
                        End If
                    End If
                        
                    nCur = nCur + 1
                    ReDim Preserve TVEntries(nCur)
                    fSubFolder = 1
                    tVI.Mask = TVIF_TEXT Or TVIF_IMAGE Or TVIF_SELECTEDIMAGE Or TVIF_PARAM Or TVIF_STATEEX 'Or TVIF_STATE
                    If (((lAtr And SFGAO_FILESYSANCESTOR) = False) And ((lAtr And SFGAO_FILESYSTEM) = False)) Or ((lAtr And SFGAO_NONENUMERATED) = SFGAO_NONENUMERATED) Then
'                        If (lAtr And SFGAO_DROPTARGET) = 0& Then
'                            fDisable = 1&
'                            TVEntries(nCur).bDisabled = True
'                        End If
                    End If
                    If fDisable = 0& Then
                    If (lAtr And SFGAO_HASSUBFOLDER) = SFGAO_HASSUBFOLDER Then
                        If (TVEntries(nCur).bZip = False) Or ((TVEntries(nCur).bZip = True) And (mExpandZip = True)) Then
                            tVI.cChildren = 1
                            tVI.Mask = tVI.Mask Or TVIF_CHILDREN
                        End If
                    Else
                        If (mShowFiles = True) Then
                            If bFolder = True Then
                                If FolderIsEmpty(siChild) = False Then
                                    tVI.cChildren = 1
                                    tVI.Mask = tVI.Mask Or TVIF_CHILDREN
                                End If
                            End If
                        End If
                    End If
                    End If
                    siChild.GetDisplayName SIGDN_NORMALDISPLAY, lpName
                    sName = LPWSTRtoStr(lpName)
                    If nCur = 2 Then
                        DebugAppend "nCur=2,name=" & sName
                    End If
'                    DebugAppend sName & " atr=" & dbg_sfgao_tostring(lAtr)
                    siChild.GetDisplayName SIGDN_DESKTOPABSOLUTEPARSING, lpFull
                    TVEntries(nCur).sFullPath = LPWSTRtoStr(lpFull)
                    TVEntries(nCur).sName = sName
                    TVEntries(nCur).bFolder = bFolder
                    TVEntries(nCur).sParentFull = StrGetPath(TVEntries(nCur).sFullPath)
                    TVEntries(nCur).sNameFull = sNameFull
                    TVEntries(nCur).bZip = bZip
                    TVEntries(nCur).dwAttrib = lAtr
                    If (lAtr And SFGAO_DROPTARGET) = SFGAO_DROPTARGET Then
                        TVEntries(nCur).bDropTarget = True
                    End If
                    Set upi = siChild
                    upi.GetParentAndItem pidlPar, psf, pidlRel
'                    Set si2Child = siChild
                    On Error Resume Next
                    If (psf Is Nothing) = False Then
            ''            DebugAppend "psf is valid, set pIcon with pidlRel=" & pidlRel & ",pidl=" & pidlPar
            '            Set pIcon = psf
                        Set pUnk = psf
                        hr = pUnk.QueryInterface(IID_IShellIcon, pIcon)
'                        DebugAppend "QueryInterface=0x" & Hex$(hr)
                        If hr = S_OK Then
                            pIcon.GetIconOf pidlRel, GIL_FORSHELL, lpIcon
                        Else
                          Dim pidlcb As Long
                          pidlcb = ILCombine(pidlPar, pidlRel)
            '              DebugAppend "pidlcb=" & pidlcb
                          lpIcon = GetFileIconIndexPIDL(pidlcb, SHGFI_SMALLICON)
                          DebugAppend "lpIcon on fallback=" & lpIcon
                        End If
                        If lpIcon = -1 Then lpIcon = 0
                    End If
                    lpIconOvr = -1
                    If mExtOverlay Then
                        Set pIconOvr = psf
                        pIconOvr.GetOverlayIndex pidlRel, VarPtr(lpIconOvr)
                    End If
                    If (lpIconOvr > 15) Or (lpIconOvr < 0) Then
                        'Overlay icons are a mess. On Win7 there's a bunch in root that return 16, which is invalid
                        'and will cause a crash later one, and doesn't show anything. Shares never get shown so I'm
                        'going to manually set those
                        lpIconOvr = -1
                        If (lAtr And SFGAO_SHARE) = SFGAO_SHARE Then
                            lpIconOvr = 1
                        End If
                        If (lAtr And SFGAO_LINK) = SFGAO_LINK Then
                            lpIconOvr = 2
                        End If
                    End If
                    On Error GoTo e0
                    TVEntries(nCur).nIcon = lpIcon
                    TVEntries(nCur).nOverlay = lpIconOvr
                    TVEntries(nCur).pidlFQPar = pidlPar
                    TVEntries(nCur).pidlRel = pidlRel
                    TVEntries(nCur).hParentNode = hitemParent
                    If fLoad Then
                        TVEntries(nCur).bIsDefItem = True
                    End If
                    If (lAtr And SFGAO_LINK) = SFGAO_LINK Then
                        TVEntries(nCur).bLink = True
                        TVEntries(nCur).sLinkTarget = GetLinkTarget(siChild)
                        If TVEntries(nCur).sLinkTarget = "" Then TVEntries(nCur).LinkPIDL = GetLinkTargetPIDL(siChild)
                        If PathIsDirectoryW(StrPtr(TVEntries(nCur).sLinkTarget)) Then
                            TVEntries(nCur).bLinkIsFolder = True
                        End If
                        tVI.Mask = tVI.Mask Or TVIF_STATE
                    End If
                    tVI.cchTextMax = Len(TVEntries(nCur).sName)
                    tVI.pszText = StrPtr(TVEntries(nCur).sName)
                    tVI.iImage = lpIcon
                    tVI.iSelectedImage = lpIcon
                    tVI.lParam = nCur
                    If lpIconOvr > -1 Then
                        tVI.Mask = tVI.Mask Or TVIF_STATE
                        tVI.StateMask = TVIS_OVERLAYMASK
                        tVI.State = INDEXTOOVERLAYMASK(lpIconOvr)
                    End If
                    tvins.hParent = hitemParent
                    tvins.hInsertAfter = hitemPrev
                    tvins.Item = tVI
                    
                    hitemPrev = SendMessage(hTVD, TVM_INSERTITEMW, 0&, tvins)
                    TVEntries(nCur).hNode = hitemPrev
                    nCount = nCount + 1
                    
                    If bFav Then
                        TreeView_SetCheckStateEx hTVD, hitemPrev, 0
                    Else
                        If TreeView_GetCheckState(hTVD, hitemParent) = 2 Then
                            TreeView_SetCheckState hTVD, hitemPrev, 1
'                                    TreeView_SetItemState hTVD, hitemPrev, IndexToStateImageMask(3), TVIS_STATEIMAGEMASK
                        End If
                    End If
                    If fDisable Then
                        DebugAppend "TVExpandFolder->Disable " & TVEntries(nCur).sName & ",atr=" & dbg_sfgao_tostring(lAtr)
                        TreeView_SetItemStateEx hTVD, hitemPrev, TVIS_EX_DISABLED
'                        TreeView_SetCheckStateEx hTVD, hitemPrev, 3
                    End If
                End If
            End If
nxt:
        Loop
        tc2 = GetTickCount()
        DebugAppend "TVExpandFolder added " & nCount & " items in " & (tc2 - tc1) & "ms", 2
        bFilling = False
        Set psfCur = psf
    End If
End If
If ppidlp Then CoTaskMemFree ppidlp
If fSubFolder = 0& Then 'false alarm; doesn't actually have em
    tVI.hItem = hitemParent
    tVI.Mask = TVIF_CHILDREN
    Call SendMessage(hTVD, TVM_GETITEM, 0, tVI)
    tVI.cChildren = 0
    Call SendMessage(hTVD, TVM_SETITEM, 0, tVI)
End If
Call Hourglass(False)
SetFocus UserControl.ContainerHwnd
SetFocus hTVD
                    
'<EhFooter>
Exit Sub

e0:
    DebugAppend "ucShellTree.TVExpandFolder->Error: " & Err.Description & ", 0x" & Hex$(Err.Number)
'</EhFooter>
End Sub

Private Function GetIndexFromNode(hNode As Long) As Long
Dim i As Long
GetIndexFromNode = -1
For i = 0 To UBound(TVEntries)
    If TVEntries(i).hNode = hNode Then
        GetIndexFromNode = i
        Exit Function
    End If
Next i
End Function

Private Function GetIndexFromPath(sPath As String) As Long
Dim i As Long
GetIndexFromPath = -1
For i = 0 To UBound(TVEntries)
    If TVEntries(i).sFullPath = sPath Then
        GetIndexFromPath = i
        Exit Function
    End If
Next i
End Function

Private Function GetIndexFromPathND(sPath As String) As Long
Dim i As Long
GetIndexFromPathND = -1
For i = 0 To UBound(TVEntries)
    If TVEntries(i).sFullPath = sPath Then
        If TVEntries(i).bDeleted = False Then
            GetIndexFromPathND = i
            Exit Function
        End If
    End If
Next i
End Function

Private Function GetSelectedTVItem(hTVD As Long, tVI As TVITEM) As Long
    tVI.hItem = TreeView_GetSelection(hTVD)
    tVI.Mask = TVIF_TEXT Or TVIF_PARAM Or TVIF_IMAGE Or TVIF_STATE Or TVIF_HANDLE Or TVIF_SELECTEDIMAGE Or TVIF_CHILDREN
    tVI.pszText = StrPtr(String$(MAX_PATH, 0))
    tVI.cchTextMax = MAX_PATH
    GetSelectedTVItem = TreeView_GetItem(hTVD, tVI)
End Function

Private Function LoWord(ByVal DWord As Long) As Integer
If DWord And &H8000& Then
    LoWord = DWord Or &HFFFF0000
Else
    LoWord = DWord And &HFFFF&
End If
End Function

Private Function CreateTreeView(hwndParent As Long, _
                                                    iid As Long, _
                                                    dwStyle As Long, _
                                                    dwExStyle As Long) As Long
  Dim rc As oleexp.RECT ' parent client rect
  Call GetClientRect(hwndParent, rc)
  ' Create the TreeView control.
  CreateTreeView = CreateWindowEx(dwExStyle, WC_TREEVIEW, "", _
                                                dwStyle, 0, 0, rc.Right, rc.Bottom, _
                                            hwndParent, iid, App.hInstance, 0)

End Function

Private Sub InitTV()
Dim dwStyle As ucst_TV_Styles '
Dim dwExStyle As WindowStylesEx
'If mBorder Then 'The old single border option in case you want to restore it
'    dwStyle = dwStyle Or WS_BORDER
'End If          'The old Property Get/Let is there and commented out as well.
Select Case mBorder
    Case STBS_Standard
        dwStyle = dwStyle Or WS_BORDER
    Case STBS_Thick
        dwStyle = dwStyle Or WS_BORDER
        dwExStyle = WS_EX_CLIENTEDGE
    Case STBS_Thicker
        dwStyle = dwStyle Or WS_BORDER Or WS_THICKFRAME
        dwExStyle = WS_EX_CLIENTEDGE
End Select
If mShowSelAlw Then
    dwStyle = dwStyle Or TVS_SHOWSELALWAYS
End If
If m_TrackSel Then
    dwStyle = dwStyle Or TVS_TRACKSELECT
End If
If mLabelEdit Then
    dwStyle = dwStyle Or TVS_EDITLABELS
End If
If mHasButtons Then
    dwStyle = dwStyle Or TVS_HASBUTTONS
End If
If mShowLines Then
    dwStyle = dwStyle Or TVS_HASLINES
End If
If mSingleExpand Then
    dwStyle = dwStyle Or TVS_SINGLEEXPAND
End If
If mHScroll = False Then
    dwStyle = dwStyle Or TVS_NOHSCROLL
End If
If mFullRowSelect Then
    dwStyle = dwStyle Or TVS_FULLROWSELECT
End If
If ((mCheckboxes = True) And (mAutocheck = False) And (mExCheckboxes = False)) Or ((mCheckboxes = True) And (IsComCtl6 = False)) Then
    dwStyle = dwStyle Or TVS_CHECKBOXES
End If
If (mInfoTipOnFiles = True) Or (mInfoTipOnFolders = True) Then
    dwStyle = dwStyle Or TVS_INFOTIP
End If
  hTVD = CreateTreeView(UserControl.hWnd, _
                                              IDD_TREEVIEW, _
                                              dwStyle Or _
                                              WS_TABSTOP Or WS_VISIBLE Or WS_CHILD, _
                                              dwExStyle)

m_hWnd = hTVD
Set Me.Font = PropFont
Me.BackColor = clrBack
Me.ForeColor = clrFore

If Ambient.UserMode Then
    DebugAppend "Subclass TV, h=" & hTVD
    If ssc_Subclass(hTVD, , , , , True, True) Then
        Call ssc_AddMsg(hTVD, MSG_BEFORE, ALL_MESSAGES)
    End If
End If

DebugAppend "frmTV.InitTV::hWnd=" & hTVD, 3
Dim dwStyleEx As ucst_TV_Ex_Styles

dwStyleEx = SendMessage(hTVD, TVM_GETEXTENDEDSTYLE, 0, ByVal 0&)
dwStyleEx = TVS_EX_DOUBLEBUFFER Or &H1000
If mFadeExpandos Then
    dwStyleEx = dwStyleEx Or TVS_EX_FADEINOUTEXPANDOS
End If
If mMultiSel Then
    dwStyleEx = dwStyleEx Or TVS_EX_MULTISELECT
End If
If mNoIndState Then
    dwStyleEx = dwStyleEx Or TVS_EX_NOINDENTSTATE
End If
If mAutoHS Then
    dwStyleEx = dwStyleEx Or TVS_EX_AUTOHSCROLL
End If
If IsComCtl6 Then
    If (mCheckboxes = True) And (mAutocheck = True) Then
        dwStyleEx = dwStyleEx Or TVS_EX_PARTIALCHECKBOXES
    End If
    If (mCheckboxes = True) And (mExCheckboxes = True) Then
        DebugAppend "SET TVS_EX_ECB"
        dwStyleEx = dwStyleEx Or TVS_EX_EXCLUSIONCHECKBOXES
    End If
End If
Call SendMessage(hTVD, TVM_SETEXTENDEDSTYLE, dwStyleEx, ByVal dwStyleEx)

If mExplorerStyle Then
    Dim sTheme As String
    sTheme = "explorer"
    Call SetWindowTheme(hTVD, StrPtr(sTheme), 0&)
End If

Call SHGetImageList(SHIL_SYSSMALL, IID_IImageList, pIML)
hSysIL = ObjPtr(pIML)
Call SendMessage(hTVD, TVM_SETIMAGELIST, TVSIL_NORMAL, ByVal hSysIL)

'TreeView_SetImageList hTVD, himlTVCheck, TVSIL_STATE 'if you want to go back to custom checks

End Sub

Private Function TrimNullW(startstr As String) As String

   TrimNullW = Left$(startstr, lstrlenW(ByVal StrPtr(startstr)))
   
End Function

Private Function Attach(hWnd As Long) As Long
If bRegDD = False Then
    Attach = RegisterDragDrop(hWnd, Me)
    bRegDD = True
End If
End Function

Private Function Detach() As Long
'There's an appcrash if revoke is called on an unregistered window
If bRegDD Then
    If RegisterDragDrop(m_hWnd, Me) = DRAGDROP_E_ALREADYREGISTERED Then
        Detach = RevokeDragDrop(m_hWnd)
        bRegDD = False
    End If
End If
End Function

Private Function FolderIsEmpty(sItem As oleexp.IShellItem) As Boolean
Dim pEnum As oleexp.IEnumShellItems
Dim pc As Long
Dim sTmp As oleexp.IShellItem
sItem.BindToHandler 0&, BHID_EnumItems, IID_IEnumShellItems, pEnum
If (pEnum Is Nothing) Then
    FolderIsEmpty = True
    Exit Function
Else
    If (pEnum.Next(1&, sTmp, pc) <> S_OK) Then
        FolderIsEmpty = True
        Exit Function
    End If
End If
    
End Function

Private Function GenerateInfoTip(iItem As Long) As String
'DebugAppend "GenerateInfoTip item=" & iItem & ",name=" & TVEntries(iItem).sName
Dim si As oleexp.IShellItem
Dim sTip As String
Dim ppidlp As Long
On Error GoTo e0

oleexp.SHCreateItemFromParsingName StrPtr(TVEntries(iItem).sFullPath), Nothing, IID_IShellItem, si
If (si Is Nothing) Then
    DebugAppend "GenerateInfoTip->ShellItem failed, trying to load from pidls...", 3
    If (TVEntries(iItem).pidlFQPar <> 0&) And (TVEntries(iItem).pidlRel <> 0&) Then
        ppidlp = ILCombine(TVEntries(iItem).pidlFQPar, TVEntries(iItem).pidlRel)
        oleexp.SHCreateItemFromIDList ppidlp, IID_IShellItem, si
        If (si Is Nothing) = False Then
            DebugAppend "GenerateInfoTip->LoadFromPidl->Success!", 3
        Else
            DebugAppend "GenerateInfoTip->LoadFromPidl->Failed.", 3
        End If
    Else
        DebugAppend "GenerateInfoTip->LoadFromPidl->Parent or child pidl not set.", 3
    End If
    If ppidlp Then CoTaskMemFree ppidlp
End If

If (si Is Nothing = False) Then
    Dim pqi As oleexp.IQueryInfo
    si.BindToHandler 0&, BHID_SFUIObject, IID_IQueryInfo, pqi
    If (pqi Is Nothing) = False Then
        Dim lpTip As Long, sQITip As String
        pqi.GetInfoTip QITIPF_USESLOWTIP, lpTip
        sQITip = LPWSTRtoStr(lpTip)
'        DebugAppend "IQueryInfo.Slo=" & sQITip
        GenerateInfoTip = sQITip
        Exit Function
    Else
        DebugAppend "Failed to get IQueryInfo"
    End If
    
    Dim lpp As Long
    Dim si2p As oleexp.IShellItem2
    Dim pl As oleexp.IPropertyDescriptionList
    Dim pd As oleexp.IPropertyDescription
    Dim lpn As Long, sPN As String
    
    Set si2p = si
    Dim pst As oleexp.IPropertyStore
    si2p.GetPropertyDescriptionList PKEY_PropList_InfoTip, IID_IPropertyDescriptionList, pl
    If (pl Is Nothing) = False Then
        pl.GetCount lpp
    '    DebugAppend "InfoTip Cnt=" & lpp
        If lpp Then
            Dim stt As String
            si2p.GetPropertyStore GPS_BESTEFFORT Or GPS_OPENSLOWITEM, IID_IPropertyStore, pst
            If (pst Is Nothing) = False Then
    '            DebugAppend "PropsList=" & GetPropertyKeyDisplayString(pst, PKEY_PropList_InfoTip)
                'We could just parse that; but going through IPropertyDescriptionList automatically skips
                'fields where there's no data (an error is raised, hence the e1/resume next)GPS_DEFAULT Or
                On Error GoTo e1
                Dim i As Long
                For i = 0 To (lpp - 1)
                    pl.GetAt i, IID_IPropertyDescription, pd
                    If (pd Is Nothing) = False Then
                        stt = GetPropertyDisplayString(pst, pd)
                        If stt <> "" Then 'well, it used to generate an error??? then suddenly stopped 10 min later!?!?
                            pd.GetDisplayName lpn
                            sPN = LPWSTRtoStr(lpn)
                            stt = sPN & ": " & stt
                            If sTip = "" Then
                                sTip = stt
                            Else
                                sTip = sTip & vbCrLf & stt
                            End If
    '                        DebugAppend "Prop=" & stt
                            stt = ""
                        End If
                        Set pd = Nothing
                    Else
    '                    DebugAppend "Prop=(missing)"
                    End If
                Next i
                Set pst = Nothing
            End If
        Else
            DebugAppend "lpp=" & lpp
        End If
    Else
        DebugAppend "No proplist"
    End If
Else
    DebugAppend "No IShellItem"
End If
GenerateInfoTip = sTip
Exit Function
e0:
DebugAppend "GenerateInfoTip->Error: " & Err.Description & ", 0x" & Hex$(Err.Number)
Exit Function
e1:
DebugAppend "GenerateInfoTip->Error: " & Err.Description & ", 0x" & Hex$(Err.Number)
Resume Next
End Function

Private Function GetPropertyDisplayString(pps As oleexp.IPropertyStore, ppd As oleexp.IPropertyDescription, Optional bFixChars As Boolean = True) As String
'Same as above if you already have the IPropertyDescription (caller is responsible for freeing it too)
Dim lpsz As Long
On Error GoTo e0
If ((pps Is Nothing) = False) And ((ppd Is Nothing) = False) Then
    PSFormatPropertyValue ObjPtr(pps), ObjPtr(ppd), PDFF_DEFAULT, lpsz
    SysReAllocString VarPtr(GetPropertyDisplayString), lpsz
    CoTaskMemFree lpsz
    If bFixChars Then
        GetPropertyDisplayString = Replace$(GetPropertyDisplayString, ChrW$(&H202A), "")
        GetPropertyDisplayString = Replace$(GetPropertyDisplayString, ChrW$(&H202C), "")
        GetPropertyDisplayString = Replace$(GetPropertyDisplayString, ChrW$(&H200E), "")
        GetPropertyDisplayString = Replace$(GetPropertyDisplayString, ChrW$(&H200F), "")
    End If
Else
    DebugAppend "GetPropertyDisplayString.Error->PropertyStore or PropertyDescription is not set."
    
End If
Exit Function
e0:
DebugAppend "GetPropertyDisplayString.Error->" & Err.Description & ", 0x" & Hex$(Err.Number)
End Function

Private Function GetPreferredEffect(pDataObj As oleexp.IDataObject) As DROPEFFECTS
Dim tSTG As oleexp.STGMEDIUM
Dim tFMT As oleexp.FORMATETC
Dim lpGlobal As Long
Dim hGlobal As Long
Dim lEf As Long

    tFMT.cfFormat = CF_PREFERREDDROPEFFECT
    tFMT.dwAspect = DVASPECT_CONTENT
    tFMT.lIndex = -1
    tFMT.TYMED = TYMED_HGLOBAL

    pDataObj.GetData tFMT, tSTG
    lpGlobal = GlobalLock(tSTG.data)
    CopyMemory lEf, ByVal lpGlobal, 4&
    Call GlobalUnlock(tSTG.data)
    ReleaseStgMedium tSTG
    
    GetPreferredEffect = lEf
End Function

Private Sub TVDoPaste()
'Handle a paste command
Dim pdo As oleexp.IDataObject
Dim lEfct As DROPEFFECTS
Dim siDrop As oleexp.IShellItem
On Error GoTo e0
OleGetClipboard pdo
If (pdo Is Nothing) Then Exit Sub
If DataObjSupportsFormat(pdo, CF_HDROP) Then 'clipboard has files for us
    If GetPreferredEffect(pdo) = DROPEFFECT_MOVE Then
        DebugAppend "Paste as move"
        lEfct = DROPEFFECT_MOVE
    Else
        lEfct = DROPEFFECT_COPY
    End If
    If ((siSelected Is Nothing) = False) Then
        Dim lAtr As SFGAO_Flags
        siSelected.GetAttributes SFGAO_FOLDER, lAtr
        If (lAtr And SFGAO_FOLDER) = SFGAO_FOLDER Then
            Set siDrop = siSelected
        End If
    End If
 
    Dim pDT As oleexp.IDropTarget
    siDrop.BindToHandler 0&, BHID_SFUIObject, IID_IDropTarget, pDT
    If (pDT Is Nothing) = False Then
        Dim lpName As Long
        siDrop.GetDisplayName SIGDN_NORMALDISPLAY, lpName
        DebugAppend "Go for paste on " & LPWSTRtoStr(lpName)
        pDT.DragEnter pdo, MK_LBUTTON, 0&, 0&, lEfct
        pDT.Drop pdo, MK_LBUTTON, 0&, 0&, lEfct
    Else
        Beep
        DebugAppend "Fail DropTarget"
    End If
 
End If
Exit Sub
e0:
    DebugAppend "TVDoPaste.Error->" & Err.Description & ", 0x" & Hex$(Err.Number)
End Sub

Private Sub TVCutSelected()
Dim si As Long

si = TreeView_GetSelection(hTVD)
TreeView_SetItemState hTVD, si, TVIS_CUT, TVIS_CUT
End Sub

Private Function ShowShellContextMenu(lp As Long) As Long
Dim cnt As Long
If m_EnableShellMenu = False Then Exit Function
'Cnt = SendMessage(hTVD, TVM_GETSELECTEDCOUNT, 0&, ByVal 0&)
If lp < 1 Then
'    ShowViewMenu True
    Exit Function
End If
'If cnt = 0 Then Exit Function
Dim i As Long
'Dim lp As Long
Dim apidl() As Long
Dim cpidl As Long
Dim pcm As oleexp.IContextMenu
Dim cmi As oleexp.CMINVOKECOMMANDINFO
Dim PT As oleexp.POINT
Dim idCmd As Long
Dim hMenu As Long
Dim uFlag As QueryContextMenuFlags
Dim mii As MENUITEMINFOW
Dim psfPar As oleexp.IShellFolder
Dim psi As oleexp.IShellItem
Dim ppidlp As Long
Const sSel As String = "Select"

ReDim apidl(0)

apidl(cpidl) = TVEntries(lp).pidlRel
cpidl = cpidl + 1

Dim upi As oleexp.IParentAndItem
Dim pidl As Long, pidlRel As Long
DebugAppend "ShowShellContextMenu->Parse to isi " & TVEntries(lp).sFullPath
oleexp.SHCreateItemFromParsingName StrPtr(TVEntries(lp).sFullPath), Nothing, IID_IShellItem, psi
If (psi Is Nothing) Then
    DebugAppend "ShowShellContextMenu->ShellItem failed, trying to load from pidls...", 3
    If (TVEntries(lp).pidlFQPar <> 0&) And (TVEntries(lp).pidlRel <> 0&) Then
        ppidlp = ILCombine(TVEntries(lp).pidlFQPar, TVEntries(lp).pidlRel)
        oleexp.SHCreateItemFromIDList ppidlp, IID_IShellItem, psi
        If (psi Is Nothing) = False Then
            DebugAppend "ShowShellContextMenu::LoadFromPidl->Success!", 3
        Else
            DebugAppend "ShowShellContextMenu::LoadFromPidl->Failed.", 3
        End If
    Else
        DebugAppend "ShowShellContextMenu::LoadFromPidl->Parent or child pidl not set.", 3
    End If
    If ppidlp Then CoTaskMemFree ppidlp
End If
If (psi Is Nothing) = False Then
    Set upi = psi
Else
    DebugAppend "ShowShellContextMenu->Failed to get IShellItem"
    Exit Function
End If
    
If (upi Is Nothing) = False Then
    upi.GetParentAndItem pidl, psfPar, pidlRel
Else
    DebugAppend "ShowShellContextMenu->Failed to get ParentAndItemCMF_EXPLORE Or "
    Exit Function
End If

psfPar.GetUIObjectOf hTVD, cpidl, apidl(0), IID_IContextMenu, 0&, pcm
uFlag = CMF_NODEFAULT Or CMF_CANRENAME Or CMF_ITEMMENU

If mAlwaysShowExtVerbs Then
    uFlag = uFlag Or CMF_EXTENDEDVERBS
Else
    If (GetKeyState(VK_SHIFT) And &H80) = &H80 Then
        uFlag = uFlag Or CMF_EXTENDEDVERBS
    End If
End If
    
If (pcm Is Nothing) = False Then
    hMenu = CreatePopupMenu()
    Dim pUnk As oleexp.IUnknown
    Set pUnk = pcm
    pUnk.QueryInterface IID_IContextMenu2, ICtxMenu2
    pUnk.QueryInterface IID_IContextMenu3, ICtxMenu3
    If (ICtxMenu3 Is Nothing) = False Then
        ICtxMenu3.QueryContextMenu hMenu, 0&, 1&, &H7FFF, uFlag
    Else
        If (ICtxMenu2 Is Nothing) = False Then
            ICtxMenu2.QueryContextMenu hMenu, 0&, 1&, &H7FFF, uFlag
        Else
            pcm.QueryContextMenu hMenu, 0&, 1&, &H7FFF, uFlag
        End If
    End If
    mii.cbSize = Len(mii)
    mii.fMask = MIIM_TYPE
    mii.fType = MFT_SEPARATOR
    InsertMenuItemW hMenu, 0&, True, mii
    
    mii.fMask = MIIM_STRING Or MIIM_ID Or MIIM_STATE
    mii.wID = wIDSel
    mii.fState = MFS_DEFAULT
    mii.cch = Len(sSel)
    mii.dwTypeData = StrPtr(sSel)
    InsertMenuItemW hMenu, 0&, True, mii
     
    Call GetCursorPos(PT)
    
    idCmd = TrackPopupMenu(hMenu, TPM_LEFTBUTTON Or TPM_RIGHTBUTTON Or TPM_LEFTALIGN Or TPM_TOPALIGN Or TPM_HORIZONTAL Or TPM_RETURNCMD, PT.X, PT.Y, 0&, hTVD, 0&)
   
    If idCmd Then
        If idCmd = wIDSel Then
            RaiseEvent ItemSelect(TVEntries(lp).sName, TVEntries(lp).sFullPath, TVEntries(lp).bFolder, TVEntries(lp).hNode)
            RaiseEvent ItemSelectByShellItem(psi, TVEntries(lp).sName, TVEntries(lp).sFullPath, TVEntries(lp).bFolder, TVEntries(lp).hNode)
        Else
            Dim sVerb As String
            sVerb = String$(MAX_PATH, 0&)
            On Error Resume Next
            pcm.GetCommandString idCmd - 1, GCS_VERBW, 0&, StrPtr(sVerb), Len(sVerb)
            sVerb = LCase$(TrimNullW(sVerb))
            DebugAppend "ShellContextMenu->Verb=" & sVerb
            On Error GoTo e0
            If sVerb = "rename" Then
                DebugAppend "ShellContextMenu->Rename"
                Call SendMessage(hTVD, TVM_EDITLABELW, 0&, ByVal TVEntries(lp).hNode)
            Else
                With cmi
                    .cbSize = Len(cmi)
                    .hWnd = hTVD
                    .lpVerb = idCmd - 1 ' MAKEINTRESOURCE(idCmd-1);
                    .nShow = SW_SHOWNORMAL
                End With
                If (ICtxMenu2 Is Nothing) = False Then
                    ICtxMenu2.InvokeCommand VarPtr(cmi)
                Else
                    pcm.InvokeCommand VarPtr(cmi)
                End If
                If sVerb = "cut" Then TVCutSelected
            End If
        End If
    End If
Else
    DebugAppend "ShowShellContextMenu->Failed to get IContextMenu"
End If
DestroyMenu hMenu
Set ICtxMenu2 = Nothing
Set pcm = Nothing
Exit Function
e0:
DebugAppend "ShowShellContextMenu->Error: " & Err.Description & ", 0x" & Hex$(Err.Number)
End Function

Private Function StartNotify(hWnd As Long, Optional pidlPath As Long = 0) As Long
  Dim tCNE As oleexp.SHChangeNotifyEntry
  Dim pidl As Long
  
  If (m_hSHNotify = 0) Then
        If pidlPath = 0 Then
'            Call SHGetSpecialFolderLocation(0, CSIDL_DESKTOP, pidl)
            tCNE.pidl = VarPtr(0) 'pidl '
        Else
            tCNE.pidl = pidlPath
        End If
      tCNE.fRecursive = 1
      
      m_hSHNotify = oleexp.SHChangeNotifyRegister(hWnd, SHCNRF_ShellLevel Or SHCNRF_InterruptLevel Or SHCNRF_NewDelivery, SHCNE_ALLEVENTS Or SHCNE_INTERRUPT, WM_SHNOTIFY, 1, tCNE)
      
      StartNotify = m_hSHNotify
        
  End If   ' (m_hSHNotify = 0)

End Function

Private Function StopNotify() As Boolean
StopNotify = SHChangeNotifyDeregister(m_hSHNotify)
End Function

Private Function dbg_LookUpSHCNE(uMsg As Long) As String

Select Case uMsg

Case &H1: dbg_LookUpSHCNE = "SHCNE_RENAMEITEM"
Case &H2: dbg_LookUpSHCNE = "SHCNE_CREATE"
Case &H4: dbg_LookUpSHCNE = "SHCNE_DELETE"
Case &H8: dbg_LookUpSHCNE = "SHCNE_MKDIR"
Case &H10: dbg_LookUpSHCNE = "SHCNE_RMDIR"
Case &H20: dbg_LookUpSHCNE = "SHCNE_MEDIAINSERTED"
Case &H40: dbg_LookUpSHCNE = "SHCNE_MEDIAREMOVED"
Case &H80: dbg_LookUpSHCNE = "SHCNE_DRIVEREMOVED"
Case &H100: dbg_LookUpSHCNE = "SHCNE_DRIVEADD"
Case &H200: dbg_LookUpSHCNE = "SHCNE_NETSHARE"
Case &H400: dbg_LookUpSHCNE = "SHCNE_NETUNSHARE"
Case &H800: dbg_LookUpSHCNE = "SHCNE_ATTRIBUTES"
Case &H1000: dbg_LookUpSHCNE = "SHCNE_UPDATEDIR"
Case &H2000: dbg_LookUpSHCNE = "SHCNE_UPDATEITEM"
Case &H4000: dbg_LookUpSHCNE = "SHCNE_SERVERDISCONNECT"
Case &H8000&: dbg_LookUpSHCNE = "SHCNE_UPDATEIMAGE"
Case &H10000: dbg_LookUpSHCNE = "SHCNE_DRIVEADDGUI"
Case &H20000: dbg_LookUpSHCNE = "SHCNE_RENAMEFOLDER"
Case &H40000: dbg_LookUpSHCNE = "SHCNE_FREESPACE"
Case &H4000000: dbg_LookUpSHCNE = "SHCNE_EXTENDED_EVENT"
Case &H8000000: dbg_LookUpSHCNE = "SHCNE_ASSOCCHANGED"
Case &H2381F: dbg_LookUpSHCNE = "SHCNE_DISKEVENTS"
Case &HC0581E0: dbg_LookUpSHCNE = "SHCNE_GLOBALEVENTS"
Case &H7FFFFFFF: dbg_LookUpSHCNE = "SHCNE_ALLEVENTS"
Case &H80000000: dbg_LookUpSHCNE = "SHCNE_INTERRUPT"

End Select
End Function

Private Function LPWSTRtoStr(lPtr As Long, Optional ByVal fFree As Boolean = True) As String
SysReAllocString VarPtr(LPWSTRtoStr), lPtr
If fFree Then
    Call CoTaskMemFree(lPtr)
End If
End Function

Private Function GetPIDLFromFolderID(hOwner As Long, _
                                                             nFolder As oleexp.CSIDLs) As Long
  Dim pidl As Long
  Call SHGetFolderLocation(hOwner, nFolder, 0, 0, pidl)
    GetPIDLFromFolderID = pidl
  
End Function

Private Function StrGetPath(ByVal sFull As String) As String
'Gets a file path
Dim sOut As String
sOut = sFull
If InStrRev(sOut, "\") = 1 Then
    sFull = Left$(sOut, Len(sOut) - 1)
End If
If InStr(sFull, "\") = 0 Then
    StrGetPath = sFull 'usually happens for C:\ top level
    Exit Function
End If
StrGetPath = Left$(sOut, InStrRev(sOut, "\") - 1)
If Right$(StrGetPath, 1) = ":" Then
    StrGetPath = StrGetPath & "\" 'write C:\ instead of C:
End If
End Function

Private Sub DebugAppend(ByVal sMsg As String, Optional ilvl As Long)
If ilvl < dbg_MinLevel Then Exit Sub
If dbg_IncludeDate Then sMsg = "[ST][" & Format$(Now, "yyyy-mm-dd Hh:Mm:Ss") & "] " & sMsg 'yyyy-mm-dd
If dbg_PrintToImmediate Then Debug.Print sMsg
If dbg_RaiseEvent Then RaiseEvent DebugMessage(sMsg, CInt(ilvl))
End Sub

Private Sub dbg_stringbytes(s As String, Optional bOut As Boolean = False)
Dim i As Long
Dim z As String
For i = 1 To Len(s)
    z = z & Format$(Hex$(AscW(Mid(s, i, 1))), "00") & " "
Next i
'If bOut Then
    DebugAppend "StringBytes(" & s & ")", 9
    DebugAppend z, 9
'Else
'    Debug.Print "StringBytes(" & s & ")"
'    Debug.Print z
'End If
End Sub

Private Sub dbg_printstrarr(sAr() As String, bJn As Boolean)
If bJn Then
    DebugAppend Join(sAr, ",")
Else
    Dim i As Long
    For i = LBound(sAr) To UBound(sAr)
        DebugAppend "Array(" & CStr(i) & ")=" & sAr(i)
    Next i
End If
End Sub

Private Function dbg_sfgao_tostring(atr As SFGAO_Flags) As String
Dim sOut As String
If (atr And SFGAO_BROWSABLE) = SFGAO_BROWSABLE Then sOut = sOut & "SFGAO_BROWSEABLE,"
If (atr And SFGAO_CANCOPY) = SFGAO_CANCOPY Then sOut = sOut & "SFGAO_CANCOPY,"
If (atr And SFGAO_CANDELETE) = SFGAO_CANDELETE Then sOut = sOut & "SFGAO_CANDELETE,"
If (atr And SFGAO_CANLINK) = SFGAO_CANLINK Then sOut = sOut & "SFGAO_CANLINK,"
If (atr And SFGAO_CANMONIKER) = SFGAO_CANMONIKER Then sOut = sOut & "SFGAO_CANMONIKER,"
If (atr And SFGAO_CANMOVE) = SFGAO_CANMOVE Then sOut = sOut & "SFGAO_CANMOVE,"
If (atr And SFGAO_CANRENAME) = SFGAO_CANRENAME Then sOut = sOut & "SFGAO_CANRENAME,"
If (atr And SFGAO_COMPRESSED) = SFGAO_COMPRESSED Then sOut = sOut & "SFGAO_COMPRESSED,"
If (atr And SFGAO_DROPTARGET) = SFGAO_DROPTARGET Then sOut = sOut & "SFGAO_DROPTARGET,"
If (atr And SFGAO_ENCRYPTED) = SFGAO_ENCRYPTED Then sOut = sOut & "SFGAO_ENCRYPTED,"
If (atr And SFGAO_FILESYSANCESTOR) = SFGAO_FILESYSANCESTOR Then sOut = sOut & "SFGAO_FILESYSANCESTOR,"
If (atr And SFGAO_FILESYSTEM) = SFGAO_FILESYSTEM Then sOut = sOut & "SFGAO_FILESYSTEM,"
If (atr And SFGAO_FOLDER) = SFGAO_FOLDER Then sOut = sOut & "SFGAO_FOLDER,"
If (atr And SFGAO_GHOSTED) = SFGAO_GHOSTED Then sOut = sOut & "SFGAO_GHOSTED,"
If (atr And SFGAO_HASPROPSHEET) = SFGAO_HASPROPSHEET Then sOut = sOut & "SFGAO_HASPROPSHEET,"
If (atr And SFGAO_HASSTORAGE) = SFGAO_HASSTORAGE Then sOut = sOut & "SFGAO_HASSTORAGE,"
If (atr And SFGAO_HASSUBFOLDER) = SFGAO_HASSUBFOLDER Then sOut = sOut & "SFGAO_HASSUBFOLDER,"
If (atr And SFGAO_HIDDEN) = SFGAO_HIDDEN Then sOut = sOut & "SFGAO_HIDDEN,"
If (atr And SFGAO_ISSLOW) = SFGAO_ISSLOW Then sOut = sOut & "SFGAO_ISSLOW,"
If (atr And SFGAO_LINK) = SFGAO_LINK Then sOut = sOut & "SFGAO_LINK,"
If (atr And SFGAO_NEWCONTENT) = SFGAO_NEWCONTENT Then sOut = sOut & "SFGAO_NEWCONTENT,"
If (atr And SFGAO_NONENUMERATED) = SFGAO_NONENUMERATED Then sOut = sOut & "SFGAO_NONENUMERATED,"
If (atr And SFGAO_READONLY) = SFGAO_READONLY Then sOut = sOut & "SFGAO_READONLY,"
If (atr And SFGAO_REMOVABLE) = SFGAO_REMOVABLE Then sOut = sOut & "SFGAO_REMOVABLE,"
If (atr And SFGAO_SHARE) = SFGAO_SHARE Then sOut = sOut & "SFGAO_SHARE,"
If (atr And SFGAO_STORAGE) = SFGAO_STORAGE Then sOut = sOut & "SFGAO_STORAGE,"
If (atr And SFGAO_STORAGEANCESTOR) = SFGAO_STORAGEANCESTOR Then sOut = sOut & "SFGAO_STORAGEANCESTOR,"
If (atr And SFGAO_STREAM) = SFGAO_STREAM Then sOut = sOut & "SFGAO_STREAM,"
If (atr And SFGAO_SYSTEM) = SFGAO_SYSTEM Then sOut = sOut & "SFGAO_SYSTEM,"
If (atr And SFGAO_VALIDATE) = SFGAO_VALIDATE Then sOut = sOut & "SFGAO_VALIDATE,"
dbg_sfgao_tostring = sOut
End Function

Private Function PathGetDisp(ByVal sPath As String) As String
If Right$(sPath, 1) = "\" Then
    sPath = Left$(sPath, Len(sPath) - 1)
End If
Dim psi As oleexp.IShellItem
Dim lp As Long
If (sPath = "0") And (mComputerAsRoot = False) Then
    sPath = sDesktopPath
End If
If (sPath = "1") Then
    sPath = sFavPath
End If
Call oleexp.SHCreateItemFromParsingName(StrPtr(sPath), Nothing, IID_IShellItem, psi)
If (psi Is Nothing) = False Then
    psi.GetDisplayName SIGDN_NORMALDISPLAY, lp
    PathGetDisp = LPWSTRtoStr(lp)
End If
End Function

'Public Sub TestLink()
'DebugAppend "testlink " & TVEntries(gCurSelIdx).sFullPath
'Dim sz As String
'Dim si As oleexp.IShellItem
'oleexp.SHCreateItemFromParsingName StrPtr(TVEntries(gCurSelIdx).sFullPath), Nothing, IID_IShellItem, si
'
'sz = GetLinkTarget(si)
'MessageBoxW UserControl.hWnd, StrPtr(sz), 0&, 1&
'
'End Sub
Private Function GetLinkTarget(siLink As oleexp.IShellItem) As String
    '<EhHeader>
    On Error GoTo e0
    '</EhHeader>
Dim isl As oleexp.IShellLinkW
Dim sTmp As String
Dim wfd As oleexp.WIN32_FIND_DATAW
siLink.BindToHandler 0&, BHID_SFUIObject, IID_IShellLinkW, isl
If (isl Is Nothing) = False Then
    sTmp = String$(MAX_PATH, 0)
    isl.GetPath sTmp, MAX_PATH, wfd, 0&
    sTmp = TrimNullW(sTmp)
'    If sTmp = "" Then
'        Dim pidl As Long
'        pidl = isl.GetIDList
'
    GetLinkTarget = sTmp
End If
 
'<EhFooter>
Exit Function

e0:
    DebugAppend "ucShellTree.GetLinkTarget->Error: " & Err.Description & ", 0x" & Hex$(Err.Number)
'</EhFooter>
End Function

Private Function GetLinkTargetPIDL(siLink As oleexp.IShellItem) As Long
    '<EhHeader>
    On Error GoTo e0
    '</EhHeader>
Dim isl As oleexp.IShellLinkW
Dim sTmp As String
Dim wfd As oleexp.WIN32_FIND_DATAW
siLink.BindToHandler 0&, BHID_SFUIObject, IID_IShellLinkW, isl
If (isl Is Nothing) = False Then
    GetLinkTargetPIDL = isl.GetIDList()
End If
 
'<EhFooter>
Exit Function

e0:
    DebugAppend "ucShellTree.GetLinkTargetPIDL->Error: " & Err.Description & ", 0x" & Hex$(Err.Number)
'</EhFooter>
End Function

Private Function LibGetDefLoc2(szPath As String) As String
'if current path is a library, get its default save location to drop there
Dim siLib As oleexp.IShellItem
Dim siItem As oleexp.IShellItem
Dim lpPath As Long
Dim pSL As oleexp.ShellLibrary
Set pSL = New oleexp.ShellLibrary

oleexp.SHCreateItemFromParsingName StrPtr(szPath), Nothing, IID_IShellItem, siLib

If (siLib Is Nothing) = False Then
    pSL.LoadLibraryFromItem siLib, STGM_READ
    pSL.GetDefaultSaveFolder DSFT_DETECT, IID_IShellItem, siItem
    If (siItem Is Nothing) = False Then
        siItem.GetDisplayName SIGDN_FILESYSPATH, lpPath
        LibGetDefLoc2 = LPWSTRtoStr(lpPath)
        siItem.GetDisplayName SIGDN_NORMALDISPLAY, lpPath
        mDragOver = LPWSTRtoStr(lpPath)
        Set siItem = Nothing
    End If
    
    Set pSL = Nothing
    Set siLib = Nothing
End If

End Function

Private Function InvokeVerb(pcm As oleexp.IContextMenu, ByVal pszVerb As String, Optional lFlags As InvokeCommandMask = 0&) As Long
    '<EhHeader>
    On Error GoTo InvokeVerb_Err
    '</EhHeader>

Dim hMenu As Long
Dim tCmdInfo As oleexp.CMINVOKECOMMANDINFO
Dim tCmdInfoEx As oleexp.CMINVOKECOMMANDINFOEX
hMenu = CreatePopupMenu()
Dim sVerb As String
Dim pcm2 As oleexp.IContextMenu2
sVerb = pszVerb
Dim lpCmd As Long
Dim PT As oleexp.POINT
Dim nItems As Long
Dim sDesc As String
Dim i As Long
Dim bGo As Boolean
lpCmd = -1
If hMenu Then
'    Set pcm2 = pcm
    If (pcm Is Nothing) = False Then
'        DebugAppend "Got pcm"
        pcm.QueryContextMenu hMenu, 0&, 1&, &H7FFF, CMF_EXPLORE Or CMF_OPTIMIZEFORINVOKE
        nItems = GetMenuItemCount(hMenu)
        DebugAppend "InvokeVerb nItems=" & nItems
        For i = 0 To nItems - 1
            lpCmd = GetMenuItemID(hMenu, i)
'            DebugAppend "i=" & i & ", lpCmd=" & lpCmd
            If lpCmd <= 0 Then GoTo nxt
            On Error Resume Next
            sVerb = String$(MAX_PATH, 0&)
            pcm.GetCommandString lpCmd - 1, GCS_VERBW, 0&, StrPtr(sVerb), Len(sVerb)
            On Error GoTo InvokeVerb_Err
            If (Err.Number = 0&) Then
                sVerb = TrimNullW(sVerb)
                DebugAppend "CmdStr=" & sVerb
'                DebugAppend "InvokeVerb hMenu=" & hMenu
                On Error Resume Next
                sDesc = String$(MAX_PATH, 0&)
                pcm.GetCommandString lpCmd - 1, GCS_HELPTEXTW, 0&, StrPtr(sDesc), Len(sDesc)
                sDesc = TrimNullW(sDesc)
                On Error GoTo InvokeVerb_Err
'                DebugAppend "HelpStr=" & sDesc
                If LCase$(sVerb) = LCase$(pszVerb) Then
                    DebugAppend "InvokeVerb go on " & sVerb
                    bGo = True
                    Exit For
                End If
            Else
                DebugAppend "InvokeVerb.GetCommandString Error::" & Err.Description
                Err.Clear
            End If
nxt:
        Next i
        If bGo Then
            tCmdInfo.cbSize = Len(tCmdInfo)
            tCmdInfo.hWnd = hTVD
            tCmdInfo.lpVerb = lpCmd - 1&
            tCmdInfo.fMask = lFlags
            tCmdInfo.nShow = SW_SHOWNORMAL
            pcm.InvokeCommand VarPtr(tCmdInfo)
            DebugAppend "command invoked"
        End If
        Call DestroyMenu(hMenu)
    End If
End If

 '<EhFooter>
 Exit Function

InvokeVerb_Err:
    DebugAppend "InvokeVerb->" & Err.Description & " (" & Err.Number & ")"
 '</EhFooter>
End Function

Private Sub UpdateStatus(sText As String)
RaiseEvent StatusMessage(sText)
End Sub

Private Sub TVDragHover(X As Long, Y As Long, lKeyState As Long)
Dim tvhti As TVHITTESTINFO
Dim ppt As oleexp.POINT
ppt.X = X
ppt.Y = Y
Call ScreenToClient(hTVD, ppt)
tvhti.PT.X = ppt.X
tvhti.PT.Y = ppt.Y
TreeView_HitTest hTVD, tvhti
If (tvhti.Flags And TVHT_ONITEM) Then
    DebugAppend "DragHover " & X & "," & Y & ",hItem=" & GetTVItemText(hTVD, tvhti.hItem)
    TreeView_Expand hTVD, tvhti.hItem, TVE_EXPAND
End If

End Sub

Private Sub TVQueryDragOverData(in_ItemIndex As Long, in_fIsGroup As Boolean, out_FullPath As String, out_Invalid As Boolean)
out_FullPath = GetTVItemPath(hTVD, in_ItemIndex)
If out_FullPath = "0" Then
    If (mComputerAsRoot = False) Then
        out_FullPath = sDesktopPath
    End If
End If
If out_FullPath = "1" Then
    out_FullPath = sFavPath
End If
If out_FullPath = "" Then
    out_Invalid = True
Else
    out_Invalid = False
End If

End Sub

Private Sub TVQueryDragOverItem(in_ptX As Long, in_ptY As Long, in_PrevIndex As Long, out_NewIndex As Long, out_fGroup As Boolean, out_fValid As Long, out_fFolder As Long)
Dim PT As oleexp.POINT
PT.X = in_ptX
PT.Y = in_ptY
out_NewIndex = TVDragOver(hTVD, PT, in_PrevIndex, out_fValid, out_fFolder)
If out_NewIndex <> in_PrevIndex Then
    DebugAppend "TVQueryDragOverItem found new item, " & in_PrevIndex & " != " & out_NewIndex & "; fldr=" & out_fFolder
End If
End Sub

Private Sub SetPreferredEffect(ido As oleexp.IDataObject, nVal As Long)
Dim fmt As oleexp.FORMATETC
Dim sTg As oleexp.STGMEDIUM
Dim tDD As oleexp.DROPDESCRIPTION
Dim hGlobal As Long, lpGlobal As Long
Dim lpFmt As Long
Dim i As Long
'Dim cfstr As String
'cfstr = "Preferred DropEffect"
DebugAppend "adddil.enter"

hGlobal = GlobalAlloc(GHND, LenB(nVal))
If hGlobal Then
    lpGlobal = GlobalLock(hGlobal)
    Call CopyMemory(ByVal lpGlobal, nVal, LenB(nVal))
    DebugAppend "adddil.copymem"
    Call GlobalUnlock(hGlobal)
    sTg.TYMED = TYMED_HGLOBAL
    sTg.data = lpGlobal
    fmt.cfFormat = CF_PREFERREDDROPEFFECT
    fmt.dwAspect = DVASPECT_CONTENT
    fmt.lIndex = -1
    fmt.TYMED = TYMED_HGLOBAL
    ido.SetData fmt, sTg, 1
Else
    DebugAppend "failed to get hglobal"
End If
End Sub

Private Function AddDropDescription(ido As oleexp.IDataObject, nType As oleexp.DROPIMAGETYPE, sMsg As String, sIns As String) As Boolean
'If bFlagBlockDD Then Exit Function
If (ido Is Nothing) Then
    DebugAppend "AddDropDescription::No data object, aborting"
    Exit Function
End If
Dim fmt As oleexp.FORMATETC
Dim sTg As oleexp.STGMEDIUM
Dim tDD As oleexp.DROPDESCRIPTION
Dim iTmp1() As Integer
Dim iTmp2() As Integer
Dim hGlobal As Long, lpGlobal As Long
Dim i As Long
On Error GoTo e0
AddDropDescription = True
DebugAppend "cDropTarget.AddDropDescription.Entry" ', 9
Str2WCHAR sMsg, iTmp1
Str2WCHAR sIns, iTmp2

For i = 0 To UBound(iTmp1)
    tDD.szMessage(i) = iTmp1(i)
Next i

For i = 0 To UBound(iTmp2)
    tDD.szInsert(i) = iTmp2(i)
Next i
tDD.Type = nType

hGlobal = GlobalAlloc(GHND, LenB(tDD))
If hGlobal Then
    lpGlobal = GlobalLock(hGlobal)
    Call CopyMemory(ByVal lpGlobal, tDD, LenB(tDD))
    Call GlobalUnlock(hGlobal)
    
    sTg.TYMED = TYMED_HGLOBAL
    sTg.data = lpGlobal
    
    fmt.cfFormat = CF_DROPDESCRIPTION
    fmt.dwAspect = DVASPECT_CONTENT
    fmt.lIndex = -1
    fmt.TYMED = TYMED_HGLOBAL
        
    ido.SetData fmt, sTg, 1
End If
Exit Function
e0:
    DebugAppend "AddDropDescription->" & Err.Description & " (" & Err.Number & ")"
    AddDropDescription = False
End Function

Private Sub Str2WCHAR(sz As String, iOut() As Integer)
Dim i As Long
ReDim iOut(255)
'If Len(sz) > MAX_PATH Then sz = Left$(sz, MAX_PATH)
For i = 1 To Len(sz)
    iOut(i - 1) = AscW(Mid$(sz, i, 1))
Next i

End Sub

Private Sub IDropTarget_DragEnter(ByVal pDataObj As oleexp.IDataObject, ByVal grfKeyState As Long, ByVal ptX As Long, ByVal ptY As Long, pdwEffect As oleexp.DROPEFFECTS)
    '<EhHeader>
    On Error GoTo e0
    '</EhHeader>
   Dim PT As oleexp.POINT
   PT.X = ptX
   PT.Y = ptY
   bNoDrop = False
   m_sAltDrop = ""
   If grfKeyState And MK_RBUTTON Then
    ddRightButton = True
Else
    ddRightButton = False
End If

'NOTE: At first, indicate no drop until a folder is dragged over.
If mFullRowSelect = False Then
    bNoDrop = True
    pdwEffect = DROPEFFECT_NONE
'    SetPreferredEffect pDataObj, DROPEFFECT_NONE
    AddDropDescription pDataObj, DROPIMAGE_NONE, "Can't drop here.", ""
End If

skp:
   pDTH.DragEnter m_hWnd, pDataObj, PT, pdwEffect
   DebugAppend "DragEnter::Effect=" & pdwEffect
  Set mDataObj = pDataObj

'    RaiseEvent DragEnter(pDataObj, grfKeyState, ptx, pty, pdwEffect)

'<EhFooter>
Exit Sub

e0:
    DebugAppend "ucShellTree.IDropTarget_DragEnter->Error: " & Err.Description & ", 0x" & Hex$(Err.Number)
'</EhFooter>
End Sub

Private Sub IDropTarget_DragLeave()
pDTH.DragLeave
TreeView_SelectDropTarget hTVD, 0&

End Sub

Private Sub IDropTarget_DragOver(ByVal grfKeyState As Long, ByVal ptX As Long, ByVal ptY As Long, pdwEffect As oleexp.DROPEFFECTS)
    '<EhHeader>
    On Error GoTo e0
    '</EhHeader>
   Dim PT As oleexp.POINT
   PT.X = ptX
   PT.Y = ptY

'mDefEffect = GetPreferredEffect(mDataObj)

    pDTH.DragOver PT, pdwEffect
 If (xHover <> ptX) Or (yHover <> ptY) Then
    lHover1 = GetTickCount()
    xHover = ptX
    yHover = ptY
    bHoverFired = False
Else
    If lHover1 = 0 Then 'initial run
        lHover1 = GetTickCount()
        xHover = ptX
        yHover = ptY
    Else
        If bHoverFired = False Then
            lHover2 = GetTickCount()
            If (lHover2 - lHover1) > lRaiseHover Then
                DebugAppend "DragHover " & ptX & "," & ptY & ",t1=" & lHover1 & ",t2=" & lHover2
'                RaiseEvent DragHover(ptx, pty, grfKeyState)
                TVDragHover ptX, ptY, grfKeyState
                bHoverFired = True 'set flag to not raise again until pt changes
'                lHover1 = 0
            End If
        End If
    End If
End If
Dim lPrevItemIndex As Long
Dim fValid As Long, fFolder As Long
'pdwEffect = mDefEffect
If DataObjSupportsFormat(mDataObj, CF_PREFERREDDROPEFFECT) Then

        pdwEffect = GetPreferredEffect(mDataObj)

Else
    pdwEffect = DROPEFFECT_MOVE
End If

If bAbort Then Exit Sub
If True Then 'was m_ActiveDrop but active drop is mandatory for this control
    lPrevItemIndex = lItemIndex
    Dim fGrp As Boolean
    Dim bNo As Boolean
    Dim sOld As String
    Dim nPrev As Long
    nPrev = lItemIndex
    
    TVQueryDragOverItem ptX, ptY, nPrev, lItemIndex, fGrp, fValid, fFolder
    If lItemIndex <> nPrev Then
    If (lItemIndex >= 0) Then
        If fGrp Then
            sOld = sFolder
            TVQueryDragOverData lItemIndex, True, sFolder, bNo
            If sOld <> sFolder Then
                mDragOver = PathGetDisp(sFolder)
            End If
        Else
            sOld = sFolder
            TVQueryDragOverData lItemIndex, False, sFolder, bNo
            If sOld <> sFolder Then
                mDragOver = PathGetDisp(sFolder)
            End If

        End If
        If (Left$(sFolder, 3) = "::{") Then

            If (Len(sFolder) > lnLibRoot) Then
                If Left$(LCase$(sFolder), Len(sLibRoot)) = LCase$(sLibRoot) Then
                    sFolder = LibGetDefLoc2(sFolder)
                    If sFolder <> "" Then
                        'override: the item itself doesn't have the SFGAO_DROPTARGET flag
                        '          but the default location does
                        fValid = 1
                        fFolder = 1
                        TreeView_SelectDropTarget hTVD, lItemIndex
                    End If
                End If
            Else
                bNo = True
            End If
        End If
        DebugAppend "dragover set dir=" & sFolder & ",idx=" & lItemIndex & ",valid=" & fValid & ",folder=" & fFolder  ', 9
    Else
 
            DebugAppend "QueryDragOverItem index=" & lItemIndex
            mDragOver = ""
            sFolder = ""
 
    End If
    End If
    If lItemIndex = -1 Then bNo = True
    If bNo Then
        pdwEffect = DROPEFFECT_NONE
        DebugAppend "DragOver::bNo=True,set->none"
    End If
    If (lItemIndex <> lPrevItemIndex) Then
    DebugAppend "IDropTarget.DragOver::New dragover item; resetting DropDesc, lPrevItem=" & lPrevItemIndex & ",cur=" & lItemIndex
'        DebugAppend "cDropTarget.DragOver::New dragover item; resetting DropDesc", 3
        If False Then 'If bQueryTip Then

        Else
DefaultTipSet:

                bNoTarget = False
                If fValid = 0 Then
                    DebugAppend "DragOver::lItemIndex=-1,set effect (NOCHANGE),cur=" & pdwEffect
                    sFolder = ""
                    mDragOver = ""
                    If bNoDrop Then
                        pdwEffect = DROPEFFECT_NONE
                    End If
'                    bNoTarget = True
                    fFolder = 1
                End If
                Dim vdt As Long
                If lItemIndex = m_hFav Then
                    fFolder = 0
                End If
                If fFolder = 0 Then
                    vdt = VerifyDropTarget(lItemIndex)
                    If vdt = 0 Then
                        DebugAppend "Drop target (nonfolder) failed verify"
                        fFolder = 1
                        pdwEffect = DROPEFFECT_NONE
                    End If
                    If vdt = 2 Then
                        fFolder = 1
                    End If
                    If vdt = 3 Then
                        pdwEffect = DROPEFFECT_LINK
                        DebugAppend "SetEffectLink"
                        fFolder = 1
                    End If
                End If
                    
                If fFolder Then
                Select Case pdwEffect
                    Case DROPEFFECT_NONE
                        AddDropDescription mDataObj, DROPIMAGE_NONE, "Can't drop here.", ""
                    Case DROPEFFECT_COPY
                        AddDropDescription mDataObj, DROPIMAGE_COPY, IIf(mDragOver = "", "Copy here", "Copy to %1"), mDragOver
                    Case DROPEFFECT_MOVE
                        AddDropDescription mDataObj, DROPIMAGE_MOVE, IIf(mDragOver = "", "Move here", "Move to %1"), mDragOver
                    Case DROPEFFECT_LINK
                        AddDropDescription mDataObj, DROPIMAGE_MOVE, IIf(mDragOver = "", "Create shortcut here", "Create shortcut in %1"), mDragOver
                End Select
                Else
                    AddDropDescription mDataObj, DROPIMAGE_COPY, IIf(mDragOver = "", "Open", "Open with %1"), mDragOver
                End If
            End If
'        End If
    End If
End If
qts:
   
'    RaiseEvent DragOver(grfKeyState, ptx, pty, pdwEffect)

'<EhFooter>
Exit Sub

e0:
    DebugAppend "ucShellTree.IDropTarget_DragOver->Error: " & Err.Description & ", 0x" & Hex$(Err.Number)
'</EhFooter>
End Sub

Private Sub IDropTarget_Drop(ByVal pDataObj As oleexp.IDataObject, ByVal grfKeyState As Long, ByVal ptX As Long, ByVal ptY As Long, pdwEffect As oleexp.DROPEFFECTS)
  DebugAppend "IDT_Drop "
   Dim PT As oleexp.POINT
   PT.X = ptX
   PT.Y = ptY

pDTH.Drop pDataObj, PT, pdwEffect
TVHandleDrop pDataObj, pdwEffect, PT, grfKeyState
sFolder = ""
TreeView_SelectDropTarget hTVD, 0&
End Sub

Private Sub TVHandleDrop(pdo As oleexp.IDataObject, pdwEffect As oleexp.DROPEFFECTS, PT As oleexp.POINT, dwKeys As Long)
    '<EhHeader>
    On Error GoTo e0
    '</EhHeader>
Dim hSel As Long
Dim lp As Long
Dim psia As oleexp.IShellItemArray
Dim siTarget As oleexp.IShellItem
Dim pDT As oleexp.IDropTarget
Dim lpTar As Long, sTar As String
Dim sFilesOut() As String
Dim nFilesOut As Long
ReDim sFilesOut(0)
Dim lButton As Long
If ddRightButton Then
    lButton = MK_RBUTTON
Else
    lButton = MK_LBUTTON
End If

'hSel = TreeView_GetSelection(hTVD)
'lp = GetTVItemlParam(hTVD, hSel)
DebugAppend "DropTarget=" & sFolder
If sFolder = "" Then Exit Sub
oleexp.SHCreateItemFromParsingName StrPtr(sFolder), Nothing, IID_IShellItem, siTarget
If (siTarget Is Nothing) = False Then
    siTarget.BindToHandler 0&, BHID_SFUIObject, IID_IDropTarget, pDT
    If (pDT Is Nothing) = False Then
        pDT.DragEnter pdo, lButton, PT.X, PT.Y, pdwEffect
        pDT.Drop pdo, lButton, PT.X, PT.Y, DROPEFFECT_MOVE Or DROPEFFECT_COPY Or DROPEFFECT_LINK
    Else
        DebugAppend "TVHandleDrop->Failed to get drop target"
    End If
    siTarget.GetDisplayName SIGDN_DESKTOPABSOLUTEPARSING, lpTar
    sTar = LPWSTRtoStr(lpTar)
Else
    DebugAppend "TVHandleDrop->Failed to get target shell item"
End If


oleexp.SHCreateShellItemArrayFromDataObject pdo, IID_IShellItemArray, psia
If (psia Is Nothing) = False Then
    Dim pEnum As oleexp.IEnumShellItems
    Dim sia As oleexp.IShellItem
    Dim lpc As Long, szc As String
    Dim pcl As Long
    psia.EnumItems pEnum
    If (pEnum Is Nothing) = False Then
        Do While pEnum.Next(1&, sia, pcl) = S_OK
            sia.GetDisplayName SIGDN_DESKTOPABSOLUTEPARSING, lpc
            ReDim Preserve sFilesOut(nFilesOut)
            sFilesOut(nFilesOut) = LPWSTRtoStr(lpc)
            nFilesOut = nFilesOut + 1
        Loop
    End If
End If

RaiseEvent DropFiles(sFilesOut, psia, pdo, sTar, siTarget, pdwEffect, dwKeys, PT.X, PT.Y)
'<EhFooter>
Exit Sub

e0:
    DebugAppend "ucShellTree.TVHandleDrop->Error: " & Err.Description & ", 0x" & Hex$(Err.Number)
'</EhFooter>
End Sub

Private Sub SetCheckedList()
ReDim gPaths(0)
nPaths = 0
'EnumPaths hTVD, m_hRoot
Dim i As Long
For i = 0 To UBound(TVEntries)
    If (TVEntries(i).Checked = True) And (TVEntries(i).bDeleted = False) Then
        ReDim Preserve gPaths(nPaths)
        gPaths(nPaths) = TVEntries(i).sFullPath
        nPaths = nPaths + 1
    End If
Next i
End Sub

Private Sub SetExCheckedList()
ReDim gExPaths(0)
nExPaths = 0
'EnumPaths hTVD, m_hRoot
Dim i As Long
For i = 0 To UBound(TVEntries)
    If (TVEntries(i).Excluded = True) And (TVEntries(i).bDeleted = False) Then
        ReDim Preserve gExPaths(nExPaths)
        gExPaths(nExPaths) = TVEntries(i).sFullPath
        nExPaths = nExPaths + 1
    End If
Next i
End Sub

Private Sub SetSelection(n As Long, hi As Long)
Dim lp As Long
Dim ppidlp As Long
If n = -1 Then
    Set siSelected = Nothing
    sSelectedItem = ""
    Exit Sub
End If
If TVEntries(n).sFullPath = "0" Then
    If mComputerAsRoot Then
        SHGetKnownFolderItem FOLDERID_ComputerFolder, KF_FLAG_DEFAULT, 0&, IID_IShellItem, siSelected
    Else
        SHGetKnownFolderItem FOLDERID_Desktop, KF_FLAG_DEFAULT, 0&, IID_IShellItem, siSelected
    End If
ElseIf TVEntries(n).sFullPath = "1" Then
        SHGetKnownFolderItem FOLDERID_Links, KF_FLAG_DEFAULT, 0&, IID_IShellItem, siSelected
Else
    If mFavorites Then
        DebugAppend "node=" & TVEntries(n).hParentNode & ",fav=" & m_hFav & ",txt=" & TVEntries(1).sName & ",tgt=" & TVEntries(n).sLinkTarget
        If TVEntries(n).hParentNode = m_hFav Then
            If TVEntries(n).LinkPIDL Then
                DebugAppend "Using link pidl"
                oleexp.SHCreateItemFromIDList TVEntries(n).LinkPIDL, IID_IShellItem, siSelected
            Else
                oleexp.SHCreateItemFromParsingName StrPtr(TVEntries(n).sLinkTarget), Nothing, IID_IShellItem, siSelected
                If (siSelected Is Nothing) Then
                    DebugAppend "SetSelection::ParsingPath->ShellItem failed, trying to load from pidls...", 3
                    If (TVEntries(n).pidlFQPar <> 0&) And (TVEntries(n).pidlRel <> 0&) Then
                        ppidlp = ILCombine(TVEntries(n).pidlFQPar, TVEntries(n).pidlRel)
                        oleexp.SHCreateItemFromIDList ppidlp, IID_IShellItem, siSelected
                        If (siSelected Is Nothing) = False Then
                            DebugAppend "SetSelection::LoadFromPidl->Success!", 3
                        Else
                            DebugAppend "SetSelection::LoadFromPidl->Failed.", 3
                        End If
                    Else
                        DebugAppend "SetSelection::LoadFromPidl->Parent or child pidl not set.", 3
                    End If
                End If
            End If
        Else
            If TVEntries(n).pidlFQ Then
                oleexp.SHCreateItemFromIDList TVEntries(n).pidlFQ, IID_IShellItem, siSelected
            Else
                oleexp.SHCreateItemFromParsingName StrPtr(TVEntries(n).sFullPath), Nothing, IID_IShellItem, siSelected
            End If
            If (siSelected Is Nothing) Then
                DebugAppend "SetSelection::ParsingPath->ShellItem failed, trying to load from pidls...", 3
                If (TVEntries(n).pidlFQPar <> 0&) And (TVEntries(n).pidlRel <> 0&) Then
                    ppidlp = ILCombine(TVEntries(n).pidlFQPar, TVEntries(n).pidlRel)
                    oleexp.SHCreateItemFromIDList ppidlp, IID_IShellItem, siSelected
                    If (siSelected Is Nothing) = False Then
                        DebugAppend "SetSelection::LoadFromPidl->Success!", 3
                    Else
                        DebugAppend "SetSelection::LoadFromPidl->Failed.", 3
                    End If
                Else
                    DebugAppend "SetSelection::LoadFromPidl->Parent or child pidl not set.", 3
                End If
            End If
        End If
    Else
        If TVEntries(n).pidlFQ Then
            oleexp.SHCreateItemFromIDList TVEntries(n).pidlFQ, IID_IShellItem, siSelected
        Else
            oleexp.SHCreateItemFromParsingName StrPtr(TVEntries(n).sFullPath), Nothing, IID_IShellItem, siSelected
        End If
        If (siSelected Is Nothing) Then
            DebugAppend "SetSelection::ParsingPath->ShellItem failed, trying to load from pidls...", 3
            If (TVEntries(n).pidlFQPar <> 0&) And (TVEntries(n).pidlRel <> 0&) Then
                ppidlp = ILCombine(TVEntries(n).pidlFQPar, TVEntries(n).pidlRel)
                oleexp.SHCreateItemFromIDList ppidlp, IID_IShellItem, siSelected
                If (siSelected Is Nothing) = False Then
                    DebugAppend "SetSelection::LoadFromPidl->Success!", 3
                Else
                    DebugAppend "SetSelection::LoadFromPidl->Failed.", 3
                End If
            Else
                DebugAppend "SetSelection::LoadFromPidl->Parent or child pidl not set.", 3
            End If
        End If
    End If
End If
        
siSelected.GetDisplayName SIGDN_DESKTOPABSOLUTEPARSING, lp
sSelectedItem = LPWSTRtoStr(lp)
Dim lpName As Long, sName As String
Dim bFl As Boolean
siSelected.GetDisplayName SIGDN_NORMALDISPLAY, lpName
sName = LPWSTRtoStr(lpName)
DebugAppend "SetSel " & sName & ",path=" & sSelectedItem
Dim lAtr As SFGAO_Flags
siSelected.GetAttributes SFGAO_FOLDER, lAtr
If (lAtr And SFGAO_FOLDER) = SFGAO_FOLDER Then
    bFl = True
End If

If (fLoad = 0&) Then RaiseEvent ItemSelect(sName, sSelectedItem, bFl, hi)
If (fLoad = 0&) Then RaiseEvent ItemSelectByShellItem(siSelected, sName, sSelectedItem, bFl, hi)

If ppidlp Then CoTaskMemFree ppidlp
End Sub

Private Function NormalizePath(sz As String) As String
'Our add routine doesn't include the trailing slash except for roots
'so normalize means removing it
If Len(sz) > 4 Then
    If Right$(sz, 1) = "\" Then
        sz = Left$(sz, Len(sz) - 1)
    End If
End If
NormalizePath = sz
End Function

Public Function PathSetCheck(ByVal sPath As String, bChecked As Boolean, Optional bInsertIfNeeded As Boolean = False) As Long
'sets check state and returns old state
'bInsertIfNeeded will expand all needed folders to show the item
If mCheckboxes = False Then Exit Function
Dim fChk As Long
Dim idx As Long
sPath = NormalizePath(sPath)
idx = GetIndexFromPath(sPath)

If (idx = -1) And (bInsertIfNeeded = True) Then
    DebugAppend "PathSetCheck->Open to path"
    OpenToPath sPath, False, False
    idx = GetIndexFromPath(sPath)
End If
DebugAppend "PathSetCheck->AfterInsertBlock idx=" & idx
If idx < 1 Then Exit Function
If TVEntries(idx).bDeleted Then Exit Function

fChk = TreeView_GetCheckState(hTVD, TVEntries(idx).hNode)
DebugAppend "PathSetCheck->idx=" & idx & ",cur=" & fChk
If bChecked Then
    If fChk = 1 Then '1=unchecked,2=checked,3=partial (0=no box at all)
        TreeView_SetCheckState hTVD, TVEntries(idx).hNode, 1
    End If
Else
    If fChk = 2 Then
        TreeView_SetCheckState hTVD, TVEntries(idx).hNode, 0
    End If
End If
PathSetCheck = fChk
End Function

Public Function PathGetCheck(ByVal sPath As String) As Long
'-1=checkbox mode off, -2=path not found, 0=no box on item, 1=unchecked, 2=checked, 3=partial check
PathGetCheck = -1
If mCheckboxes = False Then Exit Function
Dim idx As Long

PathGetCheck = -2
sPath = NormalizePath(sPath)
idx = GetIndexFromPath(sPath)
If idx < 1 Then Exit Function
If TVEntries(idx).bDeleted Then Exit Function

PathGetCheck = TreeView_GetCheckState(hTVD, TVEntries(idx).hNode)

End Function

Private Sub PropFont_FontChanged(ByVal PropertyName As String)
Dim OldFontHandle As Long
OldFontHandle = hFontTV
Set mIFMain = PropFont
Dim lftmp As LOGFONT
GetObjectW mIFMain.hFont, LenB(lftmp), lftmp
hFontTV = CreateFontIndirect(lftmp)
If hTVD <> 0 Then SendMessageW hTVD, WM_SETFONT, hFontTV, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Sub
'***************************************************
'Begin Control Properties

Public Property Get BackColor() As stdole.OLE_COLOR: BackColor = clrBack: End Property
Attribute BackColor.VB_Description = "Set the background color of the TreeView."
Public Property Let BackColor(ByVal val As stdole.OLE_COLOR)
If val = &HFFFFFFFF Then Exit Property
clrBack = val
Dim clrt As Long
OleTranslateColor clrBack, 0&, clrt
SendMessageW hTVD, TVM_SETBKCOLOR, 0, ByVal clrt
UserControl.Refresh
'RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
UserControl.PropertyChanged "BackColor"

End Property
Public Property Get ForeColor() As stdole.OLE_COLOR: ForeColor = clrFore: End Property
Attribute ForeColor.VB_Description = "Set the color of the text in the TreeView."
Public Property Let ForeColor(ByVal val As stdole.OLE_COLOR)
clrFore = val
Dim clrt As Long
OleTranslateColor clrFore, 0&, clrt
SendMessage hTVD, TVM_SETTEXTCOLOR, 0, ByVal clrt
UserControl.Refresh
RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
UserControl.PropertyChanged "ForeColor"
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Sets the font for the TreeView items. "
Set Font = PropFont
End Property

Public Property Let Font(ByVal NewFont As StdFont)
Set Me.Font = NewFont
End Property

Public Property Set Font(ByVal NewFont As StdFont)
If NewFont Is Nothing Then Set NewFont = Ambient.Font
Dim OldFontHandle As Long
Set PropFont = NewFont
OldFontHandle = hFontTV
Set mIFMain = PropFont
Dim lftmp As LOGFONT
GetObjectW mIFMain.hFont, LenB(lftmp), lftmp
hFontTV = CreateFontIndirect(lftmp)
If hTVD <> 0 Then SendMessageW hTVD, WM_SETFONT, hFontTV, ByVal 1&
If OldFontHandle <> 0 Then DeleteObject OldFontHandle
UserControl.PropertyChanged "Font"
End Property

Public Property Get Enabled() As Boolean: Enabled = UserControl.Enabled: End Property
Attribute Enabled.VB_Description = "Set whether or not the control is enabled."
Public Property Let Enabled(ByVal value As Boolean)
UserControl.Enabled = value
If hTVD Then EnableWindow hTVD, IIf(value, 1, 0)
UserControl.PropertyChanged "Enabled"
End Property
Public Property Get Checkboxes() As Boolean: Checkboxes = mCheckboxes: End Property
Attribute Checkboxes.VB_Description = "Show checkboxes next to the item names."
Public Property Let Checkboxes(bVal As Boolean)
If bVal <> mCheckboxes Then
    mCheckboxes = bVal
    If hTVD = 0& Then Exit Property
    Dim dwStyleEx As ucst_TV_Ex_Styles
    
    dwStyleEx = TVS_EX_PARTIALCHECKBOXES
    If mCheckboxes Then
        Call SendMessage(hTVD, TVM_SETEXTENDEDSTYLE, dwStyleEx, ByVal dwStyleEx)
    Else
        Call SendMessage(hTVD, TVM_SETEXTENDEDSTYLE, dwStyleEx, ByVal 0&)
    End If
End If
End Property
Public Property Get ExclusionChecks() As Boolean: ExclusionChecks = mExCheckboxes: End Property
Attribute ExclusionChecks.VB_Description = "Adds an additional checkbox state, a red x. Checkboxes must be enabled for this to have an effect."
Public Property Let ExclusionChecks(bVal As Boolean)
If bVal <> mExCheckboxes Then
    mExCheckboxes = bVal
    If mCheckboxes = False Then Exit Property
    If mExCheckboxes Then
        Dim dwStyleEx As ucst_TV_Ex_Styles
        
        dwStyleEx = TVS_EX_EXCLUSIONCHECKBOXES
        If mCheckboxes Then
            Call SendMessage(hTVD, TVM_SETEXTENDEDSTYLE, dwStyleEx, ByVal dwStyleEx)
        Else
            Call SendMessage(hTVD, TVM_SETEXTENDEDSTYLE, dwStyleEx, ByVal 0&)
        End If
    End If
End If
End Property
Public Property Get ShowSelAlways() As Boolean: ShowSelAlways = mShowSelAlw: End Property
Attribute ShowSelAlways.VB_Description = "Show selected item(s) when the control is out of focus."
Public Property Let ShowSelAlways(bVal As Boolean)
If mShowSelAlw <> bVal Then
    mShowSelAlw = bVal
    Dim dwStyle As ucst_TV_Styles
    dwStyle = GetWindowLong(hTVD, GWL_STYLE)
    If mShowSelAlw Then
        dwStyle = dwStyle Or TVS_SHOWSELALWAYS
    Else
        dwStyle = dwStyle And Not TVS_SHOWSELALWAYS
    End If
    Call SetWindowLong(hTVD, GWL_STYLE, dwStyle)
End If
End Property
Public Property Get TrackSelect() As Boolean: TrackSelect = m_TrackSel: End Property
Attribute TrackSelect.VB_Description = "Set whether hot tracking is enabled."
Public Property Let TrackSelect(ByVal bVal As Boolean)
If m_TrackSel <> bVal Then
    m_TrackSel = bVal
    Dim dwStyle As ucst_TV_Styles
    dwStyle = GetWindowLong(hTVD, GWL_STYLE)
    If m_TrackSel Then
        dwStyle = dwStyle Or TVS_TRACKSELECT
    Else
        dwStyle = dwStyle And Not TVS_TRACKSELECT
    End If
    Call SetWindowLong(hTVD, GWL_STYLE, dwStyle)
End If
End Property
Public Property Get MultiSelect() As Boolean: MultiSelect = mMultiSel: End Property
Attribute MultiSelect.VB_Description = "Enables the Mutlselect extended style. WARNING: This style is listed as 'Not supported' by MSDN, so is subject to no longer working at any time."
Public Property Let MultiSelect(bVal As Boolean)
Dim dwStyleEx As ucst_TV_Ex_Styles
If bVal <> mMultiSel Then
    mMultiSel = bVal
    dwStyleEx = TVS_EX_MULTISELECT
    If mMultiSel Then
        Call SendMessage(hTVD, TVM_SETEXTENDEDSTYLE, dwStyleEx, ByVal dwStyleEx)
    Else
        Call SendMessage(hTVD, TVM_SETEXTENDEDSTYLE, dwStyleEx, ByVal 0&)
    End If
End If
End Property
Public Property Get DisableDragDrop() As Boolean: DisableDragDrop = mDisableDD: End Property
Attribute DisableDragDrop.VB_Description = "Disables drag and drop operations."
Public Property Let DisableDragDrop(bVal As Boolean)
If bVal <> mDisableDD Then
    mDisableDD = bVal
    Dim dwStyle As ucst_TV_Styles
    dwStyle = GetWindowLong(hTVD, GWL_STYLE)
    If mDisableDD Then
        dwStyle = dwStyle Or TVS_DISABLEDRAGDROP
    Else
        dwStyle = dwStyle And Not TVS_DISABLEDRAGDROP
    End If
    SetWindowLong hTVD, GWL_STYLE, dwStyle
End If
End Property
Public Property Get Autocheck() As Boolean: Autocheck = mAutocheck: End Property
Attribute Autocheck.VB_Description = "Change parent and child states to reflect checkbox states."
Public Property Let Autocheck(bVal As Boolean)
If bVal <> mAutocheck Then
    mAutocheck = bVal
    If mCheckboxes = False Then Exit Property
    If mAutocheck Then
        Dim dwStyleEx As ucst_TV_Ex_Styles
        Dim dwStyle As ucst_TV_Styles
        
        If mAutocheck Then
            If mExCheckboxes Then
                dwStyleEx = TVS_EX_EXCLUSIONCHECKBOXES
            Else
                dwStyle = GetWindowLong(hTVD, GWL_STYLE)
                dwStyle = dwStyle And Not TVS_CHECKBOXES
                SetWindowLong hTVD, GWL_STYLE, dwStyle
            End If
            dwStyleEx = dwStyleEx Or TVS_EX_PARTIALCHECKBOXES
            Call SendMessage(hTVD, TVM_SETEXTENDEDSTYLE, dwStyleEx, ByVal dwStyleEx)
        Else
            dwStyleEx = TVS_EX_PARTIALCHECKBOXES
            If mExCheckboxes Then
                dwStyleEx = dwStyleEx Or TVS_EX_EXCLUSIONCHECKBOXES
                Call SendMessage(hTVD, TVM_SETEXTENDEDSTYLE, dwStyleEx, ByVal 0&)
            Else
                Call SendMessage(hTVD, TVM_SETEXTENDEDSTYLE, dwStyleEx, ByVal 0&)
                dwStyle = GetWindowLong(hTVD, GWL_STYLE)
                dwStyle = dwStyle Or TVS_CHECKBOXES
                SetWindowLong hTVD, GWL_STYLE, dwStyle
            End If
        End If
    End If
End If
End Property
Public Property Get ExplorerStyle() As Boolean: ExplorerStyle = mExplorerStyle: End Property
Attribute ExplorerStyle.VB_Description = "Apply the Explorer visual style to the tree."
Public Property Let ExplorerStyle(bVal As Boolean)
If bVal <> mExplorerStyle Then
    mExplorerStyle = bVal
    If mExplorerStyle Then
        SetWindowTheme hTVD, StrPtr("explorer"), 0&
    Else
        SetWindowTheme hTVD, StrPtr(" "), 0&
    End If
End If
End Property
Public Property Get ShowFiles() As Boolean: ShowFiles = mShowFiles: End Property
Attribute ShowFiles.VB_Description = "Show files in the tree in addition to folders."
Public Property Let ShowFiles(bVal As Boolean): mShowFiles = bVal: End Property
Public Property Get ExtendedOverlays() As Boolean: ExtendedOverlays = mExtOverlay: End Property
Attribute ExtendedOverlays.VB_Description = "Show the extended overlays sometimes found in programs like TortoiseSVN, Dropbox, or Github. WARNING: Extreme performance cost, reduces folder load times by 10-100x."
Public Property Let ExtendedOverlays(bVal As Boolean): mExtOverlay = bVal: End Property
Public Property Get FadingExpandos() As Boolean: FadingExpandos = mFadeExpandos: End Property
Attribute FadingExpandos.VB_Description = "Toggle whether or not the expando buttons fade in/fade out as focus changes."
Public Property Let FadingExpandos(bVal As Boolean)
If bVal <> mFadeExpandos Then
    mFadeExpandos = bVal
    Dim dwStyleEx As ucst_TV_Ex_Styles
    dwStyleEx = TVS_EX_FADEINOUTEXPANDOS
    If mFadeExpandos Then
        Call SendMessage(hTVD, TVM_SETEXTENDEDSTYLE, dwStyleEx, ByVal dwStyleEx)
    Else
        Call SendMessage(hTVD, TVM_SETEXTENDEDSTYLE, dwStyleEx, ByVal 0&)
    End If
End If
End Property
Public Property Get NoIndentState() As Boolean: NoIndentState = mNoIndState: End Property
Attribute NoIndentState.VB_Description = "Do not indent the tree view for the expando buttons."
Public Property Let NoIndentState(bVal As Boolean)
If bVal <> mNoIndState Then
    mNoIndState = bVal
    Dim dwStyleEx As ucst_TV_Ex_Styles
    dwStyleEx = TVS_EX_NOINDENTSTATE
    If mNoIndState Then
        Call SendMessage(hTVD, TVM_SETEXTENDEDSTYLE, dwStyleEx, ByVal dwStyleEx)
    Else
        Call SendMessage(hTVD, TVM_SETEXTENDEDSTYLE, dwStyleEx, ByVal 0&)
    End If
End If
End Property
Public Property Get AutoHScroll() As Boolean: AutoHScroll = mAutoHS: End Property
Attribute AutoHScroll.VB_Description = "Remove the horizontal scrollbar for some positions."
Public Property Let AutoHScroll(bVal As Boolean)
If bVal <> mAutoHS Then
    mAutoHS = bVal
    Dim dwStyleEx As ucst_TV_Ex_Styles
    dwStyleEx = TVS_EX_AUTOHSCROLL
    If mAutoHS Then
        Call SendMessage(hTVD, TVM_SETEXTENDEDSTYLE, dwStyleEx, ByVal dwStyleEx)
    Else
        Call SendMessage(hTVD, TVM_SETEXTENDEDSTYLE, dwStyleEx, ByVal 0&)
    End If
End If
End Property
Public Property Get FullRowSelect() As Boolean: FullRowSelect = mFullRowSelect: End Property
Attribute FullRowSelect.VB_Description = "Hilite the entire row instead of just the item name."
Public Property Let FullRowSelect(bVal As Boolean)
If bVal <> mFullRowSelect Then
    mFullRowSelect = bVal
    If hTVD = 0& Then Exit Property
    Dim dwStyle As ucst_TV_Styles
    dwStyle = GetWindowLong(hTVD, GWL_STYLE)
    If mFullRowSelect Then
        dwStyle = dwStyle Or TVS_FULLROWSELECT
    Else
        dwStyle = dwStyle And Not TVS_FULLROWSELECT
    End If
    SetWindowLong hTVD, GWL_STYLE, dwStyle
End If
End Property
Public Property Get ComputerAsRoot() As Boolean: ComputerAsRoot = mComputerAsRoot: End Property
Attribute ComputerAsRoot.VB_Description = "Use the Computer folder (drive list) instead of the Desktop as the root of the tree."
Public Property Let ComputerAsRoot(bVal As Boolean)
If mComputerAsRoot <> bVal Then
    mComputerAsRoot = bVal
    If Ambient.UserMode Then ResetTreeView
End If
End Property
Public Property Get SingleExpand() As Boolean: SingleExpand = mSingleExpand: End Property
Attribute SingleExpand.VB_Description = "Allow only one item to be expanded at a time. The previous item is collapsed when a new one expands."
Public Property Let SingleExpand(bVal As Boolean)
If bVal <> mSingleExpand Then
    mSingleExpand = bVal
    If hTVD = 0& Then Exit Property
    Dim dwStyle As ucst_TV_Styles
    dwStyle = GetWindowLong(hTVD, GWL_STYLE)
    If mSingleExpand Then
        dwStyle = dwStyle Or TVS_SINGLEEXPAND
    Else
        dwStyle = dwStyle And Not TVS_SINGLEEXPAND
    End If
    SetWindowLong hTVD, GWL_STYLE, dwStyle
End If
End Property
Public Property Get HasButtons() As Boolean: HasButtons = mHasButtons: End Property
Attribute HasButtons.VB_Description = "Sets whether expando buttons appear next to items with subfolders."
Public Property Let HasButtons(bVal As Boolean)
If bVal <> mHasButtons Then
    mHasButtons = bVal
    If hTVD = 0& Then Exit Property
    Dim dwStyle As ucst_TV_Styles
    dwStyle = GetWindowLong(hTVD, GWL_STYLE)
    If mHasButtons Then
        dwStyle = dwStyle Or TVS_HASBUTTONS
    Else
        dwStyle = dwStyle And Not TVS_HASBUTTONS
    End If
    SetWindowLong hTVD, GWL_STYLE, dwStyle
End If
End Property
Public Property Get ShowLines() As Boolean: ShowLines = mShowLines: End Property
Attribute ShowLines.VB_Description = "Show lines connecting the levels of the tree. Works best when ExplorerStyle is disabled."
Public Property Let ShowLines(bVal As Boolean)
If bVal <> mShowLines Then
    mShowLines = bVal
    If hTVD = 0& Then Exit Property
    Dim dwStyle As ucst_TV_Styles
    dwStyle = GetWindowLong(hTVD, GWL_STYLE)
    If mShowLines Then
        dwStyle = dwStyle Or TVS_HASLINES
    Else
        dwStyle = dwStyle And Not TVS_HASLINES
    End If
    SetWindowLong hTVD, GWL_STYLE, dwStyle
End If
End Property
Public Property Get ExpandZip() As Boolean: ExpandZip = mExpandZip: End Property
Attribute ExpandZip.VB_Description = "Display/expand zip/cab files like a folder."
Public Property Let ExpandZip(bVal As Boolean): mExpandZip = bVal: End Property
Public Property Get ItemFilter() As String: ItemFilter = mFilter: End Property
Attribute ItemFilter.VB_Description = "Set a filter such that item names not matching are not shown in the tree."
Public Property Let ItemFilter(sFilter As String): mFilter = sFilter: End Property
Public Property Get ItemFilterFilesOnly() As Boolean: ItemFilterFilesOnly = mFilterFilesOnly: End Property
Attribute ItemFilterFilesOnly.VB_Description = "Apply the ItemFilter restriction only to files (ShowFiles), not folders."
Public Property Let ItemFilterFilesOnly(bVal As Boolean): mFilterFilesOnly = bVal: End Property
Public Property Get InfoTipOnFolders() As Boolean: InfoTipOnFolders = mInfoTipOnFolders: End Property
Attribute InfoTipOnFolders.VB_Description = "Show a tooltip with item information on mouseover."
Public Property Let InfoTipOnFolders(bVal As Boolean): mInfoTipOnFolders = bVal: End Property
Public Property Get InfoTipOnFiles() As Boolean: InfoTipOnFiles = mInfoTipOnFiles: End Property
Attribute InfoTipOnFiles.VB_Description = "Show a tooltip with several properties based on the file type on mouseover."
Public Property Let InfoTipOnFiles(bVal As Boolean): mInfoTipOnFiles = bVal: End Property
Public Property Get ActiveDropHoverTime() As Long: ActiveDropHoverTime = lRaiseHover: End Property
Attribute ActiveDropHoverTime.VB_Description = "During a drag/drop operation, how long (in ms) the mouse must not move to trigger a hover event."
Public Property Let ActiveDropHoverTime(lMillisecs As Long): lRaiseHover = lMillisecs: End Property
Public Property Get SingleClickExpand() As Boolean: SingleClickExpand = mExpandOnLabelClick: End Property
Attribute SingleClickExpand.VB_Description = "Expand/collapse when the item name is clicked; otherwise, you must click the expando button or double-click the name to expand/collapse."
Public Property Let SingleClickExpand(bVal As Boolean): mExpandOnLabelClick = bVal: End Property
Public Property Get PlayNavigationSound() As Boolean: PlayNavigationSound = mNavSound: End Property
Attribute PlayNavigationSound.VB_Description = "Play the click sound you hear in Explorer when selecting a new item."
Public Property Let PlayNavigationSound(bVal As Boolean): mNavSound = bVal: End Property
Public Property Get AlwaysShowExtendedVerbs() As Boolean: AlwaysShowExtendedVerbs = mAlwaysShowExtVerbs: End Property
Attribute AlwaysShowExtendedVerbs.VB_Description = "Always show extended verbs in the shell context menu, otherwise Shift must be pressed when bringing up the menu."
Public Property Let AlwaysShowExtendedVerbs(bVal As Boolean): mAlwaysShowExtVerbs = bVal: End Property
Public Property Get ShowHiddenItems() As ST_HDN_PREF: ShowHiddenItems = m_HiddenPref: End Property
Public Property Let ShowHiddenItems(lVal As ST_HDN_PREF): m_HiddenPref = lVal: End Property
Public Property Get ShowSuperHidden() As ST_SPRHDN_PREF: ShowSuperHidden = m_SuperHiddenPref: End Property
Public Property Let ShowSuperHidden(lVal As ST_SPRHDN_PREF): m_SuperHiddenPref = lVal: End Property
Public Property Get ShowFavorites() As Boolean: ShowFavorites = mFavorites: End Property
Attribute ShowFavorites.VB_Description = "Show the 'Favorites' link group at the top like Explorer."
Public Property Let ShowFavorites(bVal As Boolean)
If mFavorites <> bVal Then
    mFavorites = bVal
    If mFavorites Then
        If Ambient.UserMode Then ResetTreeView
    End If
End If
End Property
Public Property Get EnableShellMenu() As Boolean: EnableShellMenu = m_EnableShellMenu: End Property
Public Property Let EnableShellMenu(bVal As Boolean)
m_EnableShellMenu = bVal
'UserControl.PropertyChanged "EnableShellMenu"
End Property
Public Property Get SelectedItem() As String: SelectedItem = sSelectedItem: End Property
Public Property Get SelectedShellItem() As oleexp.IShellItem: Set SelectedShellItem = siSelected: End Property
Public Property Get hWndTreeView() As Long: hWndTreeView = hTVD: End Property
Public Property Get hWndUserControl() As Long: hWndUserControl = UserControl.hWnd: End Property
'Public Property Get Border() As Boolean: Border = mBorder: End Property
'Public Property Let Border(bVal As Boolean): mBorder = bVal: End Property
Public Property Get BorderStyle() As ST_BORDERSTYLE: BorderStyle = mBorder: End Property
Attribute BorderStyle.VB_Description = "Provides options for the type of border around the tree."
Public Property Let BorderStyle(lVal As ST_BORDERSTYLE)
If lVal <> mBorder Then
    mBorder = lVal
    Dim dwStyle As WindowStyles, dwExStyle As WindowStylesEx
    dwStyle = GetWindowLong(hTVD, GWL_STYLE)
    dwExStyle = GetWindowLong(hTVD, GWL_EXSTYLE)
    Select Case mBorder
        Case STBS_None
            dwStyle = dwStyle And Not WS_BORDER
            dwStyle = dwStyle And Not WS_THICKFRAME
            dwExStyle = dwExStyle And Not WS_EX_CLIENTEDGE
        Case STBS_Standard
            dwStyle = dwStyle Or WS_BORDER
            dwStyle = dwStyle And Not WS_THICKFRAME
            dwExStyle = dwExStyle And Not WS_EX_CLIENTEDGE
        Case STBS_Thick
            dwStyle = dwStyle Or WS_BORDER
            dwStyle = dwStyle And Not WS_THICKFRAME
            dwExStyle = dwExStyle Or WS_EX_CLIENTEDGE
        Case STBS_Thicker
            dwStyle = dwStyle Or WS_BORDER
            dwStyle = dwStyle Or WS_THICKFRAME
            dwExStyle = dwExStyle Or WS_EX_CLIENTEDGE
    End Select
    Call SetWindowLong(hTVD, GWL_STYLE, dwStyle)
    Call SetWindowLong(hTVD, GWL_EXSTYLE, dwExStyle)
    SetWindowPos hTVD, 0&, 0&, 0&, 0&, 0&, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_FRAMECHANGED
End If
End Property
Public Property Get MonitorDirChanges() As Boolean: MonitorDirChanges = mSHCN: End Property
Attribute MonitorDirChanges.VB_Description = "Monitor shell notifications for files/folders created, deleted, renamed, etc, and update the tree accordingly if needed."
Public Property Let MonitorDirChanges(bVal As Boolean)
If bVal <> mSHCN Then
    If bVal Then
        mSHCN = True
        StartNotify UserControl.hWnd
    Else
        mSHCN = False
        StopNotify
    End If
End If
End Property
Public Property Get LabelEditRename() As Boolean: LabelEditRename = mLabelEdit: End Property
Attribute LabelEditRename.VB_Description = "Enables label edit, which renames the folders/files in Windows as well."
Public Property Let LabelEditRename(bVal As Boolean)
If bVal <> mLabelEdit Then
    mLabelEdit = bVal
    If hTVD = 0& Then Exit Property
    Dim dwStyle As ucst_TV_Styles
    dwStyle = GetWindowLong(hTVD, GWL_STYLE)
    If mLabelEdit Then
        dwStyle = dwStyle Or TVS_EDITLABELS
    Else
        dwStyle = dwStyle And Not TVS_EDITLABELS
    End If
    SetWindowLong hTVD, GWL_STYLE, dwStyle
End If
End Property
Public Property Get HorizontalScroll() As Boolean: HorizontalScroll = mHScroll: End Property
Attribute HorizontalScroll.VB_Description = "Sets whether the tree expands such that horizontal scrolling isn't needed."
Public Property Let HorizontalScroll(bVal As Boolean)
If bVal <> mHScroll Then
    mHScroll = bVal
    If hTVD = 0& Then Exit Property
    Dim dwStyle As ucst_TV_Styles
    dwStyle = GetWindowLong(hTVD, GWL_STYLE)
    If mHScroll Then
        dwStyle = dwStyle And Not TVS_NOHSCROLL
    Else
        dwStyle = dwStyle Or TVS_NOHSCROLL
    End If
    SetWindowLong hTVD, GWL_STYLE, dwStyle
End If
End Property

Public Property Get NameColors() As Boolean: NameColors = mNameColors: End Property
Attribute NameColors.VB_Description = "Show Windows encrypted files in green, and Windows compressed files in blue. These are the attributes set in file properties."
Public Property Let NameColors(bVal As Boolean): mNameColors = bVal: End Property
Public Property Get InitialPath() As String: InitialPath = mInitialPath: End Property
Attribute InitialPath.VB_Description = "Deprecated. Use OpenToPath."
Attribute InitialPath.VB_MemberFlags = "400"
Public Property Let InitialPath(sPath As String)
mInitialPath = sPath
If sPath <> "" Then
    OpenToPath sPath, False
End If
End Property
Public Property Get CustomRoot() As String: CustomRoot = mCustomRoot: End Property
Attribute CustomRoot.VB_Description = "Specify a custom location to use as the root folder."
Public Property Let CustomRoot(sRoot As String)
mCustomRoot = sRoot
ResetTreeView
End Property
Public Property Get RootHasCheckbox() As Boolean: RootHasCheckbox = mRootHasCheckbox: End Property
Attribute RootHasCheckbox.VB_Description = "Sets whether or not a checkbox appears next to the root item when checkboxes are enabled."
Public Property Let RootHasCheckbox(bVal As Boolean)
If bVal <> mRootHasCheckbox Then
    mRootHasCheckbox = bVal
    bSetParents = True
    If mRootHasCheckbox Then
        If Ambient.UserMode Then
            If TVNodeHasUncheckedChildren(m_hRoot) = False Then
                TreeView_SetCheckStateEx hTVD, m_hRoot, 2
            Else
                If TVNodeHasCheckedChildren(m_hRoot) Then
                    TreeView_SetCheckStateEx hTVD, m_hRoot, 3
                Else
                    TreeView_SetCheckStateEx hTVD, m_hRoot, 1
                End If
            End If
        Else
            TreeView_SetCheckStateEx hTVD, m_hRoot, 1
        End If
    Else
        TreeView_SetCheckStateEx hTVD, m_hRoot, 0
    End If
    bSetParents = False
End If
End Property
Public Property Get CheckedPaths() As Variant
    SetCheckedList
    CheckedPaths = gPaths
End Property
Public Property Get ExcludedPaths() As Variant
    SetExCheckedList
    ExcludedPaths = gExPaths
End Property 'End Control Properties
'********************************************************

Public Sub ResetTreeView()
'Removes all items and inserts from root
If Ambient.UserMode Then
    EmptyTreeView hTVD, True
    fRefreshing = 1
    bFilling = True
    EnumRoot
    bFilling = False
    fRefreshing = 0
End If
End Sub

Public Sub RefreshTreeView()
'Resets to the root then expands all the previous locations
'The lengthy code involves minimizing the OpenToPath calls by calculating
'smallest group of folders that will give us the same setup. E.g.:
'C:\a\b\(c,d,e,f) will call only C:\a\b\c, and not d,e,f or C:\a or b, because
'expanding to c will give us a b d e and f too
Dim i As Long
Dim sData() As String

If CalcRefreshData(sData) Then
    ResetTreeView
    
    For i = 0 To UBound(sData)
        OpenToPath sData(i), False
    Next i
End If

Exit Sub
e0:
    DebugAppend "PruneParentsFromRPL.Error->" & Err.Description & "(0x" & Hex$(Err.Number) & ")"

End Sub

Private Function CalcRefreshData(ByRef sList() As String) As Boolean
    '<EhHeader>
    On Error GoTo e0
    '</EhHeader>
Dim i As Long, j As Long, k As Long
On Error GoTo e0
ReDim mRefreshPaths(0)

For i = 0 To UBound(TVEntries)
    If TVEntries(i).bDeleted = False Then
        If TVEntries(i).bFolder Then
            ReDim Preserve mRefreshPaths(j)
            mRefreshPaths(j) = TVEntries(i).sFullPath
            j = j + 1
        End If
    End If
Next i
PruneParentsFromRPL
Dim sTmp As String
Dim sExp() As String, nE As Long
Dim bA As Boolean
ReDim sExp(0)
For i = 0 To UBound(mRefreshPaths)
    bA = True
    For j = 0 To UBound(TVEntries)
        If TVEntries(j).sFullPath = mRefreshPaths(i) Then
            If (TVEntries(j).bIsDefItem = True) Then
                bA = False
            End If
        End If
    Next j
    sTmp = mRefreshPaths(i)
    For k = 0 To UBound(mRefreshPaths)
        If (Left$(mRefreshPaths(k), Len(sTmp)) = sTmp) And (Len(mRefreshPaths(k)) > Len(sTmp)) Then
            bA = False
        End If
    Next k
    If bA Then
        ReDim Preserve sExp(nE)
        sExp(nE) = mRefreshPaths(i)
        nE = nE + 1
    End If
Next i
Dim sExp2() As String, nE2 As Long
ReDim sExp2(0)
For i = 0 To UBound(sExp)
    bA = True
    sTmp = GetPathParent(sExp(i))
    For k = 0 To UBound(sExp)
        If (Left$(sExp(k), Len(sTmp)) = sTmp) And (i <> k) Then
            bA = False
        End If
    Next k
    If bA Then
        ReDim Preserve sExp2(nE2)
        sExp2(nE2) = sExp(i)
        nE2 = nE2 + 1
    End If
Next i
DebugAppend "Final set for re-expansion:"
dbg_printstrarr sExp2, False
If (UBound(sExp2) = 0&) And (sExp2(0) = "") Then
    CalcRefreshData = False
Else
    sList = sExp2
    CalcRefreshData = True
End If


'<EhFooter>
Exit Function

e0:
    DebugAppend "ucShellTree.CalcRefreshData->Error: " & Err.Description & ", 0x" & Hex$(Err.Number)
'</EhFooter>
End Function

Private Sub PruneParentsFromRPL()
Dim sCur As String, sPar As String
Dim i As Long, j As Long, k As Long
Dim sOut() As String
Dim nOut As Long
On Error GoTo e0

For i = 0 To UBound(mRefreshPaths)
    sPar = GetPathParent(mRefreshPaths(i))
    If (PathHasChildInRPL(sPar) = False) Then
        mRefreshPaths(i) = ""
    End If
Next i
ReDim sOut(0)
For i = 0 To UBound(mRefreshPaths)
    If mRefreshPaths(i) <> "" Then
        ReDim Preserve sOut(nOut)
        sOut(nOut) = mRefreshPaths(i)
        nOut = nOut + 1
    End If
Next i
Dim aFn() As String, kF As Long
ReDim aFn(0)
For i = 0 To UBound(sOut)
    sPar = GetPathParent(sOut(i))
    For j = 0 To UBound(sOut)
        If j <> i Then
            If GetPathParent(sOut(j)) = sPar Then
                sOut(j) = ""
            End If
        End If
    Next j
Next i
ReDim aFn(0)
For i = 0 To UBound(sOut)
    If sOut(i) <> "" Then
        ReDim Preserve aFn(kF)
        aFn(kF) = sOut(i)
        kF = kF + 1
    End If
Next i
mRefreshPaths = aFn
Exit Sub
e0:
    DebugAppend "PruneParentsFromRPL.Error->" & Err.Description & "(0x" & Hex$(Err.Number) & ")"
End Sub

Private Function PathHasChildInRPL(sPath As String) As Boolean
Dim sCur As String, sPar As String
Dim i As Long, j As Long, k As Long

For i = 0 To UBound(mRefreshPaths)
    If Left$(mRefreshPaths(i), Len(sPath)) = sPath Then
        If Len(mRefreshPaths(i)) > Len(sPath) Then
            PathHasChildInRPL = True
            Exit Function
        End If
    End If
Next i

End Function

Private Function GetPathParent(sPath As String) As String
If Len(sPath) > 3 Then
    If InStr(sPath, "\") Then
        GetPathParent = Left$(sPath, InStrRev(sPath, "\") - 1)
    End If
End If
If Len(GetPathParent) = 2 Then
    GetPathParent = GetPathParent & "\"
End If
End Function

Private Sub UserControl_GotFocus()
SetFocus hTVD
End Sub

Private Sub UserControl_EnterFocus()
RaiseEvent EnterFocus
End Sub

Private Sub UserControl_ExitFocus()
RaiseEvent ExitFocus
End Sub

Private Sub UserControl_InitProperties()
clrBack = vbWindowBackground
clrFore = vbWindowText
Set PropFont = Ambient.Font
mInitialPath = mInitialPath_def
mCheckboxes = mCheckboxes_def
mExCheckboxes = mExCheckboxes_def
mExplorerStyle = mExplorerStyle_def
mShowFiles = mShowFiles_def
mFadeExpandos = mFadeExpandos_def
mShowSelAlw = mDefShowSelAlw
mShowLines = mShowLines_def
mHasButtons = mHasButtons_def
mSingleExpand = mSingleExpand_def
mComputerAsRoot = mComputerAsRoot_def
mFullRowSelect = mFullRowSelect_def
mExpandZip = mExpandZip_def
mFilter = mFilter_def
mFilterFilesOnly = mFilterFilesOnly_def
mInfoTipOnFolders = mInfoTipOnFolders_def
mInfoTipOnFiles = mInfoTipOnFiles_def
mExpandOnLabelClick = mExpandOnLabelClick_def
mNavSound = mNavSound_def
mFavorites = mFavorites_def
mSHCN = mSHCN_def
mLabelEdit = mLabelEdit_def
mNameColors = mNameColors_def
mBorder = mBorder_def
lRaiseHover = m_def_lRaiseHover
m_TrackSel = m_def_TrackSel
mCustomRoot = mCustomRoot_def
mRootHasCheckbox = mRootHasCheckbox_def
mHScroll = mHScroll_def
mAutocheck = mAutocheck_def
mExtOverlay = m_def_ExtOverlay
mMultiSel = mDefMultiSel
mDisableDD = mDefDisableDD
mAlwaysShowExtVerbs = mDefAlwaysShowExtVerbs
mNoIndState = mDefNoIndState
mDefExpComp = True
mAutoHS = mDefAutoHS
m_HiddenPref = m_def_HiddenPref
m_SuperHiddenPref = m_def_SuperHiddenPref
m_EnableShellMenu = m_def_EnableShellMenu
DebugAppend "InitProps"
pvCreate
End Sub

Private Sub UserControl_ReadProperties(propBag As PropertyBag)
clrBack = propBag.ReadProperty("BackColor", vbWindowBackground)
clrFore = propBag.ReadProperty("ForeColor", vbWindowText)
mBorder = propBag.ReadProperty("BorderStyle", mBorder_def)
Set PropFont = propBag.ReadProperty("Font", Nothing)
mCheckboxes = propBag.ReadProperty("Checkboxes", mCheckboxes_def)
mExCheckboxes = propBag.ReadProperty("ExclusionChecks", mExCheckboxes_def)
mAutoHS = propBag.ReadProperty("AutoHScroll", mDefAutoHS)
mShowSelAlw = propBag.ReadProperty("ShowSelAlways", mDefShowSelAlw)
mExplorerStyle = propBag.ReadProperty("ExplorerStyle", mExplorerStyle_def)
mShowFiles = propBag.ReadProperty("ShowFiles", mShowFiles_def)
m_TrackSel = propBag.ReadProperty("TrackSelect", m_def_TrackSel)
mFadeExpandos = propBag.ReadProperty("FadingExpandos", mFadeExpandos_def)
mShowLines = propBag.ReadProperty("ShowLines", mShowLines_def)
mHasButtons = propBag.ReadProperty("HasButtons", mHasButtons_def)
mSingleExpand = propBag.ReadProperty("SingleExpand", mSingleExpand_def)
mComputerAsRoot = propBag.ReadProperty("ComputerAsRoot", mComputerAsRoot_def)
mFullRowSelect = propBag.ReadProperty("FullRowSelect", mFullRowSelect_def)
mExpandZip = propBag.ReadProperty("ExpandZip", mExpandZip_def)
mFilter = propBag.ReadProperty("ItemFilter", mFilter_def)
mFilterFilesOnly = propBag.ReadProperty("ItemFilterFilesOnly", mFilterFilesOnly_def)
mInfoTipOnFolders = propBag.ReadProperty("InfoTipOnFolders", mInfoTipOnFolders_def)
mInfoTipOnFiles = propBag.ReadProperty("InfoTipOnFiles", mInfoTipOnFiles_def)
mExpandOnLabelClick = propBag.ReadProperty("SingleClickExpand", mExpandOnLabelClick_def)
mNavSound = propBag.ReadProperty("PlayNavigationSound", mNavSound_def)
mFavorites = propBag.ReadProperty("ShowFavorites", mFavorites_def)
mSHCN = propBag.ReadProperty("MonitorDirChanges", mSHCN_def)
lRaiseHover = propBag.ReadProperty("ActiveDropHoverTime", m_def_lRaiseHover)
mLabelEdit = propBag.ReadProperty("LabelEditRename", mLabelEdit_def)
mNameColors = propBag.ReadProperty("NameColors", mNameColors_def)
mBorder = propBag.ReadProperty("Border", mBorder_def)
mInitialPath = propBag.ReadProperty("InitialPath", mInitialPath_def)
mCustomRoot = propBag.ReadProperty("CustomRoot", mCustomRoot_def)
mRootHasCheckbox = propBag.ReadProperty("RootHasCheckbox", mRootHasCheckbox_def)
mHScroll = propBag.ReadProperty("HorizontalScroll", mHScroll_def)
mAutocheck = propBag.ReadProperty("Autocheck", mAutocheck_def)
mExtOverlay = propBag.ReadProperty("ExtendedOverlays", m_def_ExtOverlay)
mMultiSel = propBag.ReadProperty("Multiselect", mDefMultiSel)
mDisableDD = propBag.ReadProperty("DisableDragDrop", mDefDisableDD)
mAlwaysShowExtVerbs = propBag.ReadProperty("AlwaysShowExtendedVerbs", mDefAlwaysShowExtVerbs)
mNoIndState = propBag.ReadProperty("NoIndentState", mDefNoIndState)
m_HiddenPref = propBag.ReadProperty("ShowHiddenItems", m_def_HiddenPref)
m_SuperHiddenPref = propBag.ReadProperty("ShowSuperHidden", m_def_SuperHiddenPref)
m_EnableShellMenu = propBag.ReadProperty("EnableShellMenu", m_def_EnableShellMenu)
mDefExpComp = True
pvCreate
End Sub

Private Sub UserControl_Show()
'DebugAppend "IPAO..."
pvInitIPAO
DebugAppend "Init IPAO Ok. Init TV..."

End Sub

Private Sub UserControl_WriteProperties(propBag As PropertyBag)
propBag.WriteProperty "BackColor", clrBack, vbWindowBackground
propBag.WriteProperty "ForeColor", clrFore, vbWindowText
propBag.WriteProperty "Font", IIf(OLEFontIsEqual(PropFont, Ambient.Font) = False, PropFont, Nothing), Nothing
propBag.WriteProperty "Checkboxes", mCheckboxes, mCheckboxes_def
propBag.WriteProperty "ExclusionChecks", mExCheckboxes, mExCheckboxes_def
propBag.WriteProperty "ShowSelAlways", mShowSelAlw, mDefShowSelAlw
propBag.WriteProperty "ExplorerStyle", mExplorerStyle, mExplorerStyle_def
propBag.WriteProperty "ShowFiles", mShowFiles, mShowFiles_def
propBag.WriteProperty "FadingExpandos", mFadeExpandos, mFadeExpandos_def
propBag.WriteProperty "ShowLines", mShowLines, mShowLines_def
propBag.WriteProperty "HasButtons", mHasButtons, mHasButtons_def
propBag.WriteProperty "TrackSelect", m_TrackSel, m_def_TrackSel
propBag.WriteProperty "SingleExpand", mSingleExpand, mSingleExpand_def
propBag.WriteProperty "ComputerAsRoot", mComputerAsRoot, mComputerAsRoot_def
propBag.WriteProperty "FullRowSelect", mFullRowSelect, mFullRowSelect_def
propBag.WriteProperty "ExpandZip", mExpandZip, mExpandZip_def
propBag.WriteProperty "ItemFilter", mFilter, mFilter_def
propBag.WriteProperty "ItemFilterFilesOnly", mFilterFilesOnly, mFilterFilesOnly_def
propBag.WriteProperty "InfoTipOnFolders", mInfoTipOnFolders, mInfoTipOnFolders_def
propBag.WriteProperty "InfoTipOnFiles", mInfoTipOnFiles, mInfoTipOnFiles_def
propBag.WriteProperty "SingleClickExpand", mExpandOnLabelClick, mExpandOnLabelClick_def
propBag.WriteProperty "PlayNavigationSound", mNavSound, mNavSound_def
propBag.WriteProperty "ShowFavorites", mFavorites, mFavorites_def
propBag.WriteProperty "MonitorDirChanges", mSHCN, mSHCN_def
propBag.WriteProperty "ActiveDropHoverTime", lRaiseHover, m_def_lRaiseHover
propBag.WriteProperty "LabelEditRename", mLabelEdit, mLabelEdit_def
propBag.WriteProperty "NameColors", mNameColors, mNameColors_def
'PropBag.WriteProperty "Border", mBorder, mBorder_def
propBag.WriteProperty "BorderStyle", mBorder, mBorder_def
propBag.WriteProperty "InitialPath", mInitialPath, mInitialPath_def
propBag.WriteProperty "CustomRoot", mCustomRoot, mCustomRoot_def
propBag.WriteProperty "RootHasCheckbox", mRootHasCheckbox, mRootHasCheckbox_def
propBag.WriteProperty "HorizontalScroll", mHScroll, mHScroll_def
propBag.WriteProperty "Autocheck", mAutocheck, mAutocheck_def
propBag.WriteProperty "ExtendedOverlays", mExtOverlay, m_def_ExtOverlay
propBag.WriteProperty "Multiselect", mMultiSel, mDefMultiSel
propBag.WriteProperty "DisableDragDrop", mDisableDD, mDefDisableDD
propBag.WriteProperty "AlwaysShowExtendedVerbs", mAlwaysShowExtVerbs, mDefAlwaysShowExtVerbs
propBag.WriteProperty "NoIndentState", mNoIndState, mDefNoIndState
propBag.WriteProperty "AutoHScroll", mAutoHS, mDefAutoHS
propBag.WriteProperty "ShowHiddenItems", m_HiddenPref, m_def_HiddenPref
propBag.WriteProperty "ShowSuperHidden", m_SuperHiddenPref, m_def_SuperHiddenPref
propBag.WriteProperty "EnableShellMenu", m_EnableShellMenu, m_def_EnableShellMenu
End Sub

Private Sub UserControl_Resize()
Dim rc As oleexp.RECT
If mBkm Then
    '...
Else
    Call GetClientRect(UserControl.hWnd, rc)
    Call SetWindowPos(hTVD, 0, 0, 0, rc.Right, rc.Bottom, 0)
End If
End Sub

Private Sub UserControl_Initialize()
IsComCtl6 = (ComCtlVersion >= 6)
gCurSelIdx = -1
Set pDTH = New oleexp.DragDropHelper
CF_SHELLIDLIST = RegisterClipboardFormatW(StrPtr(CFSTR_SHELLIDLIST))
CF_DROPDESCRIPTION = RegisterClipboardFormatW(StrPtr(CFSTR_DROPDESCRIPTION))
CF_PREFERREDDROPEFFECT = RegisterClipboardFormatW(StrPtr(CFSTR_PREFERREDDROPEFFECT))
CF_COMPUTEDDRAGIMAGE = RegisterClipboardFormatW(StrPtr(CFSTR_COMPUTEDDRAGIMAGE))
CF_INDRAGLOOP = RegisterClipboardFormatW(StrPtr(CFSTR_INDRAGLOOP))
m_SysClrText = GetSysColor(COLOR_WINDOWTEXT)
Dim lp As Long
SHGetKnownFolderPath FOLDERID_Profile, KF_FLAG_DEFAULT, 0&, lp
sUserFolder = LPWSTRtoStr(lp)
lp = 0
SHGetKnownFolderPath FOLDERID_Desktop, KF_FLAG_DEFAULT, 0&, lp
sUserDesktop = LPWSTRtoStr(lp)
ReDim sUDLastPathSet(0)
ReDim TVEntries(0)
If ExplorerSettingEnabled(SSF_SHOWALLOBJECTS) Then
    mHPInExp = True
Else
    mHPInExp = False
End If
If ExplorerSettingEnabled(SSF_SHOWSUPERHIDDEN) Then
    mSHPInExp = True
Else
    mSHPInExp = False
End If

'Dim psibk As IShellItem
'Dim isiif As IShellItemImageFactory
'oleexp.SHCreateItemFromParsingName StrPtr(App.Path & "\grain2.ico"), Nothing, IID_IShellItem, psibk
'Set isiif = psibk
'isiif.GetImage 256, 256, SIIGBF_THUMBNAILONLY, hBmpBack
'Debug.Print "hBmpBack=" & hBmpBack
End Sub

Private Sub pvCreate() 'Formerly in UserControl_Show()
If hTVD <> 0& Then Exit Sub
DebugAppend "pvCreate"
Dim pidlDesktop As Long
If Ambient.UserMode Then
    If ssc_Subclass(UserControl.hWnd, , , , , True) Then
     Call ssc_AddMsg(UserControl.hWnd, MSG_BEFORE, ALL_MESSAGES)
    End If
End If
If mFavorites Then nCur = 1
InitTV
If Ambient.UserMode Then
    
    EnumRoot
    
    SetFocus hTVD
    
    If mSHCN Then StartNotify UserControl.hWnd
Else
    Dim tVI As TVITEM
    Dim tvins As TVINSERTSTRUCT
    Dim hRt As Long, hItem As Long
    
    tVI.Mask = TVIF_IMAGE Or TVIF_TEXT Or TVIF_CHILDREN
    tVI.cChildren = 1
    tVI.pszText = StrPtr(mVersionStr)
    tVI.cchTextMax = 22
    tVI.iImage = 1
    
    tvins.hInsertAfter = 0
    tvins.hParent = TVI_ROOT
    tvins.Item = tVI
    
    hRt = SendMessage(hTVD, TVM_INSERTITEMW, 0&, tvins)
    ReDim TVEntries(0)
    
End If
bLoadDone = True
If Ambient.UserMode Then
    If (mComputerAsRoot = False) And (bCustRt = False) And (mDefExpComp = True) Then
        Dim siComp As oleexp.IShellItem
        SHGetKnownFolderItem FOLDERID_ComputerFolder, KF_FLAG_DEFAULT, 0&, IID_IShellItem, siComp
        If (siComp Is Nothing) = False Then
            OpenToItem siComp, True, False
        End If
    End If
    
    RaiseEvent Initialized
End If


End Sub

Private Sub UserControl_Terminate()
DebugAppend "ShellTree Terminate Event"
fLoad = 1
'Free full pidl records
Dim i As Long
On Error Resume Next
For i = 0 To UBound(TVEntries)
    If TVEntries(i).pidlFQPar Then
        CoTaskMemFree TVEntries(i).pidlFQPar
    End If
Next i
For i = 0 To UBound(TVEntries)
    If TVEntries(i).pidlFQ Then
        CoTaskMemFree TVEntries(i).pidlFQ
    End If
Next i
Call EmptyTreeView(hTVD, False)
  
' Detach the image lists
Call TreeView_SetImageList(hTVD, 0, TVSIL_NORMAL)
Call TreeView_SetImageList(hTVD, 0, TVSIL_STATE)
'ImageList_Destroy himlTVCheck

Set pIML = Nothing
Set pDTH = Nothing

'Revoke DragDrop
Call Detach

'Clear self-subclass/callback
Call ssc_Terminate
Call scb_TerminateCallbacks

DestroyWindow hTVD

If hFontTV <> 0 Then
    DeleteObject hFontTV
    hFontTV = 0
End If

End Sub

Private Function OLEFontIsEqual(ByVal Font As StdFont, ByVal FontOther As StdFont) As Boolean
If Font Is Nothing Then
    If FontOther Is Nothing Then OLEFontIsEqual = True
ElseIf FontOther Is Nothing Then
    If Font Is Nothing Then OLEFontIsEqual = True
Else
    If Font.Name = FontOther.Name And Font.Size = FontOther.Size And Font.Charset = FontOther.Charset And Font.Weight = FontOther.Weight And _
    Font.Underline = FontOther.Underline And Font.Italic = FontOther.Italic And Font.Strikethrough = FontOther.Strikethrough Then
        OLEFontIsEqual = True
    End If
End If
End Function

Private Function CloneOLEFont(ByVal Font As IFont) As StdFont
Font.Clone CloneOLEFont
End Function

Private Function GDIFontFromOLEFont(ByVal Font As IFont) As Long
GDIFontFromOLEFont = Font.hFont
End Function

Private Function ImageList_AddIcon(himl As Long, hIcon As Long) As Long
  ImageList_AddIcon = ImageList_ReplaceIcon(himl, -1, hIcon)
End Function

Private Function AddBackslash(s As String) As String

   If Len(s) > 0 Then
      If Right$(s, 1) <> "\" Then
         AddBackslash = s & "\"
      Else
         AddBackslash = s
      End If
   Else
      AddBackslash = "\"
   End If

End Function

Private Sub FastDoEvents()
    Dim uMsg As MsgType
    '
    Do While PeekMessage(uMsg, 0&, 0&, 0&, PM_REMOVE)   ' Reads and deletes message from queue.
        TranslateMessage uMsg                           ' Translates virtual-key messages into character messages.
        DispatchMessage uMsg                            ' Dispatches a message to a window procedure.
    Loop
End Sub

Private Function ComCtlVersion() As Long
Dim tVI As DLLVERSIONINFO
On Error Resume Next
tVI.cbSize = LenB(tVI)
If DllGetVersion(tVI) = S_OK Then ComCtlVersion = tVI.dwMajor
End Function

Private Sub HandleShellNotify(pidl1 As Long, pidl2 As Long, lEvent As SHCN_Events)
    '<EhHeader>
    On Error GoTo e0
    '</EhHeader>
'if (levent =

Dim siItem1 As oleexp.IShellItem
Dim szItem1 As String
Dim lpName As Long
Dim siItem2 As oleexp.IShellItem
Dim szItem2 As String
Dim lIdx As Long, lIdx2 As Long
Dim siPar As oleexp.IShellItem
Dim lpPar As Long, sPar As String
Dim hNode As Long, hNodePar As Long
Dim lpi1 As Long, lpi2 As Long
Dim tVI As TVITEM

If pidl1 Then
    oleexp.SHCreateItemFromIDList pidl1, IID_IShellItem, siItem1
    If (siItem1 Is Nothing) = False Then
        siItem1.GetDisplayName SIGDN_DESKTOPABSOLUTEPARSING, lpi1
        szItem1 = LPWSTRtoStr(lpi1)
    End If
End If
If pidl2 Then
    oleexp.SHCreateItemFromIDList pidl2, IID_IShellItem, siItem2
    If (siItem2 Is Nothing) = False Then
        siItem2.GetDisplayName SIGDN_DESKTOPABSOLUTEPARSING, lpi2
        szItem2 = LPWSTRtoStr(lpi2)
    End If
End If
DebugAppend "HandleShellNotify::code=" & dbg_LookUpSHCNE(lEvent) & ",itm1=" & szItem1 & ",itm2=" & szItem2
Select Case lEvent

    Case SHCNE_CREATE, SHCNE_MKDIR, SHCNE_DRIVEADD
DebugAppend "HandleShellNotify::code=" & dbg_LookUpSHCNE(lEvent) & ",itm1=" & szItem1 & ",itm2=" & szItem2
        siItem1.GetParent siPar
        If (siPar Is Nothing) = False Then
            If (mShowFiles = False) And (PathIsDirectoryW(StrPtr(szItem1)) = 0&) Then Exit Sub
            siPar.GetDisplayName SIGDN_DESKTOPABSOLUTEPARSING, lpPar
            sPar = LPWSTRtoStr(lpPar)
            hNodePar = GetNodeByPath(sPar)
            hNode = GetNodeByPath(szItem1)
            If (hNode > 0) Then
                lIdx2 = GetIndexFromNode(hNode)
                If TVEntries(lIdx2).bDeleted = False Then Exit Sub 'item already exists and is visible
            End If
            DebugAppend "ADD hNodePar=" & hNodePar
            If hNodePar > 0 Then
                If (TreeView_GetItemState(hTVD, hNodePar, TVIS_EXPANDEDONCE) And TVIS_EXPANDEDONCE) = TVIS_EXPANDEDONCE Then
                    lIdx = GetIndexFromNode(hNodePar)
                    If TVEntries(lIdx).bDeleted = False Then
                        TVAddItem siItem1, hNodePar
                    End If
                Else
                    If (lEvent = SHCNE_MKDIR) Or ((lEvent = SHCNE_CREATE) And (mShowFiles = True)) Then
                        'Not adding it; but if it's a folder or file (when showing them), make sure the dest has children
                        tVI.Mask = TVIF_HANDLE Or TVIF_CHILDREN
                        tVI.hItem = hNodePar
                        TreeView_GetItem hTVD, tVI
                        
                        If tVI.cChildren = 0& Then
                            tVI.Mask = TVIF_CHILDREN
                            tVI.hItem = hNodePar
                            tVI.cChildren = 1&
                            TreeView_SetItem hTVD, tVI
                        End If
                    End If
                End If
            End If
        End If
        
    Case SHCNE_DELETE, SHCNE_RMDIR, SHCNE_DRIVEREMOVED
DebugAppend "HandleShellNotify::code=" & dbg_LookUpSHCNE(lEvent) & ",itm1=" & szItem1 & ",itm2=" & szItem2
        hNode = GetNodeByPath(szItem1)
        TreeView_DeleteItem hTVD, hNode
        'item data is updated by TVN_DELETE notification so all children receive it too
             
    Case SHCNE_RENAMEITEM, SHCNE_RENAMEFOLDER
DebugAppend "HandleShellNotify::code=" & dbg_LookUpSHCNE(lEvent) & ",itm1=" & szItem1 & ",itm2=" & szItem2
        If bRNf Then Exit Sub
        hNode = GetNodeByPath(szItem1)
        If hNode > 0& Then
            Dim lpFull As Long, lpNameFull As Long
            Dim sName As String, sFull As String, sNameFull As String
            Dim nIcon As Long
            siItem2.GetDisplayName SIGDN_NORMALDISPLAY, lpName
            sName = LPWSTRtoStr(lpName)
            siItem2.GetDisplayName SIGDN_DESKTOPABSOLUTEPARSING, lpFull
            sFull = LPWSTRtoStr(lpFull)
            siItem2.GetDisplayName SIGDN_PARENTRELATIVEPARSING, lpNameFull
            sNameFull = LPWSTRtoStr(lpNameFull)
            nIcon = GetFileIconIndexPIDL(pidl2, SHGFI_SMALLICON)
            
            lIdx = GetIndexFromNode(hNode)
            
            With TVEntries(lIdx)
                .sName = sName
                .sNameFull = sNameFull
                .sFullPath = sFull
                If .nIcon <> nIcon Then
                    .nIcon = nIcon
                    tVI.Mask = TVIF_IMAGE
                    tVI.iImage = nIcon
                    tVI.iSelectedImage = nIcon
                End If
            End With
            
            tVI.Mask = tVI.Mask Or TVIF_TEXT Or TVIF_HANDLE
            tVI.hItem = hNode
            tVI.cchTextMax = Len(TVEntries(lIdx).sName)
            tVI.pszText = StrPtr(TVEntries(lIdx).sName)
            
            SendMessage hTVD, TVM_SETITEMW, 0&, tVI
        End If
    Case SHCNE_MEDIAINSERTED
DebugAppend "HandleShellNotify::code=" & dbg_LookUpSHCNE(lEvent) & ",itm1=" & szItem1 & ",itm2=" & szItem2
        lIdx = GetIndexFromPath(szItem1)
        If lIdx >= 0 Then
            If TVEntries(lIdx).bDisabled = True Then
                TreeView_SetItemStateEx hTVD, TVEntries(lIdx).hNode, 0&
                TVEntries(lIdx).bDisabled = False
            End If
        End If
    Case SHCNE_MEDIAREMOVED
DebugAppend "HandleShellNotify::code=" & dbg_LookUpSHCNE(lEvent) & ",itm1=" & szItem1 & ",itm2=" & szItem2
        lIdx = GetIndexFromPath(szItem1)
        If lIdx >= 0 Then
            TreeView_SetItemStateEx hTVD, TVEntries(lIdx).hNode, TVIS_EX_DISABLED
            TVEntries(lIdx).bDisabled = True
        End If
Case SHCNE_UPDATEDIR
    'This shouldn't be raised besides add/remove device
    'Normal drives will get added via SHCNE_DRIVEADD
    'UPDATEDIR seems to be the only message we get if phone/camera type devices are
    'connected. So all we should need to do is scan specifically for those with the
    'Portable Device Manager... (Note: They do get SHCNE_DRIVEREMOVE, so we need only
    'worry about adding new ones with this update message).
    Dim tcc As Long, cdif As Long
    tcc = GetTickCount()
    If lUDTC <> 0 Then
        cdif = tcc - lUDTC
    Else
        cdif = 1001
    End If
    DebugAppend "UD Last=" & lUDTC & ",now=" & tcc & ",dif=" & CStr(cdif)
    lUDTC = tcc
    If cdif > 1000 Then
        ScanDevices
    End If
    
End Select
'<EhFooter>
Exit Sub

e0:
    DebugAppend "ucShellBrowse.HandleShellNotify->Error: " & Err.Description & ", 0x" & Hex$(Err.Number)
'</EhFooter>
End Sub

Private Function ScanDevices() As Long
'Drives are detected via their own message that doesn't apply to devices
'All devices get are a SHCNE_UPDATEDIR message in Computer/This PC
'But we don't want to rescan and potentially spin up normal drives, so
'we'll just scan for portable devices like phones and cameras, because
'that's what we're really looking for.
Dim pMgr As oleexp.PortableDeviceManager
Set pMgr = New oleexp.PortableDeviceManager
Dim nDev As Long, nDevAdded As Long
Dim lDev() As Long
Dim sID() As String
Dim sFN As String
Dim wch() As Integer
Dim cchFN As Long
Dim sParse As String
Dim psi As oleexp.IShellItem
Dim li As Long
Dim kCt As Long

pMgr.GetDevices 0&, nDev
DebugAppend "nDev=" & nDev
If nDev > 0 Then
    ReDim lDev(nDev - 1)
    pMgr.GetDevices VarPtr(lDev(0)), nDev
    ReDim sID(nDev - 1)
    Dim i As Long, j As Long
    For i = 0 To UBound(lDev)
        sID(i) = LPWSTRtoStr(lDev(i))
        DebugAppend "Got PortableDevice(" & i & ")=" & sID(i)
'        pMgr.GetDeviceFriendlyName sID(i), 0&, cchFN
'        ReDim wch(cchFN)
'        pMgr.GetDeviceFriendlyName sID(i), VarPtr(wch(0)), cchFN
'        sFN = WCHARtoSTR(wch)
'        DebugAppend "devid(" & i & ") Name=" & sFN
        sParse = AddBackslash(sComp) & sID(i)
        oleexp.SHCreateItemFromParsingName StrPtr(sParse), Nothing, IID_IShellItem, psi
        If (psi Is Nothing) = False Then
            Dim hItemComp As Long
            If mComputerAsRoot Then
                hItemComp = TreeView_GetRoot(hTVD)
            Else
                hItemComp = GetNodeByPath(sComp)
                
            End If
            If hItemComp Then
                Dim dwSt As ucst_TVITEM_State
                dwSt = TreeView_GetItemState(hTVD, hItemComp, TVIS_EXPANDED Or TVIS_EXPANDEDONCE)
                If dwSt Then
                    
                    Dim lIdxDV As Long
                    lIdxDV = GetIndexFromPath(sParse)
                    If (lIdxDV = -1&) Then
                        TVAddItem psi, hItemComp
                        nDevAdded = nDevAdded + 1
                    Else
                        If TVEntries(lIdxDV).bDeleted Then
                            TVAddItem psi, hItemComp
                            nDevAdded = nDevAdded + 1
                        End If
                    End If
                End If
            End If
        End If
nxt:
    Next i
    ScanDevices = nDevAdded
End If
End Function

Public Sub SetFocusOnTree()
If mTmrProc = 0& Then mTmrProc = scb_SetCallbackAddr(4, 4, , , True)
SetTimer hTVD, WM_USER + 99, 10, mTmrProc
End Sub

Private Sub DoSetFocus()
FastDoEvents 'DoEvents
SetFocus hTVD
End Sub

Private Function pvShiftState() As Integer
  Dim lS As Integer
    If (GetAsyncKeyState(vbKeyShift) < 0) Then lS = lS Or vbShiftMask
    If (GetAsyncKeyState(vbKeyMenu) < 0) Then lS = lS Or vbAltMask
    If (GetAsyncKeyState(vbKeyControl) < 0) Then lS = lS Or vbCtrlMask
    pvShiftState = lS
End Function

' === Call InterfaceMethod ===============================================
' This function was made by ANDRay, wich can be found in http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=72856
Private Function CallInterface(ByVal pInterface As Long, ByVal Member As Long, ByVal ParamsCount As Long, Optional ByVal p1 As Long = 0, Optional ByVal p2 As Long = 0, Optional ByVal p3 As Long = 0, Optional ByVal p4 As Long = 0, Optional ByVal p5 As Long = 0, Optional ByVal p6 As Long = 0, Optional ByVal p7 As Long = 0, Optional ByVal p8 As Long = 0, Optional ByVal p9 As Long = 0, Optional ByVal p10 As Long = 0) As Long
  Dim i As Long, t As Long
  Dim hGlobal As Long, hGlobalOffset As Long
  
  If ParamsCount < 0 Then Err.Raise 5 'invalid call
  If pInterface = 0 Then Err.Raise 5
  
  ' 5 Bytes por parametro (4 bytes + PUSH)
  ' 5 Bytes = 1 push + Puntero a interfaz
  hGlobal = GlobalAlloc(GMEM_FIXED, 5 * ParamsCount + 5 + 5 + 3 + 1)
  If hGlobal = 0 Then Err.Raise 7 'insuff. memory
  hGlobalOffset = hGlobal
  
  If ParamsCount > 0 Then
    t = VarPtr(p1)
    For i = ParamsCount - 1 To 0 Step -1
      Call PutMem2(hGlobalOffset, asmPUSH_imm32)
      hGlobalOffset = hGlobalOffset + 1
      Call GetMem4(t + i * 4, hGlobalOffset)
      hGlobalOffset = hGlobalOffset + 4
    Next
  End If
  
  ' PUSH y ponemos el puntero a la interfas
  Call PutMem2(hGlobalOffset, asmPUSH_imm32)
  hGlobalOffset = hGlobalOffset + 1
  Call PutMem4(hGlobalOffset, pInterface)
  hGlobalOffset = hGlobalOffset + 4
  
  ' Llamamos
  Call PutMem2(hGlobalOffset, asmCALL_rel32)
  hGlobalOffset = hGlobalOffset + 1
  Call GetMem4(pInterface, VarPtr(t))     'дереференс: находим положение vTable
  Call GetMem4(t + Member * 4, VarPtr(t)) 'смещение по vTable, после чего дереференс оного
  Call PutMem4(hGlobalOffset, t - hGlobalOffset - 4)
  hGlobalOffset = hGlobalOffset + 4
  
  Call PutMem4(hGlobalOffset, &H10C2&)        'ret 0x0010
  CallInterface = CallWindowProcA(hGlobal, 0, 0, 0, 0)
  Call GlobalFree(hGlobal)
End Function

'----------------------------------------------------------------------------------------
' IOLEInPlaceActiveObject interface
'----------------------------------------------------------------------------------------
Private Sub pvInitIPAO()
    Dim uiid As oleexp.UUID
    ptrMe = ObjPtr(Me)
    With m_uIPAO
        .lpVTable = GetVTable
        Call IIDFromString(StrPtr(szIID_IOleInPlaceActive), uiid)
        Call CallInterface(ptrMe, ucst_IUnknown_Exports.QueryInterface, 2, VarPtr(uiid), VarPtr(.IPAOReal))
        .ThisPointer = VarPtr(m_uIPAO)
    End With
End Sub

Private Sub pvSetIPAO()
DebugAppend "pvSetIPAO", 3
    Const IOleObject_GetClientSite As Long = 4 ' 2 From IUnknown + 2є Ordinal
    Const IOleObject_DoVerb As Long = 11
    Const IOleInPlaceSite_GetWindowContext As Long = 8 ' 2 from IUnknown + 2 IOleWindow + 4є Ordinal
    Const IOleInPlaceFrame_SetActiveObject As Long = 8 ' 2 from IUnknown + 2 IOleWindow + 4є Ordinal
    Const IOleInPlaceUIWindow_SetActiveObject As Long = 8 ' IOleInPlaceFrame inherits from IOleInPlaceUIWindow
    
    Const OLEIVERB_UIACTIVATE As Long = -4
    Dim uiid As oleexp.UUID, lResult As Long
    Dim pOleObject          As Long 'IOleObject
    Dim pOleInPlaceSite     As Long 'IOleInPlaceSite
    Dim pOleInPlaceFrame    As Long 'IOleInPlaceFrame
    Dim pOleInPlaceUIWindow As Long 'IOleInPlaceUIWindow
    Dim rcPos               As oleexp.RECT
    Dim rcClip              As oleexp.RECT
    Dim uFrameInfo          As OLEINPLACEFRAMEINFO
    
    On Error GoTo e0
    Call IIDFromString(StrPtr(szIID_IOleObject), uiid)
    Call CallInterface(ptrMe, ucst_IUnknown_Exports.QueryInterface, 2, VarPtr(uiid), VarPtr(pOleObject))
    Call CallInterface(pOleObject, IOleObject_GetClientSite, 1, VarPtr(pOleInPlaceSite))
    
    If pOleInPlaceSite <> 0 Then
        Call IIDFromString(StrPtr(szIID_IOleInPlaceSite), uiid)
        Call CallInterface(pOleInPlaceSite, ucst_IUnknown_Exports.QueryInterface, 2, VarPtr(uiid), VarPtr(pOleInPlaceSite))
    DebugAppend "Check1"
        Call CallInterface(pOleInPlaceSite, IOleInPlaceSite_GetWindowContext, 5, VarPtr(pOleInPlaceFrame), VarPtr(pOleInPlaceUIWindow), VarPtr(rcPos), VarPtr(rcClip), VarPtr(uFrameInfo))
        
        DebugAppend "Window=" & pOleInPlaceUIWindow & ",Frame=" & pOleInPlaceFrame & ",rc={" & rcPos.Left & "," & rcPos.Right & "," & rcPos.Top & "," & rcPos.Bottom & "}", 2
        If pOleInPlaceFrame <> 0 Then
            ' The original was pOleInPlaceFrame.SetActiveObject but IOleInPlaceUIWindow has the definition :/
    DebugAppend "Check2"
            Call CallInterface(pOleInPlaceFrame, IOleInPlaceFrame_SetActiveObject, 2, m_uIPAO.ThisPointer, StrPtr(vbNullString))
    DebugAppend "Check3"
        End If
        If pOleInPlaceUIWindow <> 0 Then  '-- And Not m_bMouseActivate
    DebugAppend "Check4"
            Call CallInterface(pOleInPlaceUIWindow, IOleInPlaceUIWindow_SetActiveObject, 2, VarPtr(m_uIPAO.ThisPointer), StrPtr(vbNullString))
        Else
            Call CallInterface(pOleObject, IOleObject_DoVerb, 6, OLEIVERB_UIACTIVATE, 0, pOleInPlaceSite, 0, UserControl.hWnd, VarPtr(rcPos))
        End If
    End If
    DebugAppend "Check5"
    
    On Error GoTo 0
    Exit Sub
e0:
    DebugAppend "pvSetIPAO.Error->" & Err.Description & "," & Err.Number, 4
    Resume Next

End Sub

Private Function pvTranslateAccel(pMsg As Msg) As Boolean
    
    Const IOleObject_GetClientSite As Long = 4 ' 2 From IUnknown + 2є Ordinal
    Dim pOleObject      As Long 'IOleObject
    Dim pOleControlSite As Long 'IOleControlSite
    Dim uiid As oleexp.UUID, hEdit As Long
    
    On Error Resume Next
    Select Case pMsg.message
        Case WM_KEYDOWN, WM_KEYUP
            Select Case pMsg.wParam
                Case vbKeyTab
                    DebugAppend "vbKeyTab"
                    If (pvShiftState() And vbCtrlMask) Then
                        Call IIDFromString(StrPtr(szIID_IOleObject), uiid)
                        Call CallInterface(ptrMe, ucst_IUnknown_Exports.QueryInterface, 2, VarPtr(uiid), VarPtr(pOleObject))
                        Call CallInterface(pOleObject, IOleObject_GetClientSite, 1, VarPtr(pOleControlSite))
                        If pOleControlSite Then
                            Call IIDFromString(StrPtr(szIID_IOleControlSite), uiid)
                            Call CallInterface(pOleControlSite, ucst_IUnknown_Exports.QueryInterface, 2, VarPtr(uiid), VarPtr(pOleControlSite))
                            Call CallInterface(pOleControlSite, 7, 2, VarPtr(pMsg), pvShiftState() And vbShiftMask)
                        End If
                    End If
                    pvTranslateAccel = False
                Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd, vbKeyPageDown, vbKeyPageUp
                    DebugAppend "vbArrowKey"
'                    hEdit = pvEdithWnd()
                    If hEdit Then
                        Call SendMessage(hLEEdit, pMsg.message, pMsg.wParam, ByVal pMsg.lParam)
                    Else
                        Call SendMessage(hTVD, pMsg.message, pMsg.wParam, ByVal pMsg.lParam)
                    End If
                    pvTranslateAccel = True
            End Select
    End Select
    On Error GoTo 0
End Function

Private Function GetVTable() As Long
    ' Set up the vTable for the interface and return a pointer to it
    If (m_IPAOVTable(0) = 0) Then
        m_IPAOVTable(0) = scb_SetCallbackAddr(2, 12, Me) ' QueryInterface
        m_IPAOVTable(1) = scb_SetCallbackAddr(1, 14, Me) ' Addref
        m_IPAOVTable(2) = scb_SetCallbackAddr(1, 13, Me) ' Release
        m_IPAOVTable(3) = scb_SetCallbackAddr(2, 11, Me)  ' GetWindow
        m_IPAOVTable(4) = scb_SetCallbackAddr(2, 10, Me)  ' ContextSensitiveHelp
        m_IPAOVTable(5) = scb_SetCallbackAddr(2, 9, Me)  ' TranslateAccelerator
        m_IPAOVTable(6) = scb_SetCallbackAddr(2, 8, Me)  ' OnFrameWindowActivate
        m_IPAOVTable(7) = scb_SetCallbackAddr(2, 7, Me)  ' OnDocWindowActivate
        m_IPAOVTable(8) = scb_SetCallbackAddr(4, 6, Me)  ' ResizeBorder
        m_IPAOVTable(9) = scb_SetCallbackAddr(2, 5, Me)  ' EnableModeless
        '--- init guid
        With IID_IOleInPlaceActiveObject
            .Data1 = &H117&
            .Data4(0) = &HC0
            .Data4(7) = &H46
        End With
    End If
    GetVTable = VarPtr(m_IPAOVTable(0))
End Function
'=======================================================================================
'SUBCLASSING/IPAO PROCEDURES
'WARNING: Do not add any additional procedures or otherwise alter the order of the following
'         as they're used with self-subclass/self-callback code-- order dependent!
'

Private Function pvIPAO_AddRef(This As IPAOHookStruct) As Long
    pvIPAO_AddRef = CallInterface(This.IPAOReal, ucst_IUnknown_Exports.AddRef, 0)
End Function

Private Function pvIPAO_Release(This As IPAOHookStruct) As Long
    pvIPAO_Release = CallInterface(This.IPAOReal, ucst_IUnknown_Exports.Release, 0)
End Function

Private Function pvIPAO_QueryInterface(This As IPAOHookStruct, riid As oleexp.UUID, pvObj As Long) As Long
    If (IsEqualGUID(riid, IID_IOleInPlaceActiveObject)) Then
        pvObj = VarPtr(This)
        Call pvIPAO_AddRef(This)
        pvIPAO_QueryInterface = 0
      Else
        pvIPAO_QueryInterface = CallInterface(This.IPAOReal, ucst_IUnknown_Exports.QueryInterface, 2, VarPtr(riid), VarPtr(pvObj))
    End If
End Function

Private Function pvIPAO_GetWindow(This As IPAOHookStruct, phwnd As Long) As Long
    pvIPAO_GetWindow = CallInterface(This.IPAOReal, ucst_IPAO_Exports.GetWindow, 1, VarPtr(phwnd))
End Function

Private Function pvIPAO_ContextSensitiveHelp(This As IPAOHookStruct, ByVal fEnterMode As Long) As Long
    pvIPAO_ContextSensitiveHelp = CallInterface(This.IPAOReal, ucst_IPAO_Exports.ContextSensitiveHelp, 1, VarPtr(fEnterMode))
End Function

Private Function pvIPAO_TranslateAccelerator(This As IPAOHookStruct, lpMsg As Msg) As Long
    ' Check if we want to override the handling of this key code:
    If (pvTranslateAccel(lpMsg)) Then
        pvIPAO_TranslateAccelerator = S_OK
    Else
        pvIPAO_TranslateAccelerator = CallInterface(This.IPAOReal, ucst_IPAO_Exports.TranslateAccelerator, 1, VarPtr(lpMsg))
    End If
End Function

Private Function pvIPAO_OnFrameWindowActivate(This As IPAOHookStruct, ByVal fActivate As Long) As Long
    pvIPAO_OnFrameWindowActivate = CallInterface(This.IPAOReal, ucst_IPAO_Exports.OnFrameWindowActivate, 1, VarPtr(fActivate))
End Function

Private Function pvIPAO_OnDocWindowActivate(This As IPAOHookStruct, ByVal fActivate As Long) As Long
    pvIPAO_OnDocWindowActivate = CallInterface(This.IPAOReal, ucst_IPAO_Exports.OnDocWindowActivate, 1, VarPtr(fActivate))
End Function

Private Function pvIPAO_ResizeBorder(This As IPAOHookStruct, prcBorder As oleexp.RECT, ByVal puiWindow As Long, ByVal fFrameWindow As Long) As Long
    pvIPAO_ResizeBorder = CallInterface(This.IPAOReal, ucst_IPAO_Exports.ResizeBorder, 3, VarPtr(prcBorder), puiWindow, VarPtr(fFrameWindow))
End Function

Private Function pvIPAO_EnableModeless(This As IPAOHookStruct, ByVal fEnable As Long) As Long
    pvIPAO_EnableModeless = CallInterface(This.IPAOReal, ucst_IPAO_Exports.EnableModeless, 1, VarPtr(fEnable))
End Function

'@4
Private Sub FocusTimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal TimerID As Long, ByVal Tick As Long)
    KillTimer hWnd, TimerID
    If hWnd = hTVD Then DoSetFocus
End Sub

'@3 - This procedure must be third to last in this module
Private Sub LabelEditWndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, _
                      ByRef lParamUser As Long)
'Label Edit handler
'Label edits are used to rename files, so the in-place edit control is subclassed here
'to check whether an invalid character has been entered, and if so block it

Dim iCheck As Integer
Select Case uMsg
   Case WM_NOTIFYFORMAT
       lReturn = NFR_UNICODE
       bHandled = True
       DebugAppend "NFR on LEWndProc"
       Exit Sub
   Case WM_CHAR
      Select Case wParam
         Case VK_LEFT
             DebugAppend "LeftArrow"
         Case 47, 92, 60, 62, 58, 42, 124, 63, 34 'Illegal chars /\<>:*|?"
             Call ShowBalloonTipEx(lng_hWnd, "", "File names may not contain any of the following characters:" & vbCrLf & " / \ < > : ? * | " & Chr$(34), TTI_NONE) ' TTI_ERROR)
             wParam = 0
             bHandled = True
      End Select
     Case WM_UNICHAR
        If wParam = UNICODE_NOCHAR Then
            lReturn = 1
            bHandled = True
        Else
            SendMessage lng_hWnd, WM_CHAR, wParam, ByVal lParam
            bHandled = True
        End If
    Case WM_IME_CHAR
        SendMessage lng_hWnd, WM_CHAR, wParam, ByVal lParam
        bHandled = True
     
     Case WM_PASTE
         iCheck = IsClipboardValidFileName()
         If iCheck = -1 Then
             Beep
             Call ShowBalloonTipEx(lng_hWnd, "", "File names may not contain any of the following characters:" & vbCrLf & " / \ < > : ? * | " & Chr$(34), TTI_NONE) ' TTI_ERROR)
             bHandled = True
             Exit Sub
         ElseIf iCheck = -2 Then
             Beep
             Call ShowBalloonTipEx(lng_hWnd, "", "The file name you have entered is too long. The total length of the path and file name cannot exceed " & CStr(MAX_PATH) & " characters.", TTI_NONE) ' TTI_ERROR)
             bHandled = True
             Exit Sub

         End If
     Case WM_NOTIFY
         Dim NM As NMHDR
         CopyMemory NM, ByVal lParam, LenB(NM)
         Select Case NM.Code
             Case WM_NOTIFYFORMAT
                lReturn = NFR_UNICODE
                bHandled = True
                DebugAppend "NFR on LEWndProc@WM_NOTIFY"
         End Select
End Select

End Sub

'@2 - This procedure must be second to last in this module
Private Function TVSortProc(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal lParamSort As Long) As Long
    '<EhHeader>
    On Error GoTo e0
    '</EhHeader>
Dim hr As Long
Dim pidlRel1 As Long, pidlRel2 As Long

pidlRel1 = TVEntries(lParam1).pidlRel
pidlRel2 = TVEntries(lParam2).pidlRel

hr = psfCur.CompareIDs(0&, pidlRel1, pidlRel2)

If (hr >= NOERROR) Then
  If (lParamSort And SORT_ASCENDING) = 0 Then
    TVSortProc = LoWord(hr)
  Else
    TVSortProc = LoWord(hr) * -1
  End If
End If
 
'<EhFooter>
Exit Function

e0:
    DebugAppend "ucShellTree.TVSortProc->Error: " & Err.Description & ", 0x" & Hex$(Err.Number)
'</EhFooter>
End Function

'*3
'@1 - This procedure must be the last in this module
Private Sub ucWndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, _
                      ByRef lParamUser As Long)
    '<EhHeader>
    On Error GoTo e0
    '</EhHeader>
Dim PT As oleexp.POINT
Dim tvhti As TVHITTESTINFO
Dim tVI As TVITEM
Dim tvix As TVITEMEX
Dim i As Long
Dim ulAtrrs As Long
Dim lLen As Long, lSel As Long
Dim sText As String
Dim lp As Long

Select Case uMsg
    Case WM_NOTIFYFORMAT
        DebugAppend "Got NFMT on ftv main"
        lReturn = NFR_UNICODE
        bHandled = True
        Exit Sub
    
    
'  Case WM_PAINT
'     If lng_hWnd = hTVD Then
'        Dim hdcInst As Long, hdcBitmap As Long
'        Dim ps As PAINTSTRUCT
'        Dim bp As BITMAP
'        Call GetObject(hBmpBack, LenB(bp), bp)
'        hdcInst = BeginPaint(lng_hWnd, ps)
''        hdcInst = GetDC(lng_hWnd)
'        SetBkMode hdcInst, TRANSPARENT
'
'        '// Create a memory device compatible with the above DC variable
'
'        hdcBitmap = CreateCompatibleDC(hdcInst) '
'
'        '// Select the new bitmap
'        SetBkMode hdcBitmap, TRANSPARENT
'
'        Call SelectObject(hdcBitmap, hBmpBack) '
'
'        '// Get client coordinates for the StretchBlt() function
'
'        Dim R As oleexp.RECT 'r;
'        GetClientRect lng_hWnd, R
'
'        '// troublesome part, in my oppinion
''        StretchBlt hdcInst, 0, 0, r.Right - r.Left, r.Bottom - r.Top, hdcBitmap, 0, 0, bp.BMWidth, bp.BMHeight, MERGECOPY
'        TransparentBlt hdcInst, 0, 0, R.Right - R.Left, R.Bottom - R.Top, hdcBitmap, 0, 0, bp.BMWidth, bp.BMHeight, MERGECOPY
'
'        '// Cleanup
'
'        DeleteDC hdcBitmap
'        EndPaint lng_hWnd, ps
'   End If
'
'    Case WM_ERASEBKGND
'        lReturn = 1
'        bHandled = True
        
    Case WM_SHOWWINDOW
        If Ambient.UserMode Then
            If wParam = 1 Then
                If hTVD Then Attach hTVD
            Else
                Detach
            End If
        End If
    
    Case WM_UNICHAR
        If wParam = UNICODE_NOCHAR Then
            lReturn = 1
            bHandled = True
        Else
            SendMessage lng_hWnd, WM_CHAR, wParam, ByVal lParam
            bHandled = True
        End If
        
    Case WM_IME_CHAR
        SendMessage lng_hWnd, WM_CHAR, wParam, ByVal lParam
        bHandled = True
        
    Case WM_MOUSEACTIVATE
        If lng_hWnd = UserControl.hWnd Then
            'DebugAppend "WM_MA UC.hWnd", 2
            Call pvSetIPAO
            DebugAppend "WM_MOUSEACTIVATE pvSetIPAO Return"
'            bFirstIPAOEvt = True
        End If
        
    Case WM_SETFOCUS '        If lng_hWnd = hTVD Then
'            If bTopLostFocus Then
'            SetFocus hTVD
'            pvSetIPAO
'            bTopLostFocus = False
'        End If
''            Call pvSetIPAO
'        End If
        If lng_hWnd = hTVD Then
          If bHasFocus = False Then
'            pvSetIPAO
            bHasFocus = True
            DebugAppend "WM_SETFOCUS pvSetIPAO Return"
          End If
        End If
        
    Case WM_KILLFOCUS
        bTopLostFocus = True
        If lng_hWnd = hTVD Then
            bHasFocus = False
        End If
        DebugAppend "WM_KILLFOCUS"
    Case WM_SHNOTIFY
        Dim lEvent As Long
        Dim pInfo As Long
        Dim tInfo As oleexp.SHNOTIFYSTRUCT
        Dim hNotifyLock As Long
        hNotifyLock = oleexp.SHChangeNotification_Lock(wParam, lParam, pInfo, lEvent)
        If hNotifyLock Then
            CopyMemory tInfo, ByVal pInfo, LenB(tInfo)
            HandleShellNotify tInfo.dwItem1, tInfo.dwItem2, lEvent
            Call oleexp.SHChangeNotification_Unlock(hNotifyLock)
        End If
        
    Case WM_MENUSELECT
            Dim szt As String
            Dim lw As Long
            On Error Resume Next
            lw = LoWord(wParam)
            If lw = 0 Then
                UpdateStatus ""
            ElseIf lw > 0 Then
                szt = String$(MAX_PATH, 0&)
                If (ICtxMenu2 Is Nothing) = False Then
                    If lw = wIDSel Then
                        szt = "Selects the current file."
                    Else
                        Call ICtxMenu2.GetCommandString(lw - 1, GCS_HELPTEXTW, 0&, StrPtr(szt), Len(szt))
                        szt = TrimNullW(szt)
                    End If
                End If
                UpdateStatus szt
            End If

    Case WM_NOTIFY
        Dim nmtv As NMTREEVIEW
        Dim hTVDW As Long
        Dim nmtvic As NMTVITEMCHANGE
        Dim hItemChild As Long
        Dim hItemPar As Long
        Dim hItemRoot As Long
      
        CopyMemory nmtv, ByVal lParam, Len(nmtv)
        hTVDW = nmtv.hdr.hwndFrom
        
        Select Case nmtv.hdr.Code
            Case TVN_KEYDOWN
                Dim nmkd As NMTVKEYDOWN
                Dim pCTM As oleexp.IContextMenu
                CopyMemory nmkd, ByVal lParam, Len(nmkd)
'                DebugAppend "KeyDown 0x" & Hex$(nmkd.wVKey)
                RaiseEvent TreeKeyDown(nmkd.wVKey)
                i = TreeView_GetSelection(hTVD)
                Select Case nmkd.wVKey
                
                    Case vbKeyF2
                        If i Then
                            SendMessage hTVD, TVM_EDITLABELW, 0&, ByVal i
                        End If
                    Case vbKeyDelete
                        DebugAppend "DelPress i=" & i
                        If i Then
                            lp = GetTVItemlParam(hTVD, i)
                            If (TVEntries(lp).dwAttrib And SFGAO_CANDELETE) = SFGAO_CANDELETE Then
                               If GetAsyncKeyState(vbKeyShift) Then
                                   Dim sL() As String
                                   ReDim sL(0)
                                   sL(0) = TVEntries(lp).sFullPath
                                   Call DeleteFile(sL)
    
                               Else
                                   SetSelection lp, i
                                   siSelected.BindToHandler 0&, BHID_SFUIObject, IID_IContextMenu, pCTM
                                   If (pCTM Is Nothing) = False Then
                                       InvokeVerb pCTM, "delete"
                                       Exit Sub
                                   End If
                               End If
                            End If
                        End If
                    Case vbKeyC
                        If GetAsyncKeyState(VK_CONTROL) Then
                            SetSelection lp, i
                            siSelected.BindToHandler 0&, BHID_SFUIObject, IID_IContextMenu, pCTM
                            If (pCTM Is Nothing) = False Then
                                InvokeVerb pCTM, "copy"
                                Exit Sub
                            End If
                        End If
                    Case vbKeyX
                        If GetAsyncKeyState(VK_CONTROL) Then
                            lp = GetTVItemlParam(hTVD, i)
                            DebugAppend "Cut " & TVEntries(lp).sName
                            SetSelection lp, i
                            siSelected.BindToHandler 0&, BHID_SFUIObject, IID_IContextMenu, pCTM
                            If (pCTM Is Nothing) = False Then
                                InvokeVerb pCTM, "cut"
                                TVCutSelected
                                lReturn = 1
                                bHandled = True
                                Exit Sub
                            End If
                        End If
                    Case vbKeyV
                        If GetAsyncKeyState(VK_CONTROL) Then
                            lp = GetTVItemlParam(hTVD, i)
                            DebugAppend "Paste to " & TVEntries(lp).sName
                            SetSelection lp, i
                            TVDoPaste
                            lReturn = 1
                            bHandled = True
                        End If
                            
                End Select
        
            Case TVN_DELETEITEMW
                If g_fDeleting = False Then
                    Dim hItemDel As Long, hSibDel As Long
                    TVEntries(nmtv.itemOld.lParam).bDeleted = True
                    hItemPar = TreeView_GetParent(hTVD, nmtv.itemOld.hItem)
                    hItemDel = TreeView_GetChild(hTVD, hItemPar)
                    hSibDel = TreeView_GetNextSibling(hTVD, hItemDel)
                    'If the last child has been deleted, clear the expando button
                    If (hItemDel = 0) Or ((hItemDel = nmtv.itemOld.hItem) And (hSibDel = 0&)) Then
                        tVI.Mask = TVIF_CHILDREN Or TVIF_HANDLE
                        tVI.hItem = hItemPar
                        tVI.cChildren = 0
                        TreeView_SetItem hTVD, tVI
                    End If
                End If
            Case TVN_BEGINLABELEDITW
                DebugAppend "ucWndProc@BgLabelEditW"
                
                i = TreeView_GetSelection(hTVD)
                 SetFocus UserControl.ContainerHwnd
                SetFocus hTVD
                Dim ni5 As Long
                ni5 = GetTVItemlParam(hTVD, i)
                ulAtrrs = TVEntries(ni5).dwAttrib
                If (ulAtrrs And SFGAO_CANRENAME) = SFGAO_CANRENAME Then
    
                'deselect the file name extension by default when renaming by label edit, like in explorer.
    
                    hLEEdit = SendMessage(hTVD, TVM_GETEDITCONTROL, 0, ByVal 0&)
                    If hLEEdit Then
                        lLen = SendMessageW(hLEEdit, WM_GETTEXTLENGTH, 0, ByVal 0&)
                        sText = String$(lLen, 0)
                        Call SendMessageW(hLEEdit, WM_GETTEXT, lLen, ByVal StrPtr(sText))
                        dbg_stringbytes sText
                        dbg_stringbytes TVEntries(ni5).sName
                        DebugAppend "LabelEdit:Orig text=" & sText & ",len=" & lLen
                        sOldLEText = sText
                        If (TVEntries(ni5).bFolder = False) Or (TVEntries(ni5).bZip = True) Then
                            'for files (and zip files), deselect the extension by default
                            If InStr(sText, ".") Then
                                lSel = InStrRev(sText, ".") - 1
                                If lSel > 0 Then
                                    DebugAppend "LabelEdit:EM_SETSEL=" & lSel
                                    Call SendMessage(hLEEdit, EM_SETSEL, 0, ByVal lSel)
                                End If
                            End If
                        End If
                        'We also want to detect invalid characters, to do it properly we need to
                        'subclass the edit control- where the WM_CHAR message is sent
                        If ssc_Subclass(hLEEdit, , 3, , , True, True) Then
                            Call ssc_AddMsg(hLEEdit, MSG_BEFORE, ALL_MESSAGES)
                        End If
         
                        Exit Sub
                    End If
    
                Else 'cannot rename
        
                    DebugAppend "TVN_BEGINLABELEDIT: Cancelling, SFGAO_CANRENAME=False"
                    Beep
                    lReturn = 1
                    bHandled = True
        
                End If
            Case TVN_ENDLABELEDITW
                DebugAppend "ucWndProc@EndLabelEditW"
                Call zUnThunk(hLEEdit, SubclassThunk)
                hLEEdit = 0&
                Dim NMTVDI As NMTVDISPINFO
                CopyMemory NMTVDI, ByVal lParam, Len(NMTVDI)
                Dim lSL As Long
                Dim sBuf As String
                Dim sNewText As String
                lLen = 0&
                i = NMTVDI.Item.hItem
                hLEEdit = SendMessage(hTVD, TVM_GETEDITCONTROL, 0, ByVal 0&)
                If hLEEdit Then
                    lLen = SendMessageW(hLEEdit, WM_GETTEXTLENGTH, 0, ByVal 0&)
                    sBuf = String$(lLen, 20)
                    Call SendMessageW(hLEEdit, WM_GETTEXT, lLen + 1, ByVal StrPtr(sBuf))
                    lSL = lstrlenW(ByVal StrPtr(sBuf))
                    DebugAppend "buflen=" & lSL & ",len=" & lLen
                    sBuf = Left$(sBuf, lSL)
                    dbg_stringbytes sBuf
                Else
                    DebugAppend "No edit control @EndLabelEdit"
                End If
                Dim sBuf2 As String

                sBuf2 = LPWSTRtoStr(NMTVDI.Item.pszText, False)
                If (sBuf2 <> "") And (sBuf2 <> CStr(vbNullChar)) And (sBuf <> sOldLEText) Then
                    Dim tOld As TVEntry, tNew As TVEntry
                    Dim idx As Long
                    Dim sOld As String, sNew As String
                    idx = GetTVItemlParam(hTVD, i)
                    tOld = TVEntries(idx)
                    If tOld.sName <> sBuf Then
                        tNew = tOld
                        tNew.sName = sBuf
                        If tOld.sName <> tOld.sNameFull Then 'hidden extension
                            sOld = AddBackslash(tOld.sParentFull) & tOld.sNameFull
                            sNew = tNew.sName
                            tNew.sNameFull = sNew
                            tNew.sFullPath = AddBackslash(tNew.sParentFull) & tNew.sNameFull
                        Else
                            sOld = tOld.sFullPath
                            sNew = tNew.sName
                            tNew.sNameFull = tNew.sName
                            tNew.sFullPath = AddBackslash(tNew.sParentFull) & tNew.sNameFull
                        End If
                        DebugAppend "Name " & sOld & " As " & tNew.sFullPath
                        Dim hrrf As Long
                        bRNf = True
                        hrrf = RenameFile(sOld, sNew)
                        If PathFileExistsW(StrPtr(tNew.sFullPath)) Then
                            DebugAppend "Rename::Notify"
                            'notify the shell that we have renamed the item
    '                            Call SHChangeNotify(SHCNE_RENAMEITEM, SHCNF_PATHW, StrPtr(tOld.sFullPath), StrPtr(tNew.sFullPath))
    '                            bRNf = False
                            TVEntries(idx) = tNew
                            Dim tvi2 As TVITEM
                            tvi2.Mask = TVIF_TEXT
                            tvi2.hItem = i
                            tvi2.cchTextMax = Len(sNew)
                            tvi2.pszText = StrPtr(sNew)
                            SendMessageW hTVD, TVM_SETITEMW, 0&, tvi2
    
                            DebugAppend "Rename::Out"
                            dbg_stringbytes sNew
                            RaiseEvent ItemRename(tOld.sName, tNew.sName, tOld.sFullPath, tNew.sFullPath, tOld.bFolder, NMTVDI.Item.hItem)
                            lReturn = 1
    '                            bHandled = True
                        Else
                            Beep
                            DebugAppend "Rename::Fail, 0x" & Hex$(hrrf)
                            UpdateStatus "Could not rename file."
                        End If
                        bRNf = False
                    End If
                End If

        'Handle check/uncheck logic that should be built in........
            Case TVN_ITEMCHANGINGW
                If (bFilling = False) And (g_fDeleting = False) And (mAutocheck = True) Then
                    Call CopyMemory(nmtvic, ByVal lParam, Len(nmtvic))
                    tvix.Mask = TVIF_STATEEX
                    tvix.hItem = nmtvic.hItem
                    Call SendMessage(hTVDW, TVM_GETITEM, 0, tvix)
                    If (tvix.uStateEx And TVIS_EX_DISABLED) = TVIS_EX_DISABLED Then 'for some reason disabled items can still be checked...
                        DebugAppend "ChangeBlock: Disabled"
                        lReturn = 1
                        bHandled = True
                        Exit Sub
                    End If
                    If (nmtvic.hItem = m_hRoot) And (mRootHasCheckbox = False) Then
                        If (nmtvic.uStateNew And TVIS_STATEIMAGEMASK) <> 0& Then
                            DebugAppend "ChangeBlock: Root check disabled"
                            lReturn = 1
                            bHandled = True
                            Exit Sub
                        End If
                    End If
                End If
            Case TVN_ITEMCHANGEDW
 
                 If (bFilling = False) And (g_fDeleting = False) Then
                    Call CopyMemory(nmtvic, ByVal lParam, Len(nmtvic))
                    If (mAutocheck = True) Then
                        If (nmtvic.uStateNew And TVIS_STATEIMAGEMASK) = (nmtvic.uStateOld And TVIS_STATEIMAGEMASK) Then Exit Sub
                        DebugAppend "itemchanging, mask=" & Hex$((nmtvic.uStateNew And TVIS_STATEIMAGEMASK)) & ",item=" & GetTVItemText(hTVD, nmtvic.hItem)
                        
                        If (nmtvic.uStateNew And TVIS_STATEIMAGEMASK) = &H2000 Then 'item has been checked
                            Dim sPath As String
                            RaiseEvent ItemCheck(TVEntries(nmtvic.lParam).sName, TVEntries(nmtvic.lParam).sFullPath, TVEntries(nmtvic.lParam).bFolder, 1&, nmtvic.hItem)
                            If bSetParents = False Then
                                'set child states
                                If (nmtvic.uStateNew And TVIS_EXPANDEDONCE) = TVIS_EXPANDEDONCE Then
                                    hItemChild = TreeView_GetChild(hTVDW, nmtvic.hItem)
                                    If hItemChild Then
                                        bSetParents = True
                                        TVCheckChildren hItemChild
                                        bSetParents = False
                                    End If
                                End If
                                'set parent states
                                bSetParents = True
                                hItemPar = TreeView_GetParent(hTVD, nmtvic.hItem)
                                If (hItemPar <> m_hRoot) And (hItemPar <> m_hFav) Then
                                    TVSetParentAfterCheck hItemPar
                                End If
                                bSetParents = False
                            End If
                        ElseIf (nmtvic.uStateNew And TVIS_STATEIMAGEMASK) = &H4000 Then
                            If bSetParents = False Then
                                If (nmtvic.uStateNew And TVIS_EXPANDEDONCE) = TVIS_EXPANDEDONCE Then
                                    hItemChild = TreeView_GetChild(hTVDW, nmtvic.hItem)
                                    If hItemChild Then
                                        bSetParents = True
                                        TVExcludeChildren hItemChild
                                        bSetParents = False
                                    End If
                                End If
                           End If
                        ElseIf (nmtvic.uStateNew And TVIS_STATEIMAGEMASK) = &H1000 Then
                            DebugAppend "UncheckEntry " & GetTVItemText(hTVD, nmtvic.hItem)
                            RaiseEvent ItemCheck(TVEntries(nmtvic.lParam).sName, TVEntries(nmtvic.lParam).sFullPath, TVEntries(nmtvic.lParam).bFolder, 0&, nmtvic.hItem)
                            If bSetParents = False Then
                                Dim bSib As Boolean
                                hItemRoot = TreeView_GetRoot(hTVD)
                                DebugAppend "got root"
                                If (nmtvic.uStateNew And TVIS_EXPANDEDONCE) = TVIS_EXPANDEDONCE Then
                                    'uncheck all children
                                    bSetParents = True
                                    hItemChild = TreeView_GetChild(hTVD, nmtvic.hItem)
                                    TVUncheckAllChildren hTVD, hItemChild
                                    bSetParents = False
                                End If
                                bSetParents = True
                                hItemPar = TreeView_GetParent(hTVD, nmtvic.hItem)
                                If (hItemPar <> m_hRoot) And (hItemPar <> m_hFav) Then
                                    TVSetParentAfterUncheck hItemPar
                                End If
                                bSetParents = False
        
                            End If
                        ElseIf (nmtvic.uStateNew And TVIS_STATEIMAGEMASK) = &H3000 Then
                            If bSetParents = False Then
                                'Roll over to excluded or empty; don't want partial user-set
                                If mExCheckboxes Then
                                    SetTVItemStateImage nmtvic.hItem, tvcsExclude
                                Else
                                    SetTVItemStateImage nmtvic.hItem, tvcsEmpty
                                End If
                            End If
                        End If
                    End If
                End If
                If (nmtvic.uStateNew And TVIS_STATEIMAGEMASK) = &H1000 Then
                    TVEntries(nmtvic.lParam).Checked = False
                    TVEntries(nmtvic.lParam).Excluded = False
                ElseIf (nmtvic.uStateNew And TVIS_STATEIMAGEMASK) = &H2000 Then
                    DebugAppend "ItemCheck"
                    TVEntries(nmtvic.lParam).Checked = True
                    TVEntries(nmtvic.lParam).Excluded = False
                ElseIf ((nmtvic.uStateNew And TVIS_STATEIMAGEMASK) = &H3000) And (mAutocheck = False) Then
                    DebugAppend "ItemExclude"
                    TVEntries(nmtvic.lParam).Checked = False
                    TVEntries(nmtvic.lParam).Excluded = True
                ElseIf (nmtvic.uStateNew And TVIS_STATEIMAGEMASK) = &H4000 Then
                    TVEntries(nmtvic.lParam).Checked = False
                    TVEntries(nmtvic.lParam).Excluded = True
                End If
            Case TVN_SELCHANGEDW
                tvix.Mask = TVIF_STATEEX
                tvix.hItem = nmtv.itemNew.hItem
                Call SendMessage(hTVDW, TVM_GETITEM, 0, tvix)
                If (tvix.uStateEx And TVIS_EX_DISABLED) = TVIS_EX_DISABLED Then 'for some reason disabled items can still be checked...
'                    lReturn = 1
'                    bHandled = True
                    Exit Sub
                End If
                If fLoad = 0 Then PlayNavSound
                lp = nmtv.itemNew.lParam ' GetTVItemlParam(hTVD, nmtv.itemNew.)

                DebugAppend "SelectionChange " & TVEntries(lp).sName & ",Raise=" & fLoad & ",lp=" & lp
                gCurSelIdx = lp
                SetSelection lp, nmtv.itemNew.hItem
                If (mExpandOnLabelClick = False) And (bNavigating = False) And (fLoad = 0&) Then
                    bFlagBlockClkExp = True
                End If
                
            Case TVN_ITEMEXPANDINGW
                DebugAppend "TVN_ITEMEXPANDINGW"
                If (fNoExpand > 0) Or (bFlagBlockClkExp = True) Then
                    If fRefreshing = 0 Then
                    
                        DebugAppend "Blocked Expand " & fNoExpand & "," & bFlagBlockClkExp
                        lReturn = 1
                        bHandled = True
                        If bFlagBlockClkExp Then
                            bBlockExec = True
                        End If
                        Exit Sub
                    End If
                End If
                If (nmtv.Action And TVE_EXPAND) Then
                    RaiseEvent ItemExpand(TVEntries(nmtv.itemNew.lParam).sName, TVEntries(nmtv.itemNew.lParam).sFullPath, nmtv.itemNew.hItem)
                ElseIf (nmtv.Action And TVE_COLLAPSE) Then
                    RaiseEvent ItemCollapse(TVEntries(nmtv.itemNew.lParam).sName, TVEntries(nmtv.itemNew.lParam).sFullPath, nmtv.itemNew.hItem)
                End If
                If (nmtv.itemNew.State And TVIS_EXPANDEDONCE) Then Exit Sub
                DebugAppend "TVN_ITEMEXPANDINGW->" & TVEntries(nmtv.itemNew.lParam).sName & "|" & GetTVItemText(hTVD, nmtv.itemNew.hItem)
                Dim tvif As TVITEM
                tvif.Mask = TVIF_PARAM
                tvif.hItem = nmtv.itemNew.hItem
                TreeView_GetItem hTVD, tvif
                
                Call TVExpandFolder(tvif.lParam, nmtv.itemNew.hItem)
                DebugAppend "lParamOrig=" & nmtv.itemNew.lParam & ",resolve=" & tvif.lParam

                ' Setup the callback and sort the items
                Dim tvscb As TVSORTCB
                tvscb.hParent = nmtv.itemNew.hItem
                If m_cbSort = 0& Then m_cbSort = scb_SetCallbackAddr(3, 2)
                tvscb.lpfnCompare = m_cbSort
                tvscb.lParam = 0&
                Call TreeView_SortChildrenCB(hTVDW, tvscb, 0&)
            Case TVN_ITEMEXPANDEDW
                bBlockExec = False
                hItemBlocked = 0
            Case TVN_BEGINDRAGW, TVN_BEGINRDRAGW
                Dim sSel() As String, nSel As Long
                Dim apidl() As Long, cpidl As Long
                Dim hItemSel As Long
                Dim lpSel As Long
                ReDim sSel(0): ReDim apidl(0)
                DebugAppend "TVN_BEGINDRAG"
                Do
                    hItemSel = TreeView_GetNextSelected(hTVD, hItemSel)
                    If hItemSel Then
                        lpSel = GetTVItemlParam(hTVD, hItemSel)
                        ReDim Preserve sSel(nSel)
                        sSel(nSel) = TVEntries(lpSel).sFullPath
                        DebugAppend "Drag " & sSel(nSel)
                        nSel = nSel + 1
                        ReDim Preserve apidl(cpidl)
                        apidl(cpidl) = ILCombine(TVEntries(lpSel).pidlFQPar, TVEntries(lpSel).pidlRel)
                        cpidl = cpidl + 1
                    End If
                Loop While hItemSel
                If nSel > 0& Then
                    Dim psia As oleexp.IShellItemArray
                    Dim iData As oleexp.IDataObject
                    DebugAppend "Dragging " & nSel & " items"
                    oleexp.SHCreateShellItemArrayFromIDLists cpidl, VarPtr(apidl(0)), psia
                    If (psia Is Nothing) = False Then
                        psia.BindToHandler 0&, BHID_DataObject, IID_IDataObject, iData
                        If (iData Is Nothing) = False Then
                            RaiseEvent DragStart(sSel, psia, iData)
                            Dim hr0 As Long, lRetDD As Long
                            hr0 = SHDoDragDrop(0&, ObjPtr(iData), 0&, DROPEFFECT_COPY Or DROPEFFECT_MOVE Or DROPEFFECT_LINK, lRetDD)
                            DebugAppend "Executed drag " & Join(sSel, "|")
                        Else
                            DebugAppend "Drag::Failed to create iData"
                        End If
                    Else
                        DebugAppend "Drag::Failed to create array"
                    End If
                    For i = 0 To UBound(apidl)
                        CoTaskMemFree apidl(i)
                    Next i
                End If
            
''''                    DebugAppend "TVN_BEGINDRAGW"
''''             ' Get the highlighted treeview item (instead of just TVHT_ONITEM)
''''                Call GetCursorPos(PT)
''''                Call ScreenToClient(hTVDW, PT)
''''                tvhti.PT = PT
''''                Call TreeView_HitTest(hTVDW, tvhti)
''''                If (tvhti.Flags And TVHT_ONITEMLINE) Then
''''
''''                    tVI.Mask = TVIF_PARAM
''''                    tVI.hItem = tvhti.hItem
''''                    If TreeView_GetItem(hTVDW, tVI) Then
''''                        Dim hr0 As Long
''''                        Dim iData As oleexp.IDataObject
''''                        Dim apidl() As Long
''''                        Dim cpidl As Long
''''                        Dim pidlDesk As Long
''''                        Dim lRetDD As Long
''''                        Dim AllowedEffects As DROPEFFECTS
''''                        Dim psi As oleexp.IShellItem
''''                        Dim ppidlp As Long
''''                        oleexp.SHCreateItemFromParsingName StrPtr(TVEntries(tVI.lParam).sFullPath), Nothing, IID_IShellItem, psi
''''                        If (psi Is Nothing) Then
''''                            DebugAppend "ucWndProc::BeginDrag->ShellItem failed, trying to load from pidls...", 3
''''                            If (TVEntries(tVI.lParam).pidlFQPar <> 0&) And (TVEntries(tVI.lParam).pidlRel <> 0&) Then
''''                                ppidlp = ILCombine(TVEntries(tVI.lParam).pidlFQPar, TVEntries(tVI.lParam).pidlRel)
''''                                oleexp.SHCreateItemFromIDList ppidlp, IID_IShellItem, psi
''''                                If (psi Is Nothing) = False Then
''''                                    DebugAppend "ucWndProc::LoadFromPidl->Success!", 3
''''                                Else
''''                                    DebugAppend "ucWndProc::LoadFromPidl->Failed.", 3
''''                                End If
''''                            Else
''''                                DebugAppend "ucWndProc::LoadFromPidl->Parent or child pidl not set.", 3
''''                            End If
''''                            If ppidlp Then CoTaskMemFree ppidlp
''''                        End If
''''
''''                        If (psi Is Nothing) = False Then
''''                            psi.BindToHandler 0&, BHID_DataObject, IID_IDataObject, iData
''''                            If (iData Is Nothing) = False Then
''''                                DebugAppend "tvdrag->valid ido"
'''''                                If DataObjSupportsFormat(iData, CF_PREFERREDDROPEFFECT) Then
''''
'''''                                    If GetPreferredEffect(iData) = DROPEFFECT_LINK Then
'''''                                        AllowedEffects = DROPEFFECT_LINK
'''''                                    Else
''''                                        AllowedEffects = DROPEFFECT_COPY Or DROPEFFECT_MOVE Or DROPEFFECT_LINK
'''''                                    End If
'''''                                End If
''''
''''                                hr0 = SHDoDragDrop(0&, ObjPtr(iData), 0&, AllowedEffects, lRetDD)
''''                                DebugAppend "tvdrag->hr=" & Hex$(hr0)
''''                            Else
''''                                DebugAppend "tvdrag->no ido"
''''                            End If
''''                        Else
''''                            DebugAppend "tvdrag->no psi"
''''                        End If
''''                    End If
''''                End If
            Case TVN_GETINFOTIPW
    '                DebugAppend "TVN_GETINFOTIPW"
                Dim nmtvgit As NMTVGETINFOTIP
                CopyMemory nmtvgit, ByVal lParam, LenB(nmtvgit)
                With nmtvgit
                If .hItem > 0 And .pszText <> 0 Then
                    Dim sInfoTip As String
                    Dim nTipItem As Long
                    nTipItem = GetTVItemlParam(hTVD, .hItem)
'                        DebugAppend "QueryInfoTip " & TVEntries(nTipItem).sName & ",bFolder=" & TVEntries(nTipItem).bFolder
                    If (TVEntries(nTipItem).sInfoTip <> "") Then
                        sInfoTip = TVEntries(nTipItem).sInfoTip
                    Else
                        If (TVEntries(nTipItem).bFolder = True) And (mInfoTipOnFolders = True) Then
                            sInfoTip = GenerateInfoTip(nTipItem)
                            TVEntries(nTipItem).sInfoTip = sInfoTip
                        End If
                        If (TVEntries(nTipItem).bFolder = False) And (mInfoTipOnFiles = True) Then
                            sInfoTip = GenerateInfoTip(nTipItem)
                            TVEntries(nTipItem).sInfoTip = sInfoTip
                        End If
                    End If
                    If Not sInfoTip = vbNullString Then
                        sInfoTip = Left$(sInfoTip, .cchTextMax - 1) & vbNullChar
                        CopyMemory ByVal .pszText, ByVal StrPtr(sInfoTip), LenB(sInfoTip)
                    Else
                        CopyMemory ByVal .pszText, 0&, 4&
                    End If

                End If
                End With
            Case NM_DBLCLK
                If nmtv.hdr.hwndFrom = hTVD Then
                    DebugAppend "DoubleClick"
                    If mExpandOnLabelClick = False Then
                        If bBlockExec Then
                            If hItemBlocked Then
                            bFlagBlockClkExp = False
                            TreeView_Expand hTVD, hItemBlocked, TVE_TOGGLE
                            bBlockExec = False
                            End If
                        End If
                    End If
                    
                End If
            Case NM_CLICK
                If nmtv.hdr.hwndFrom = hTVD Then
'                    SetFocus UserControl.ContainerHwnd
'                    SetFocus hTVD
                    Call GetCursorPos(PT)
                    Call ScreenToClient(hTVDW, PT)
                    tvhti.PT = PT
                    Call TreeView_HitTest(hTVDW, tvhti)
                    bFlagBlockClkExp = False
                    If (tvhti.Flags And TVHT_ONITEMLINE) Then
                        tVI.Mask = TVIF_PARAM Or TVIF_CHILDREN
                        tVI.hItem = tvhti.hItem
                        nItr = tvhti.hItem
                        If TreeView_GetItem(hTVDW, tVI) Then
                            DebugAppend "RaiseEvent(Click)->" & TVEntries(tVI.lParam).sFullPath
                            If mMultiSel Then
                                SetMultiSel
                            End If
                            RaiseEvent ItemClick(TVEntries(tVI.lParam).sName, TVEntries(tVI.lParam).sFullPath, TVEntries(tVI.lParam).bFolder, MK_LBUTTON, tvhti.hItem)
                            RaiseEvent ItemClickByShellItem(siSelected, TVEntries(tVI.lParam).sName, TVEntries(tVI.lParam).sFullPath, TVEntries(tVI.lParam).bFolder, MK_LBUTTON, tvhti.hItem)
                            If mExpandOnLabelClick = False Then
                                If ((tvhti.Flags And TVHT_ONITEMBUTTON) = 0&) And (tVI.cChildren > 0) Then
                                    If (bNavigating = False) Then
                                        bFlagBlockClkExp = True
                                    End If
                                    hItemBlocked = tVI.hItem
                                    DebugAppend "SetBlockFlag"
                                End If
                            End If
                            DoSetFocus
                            pvSetIPAO
                        End If
                    End If
                End If
    
            Case NM_RCLICK   ' lParam = lp NMHDR
                If nmtv.hdr.hwndFrom = hTVD Then
            
                  ' Get the highlighted treeview item (instead of just TVHT_ONITEM)
                  Call GetCursorPos(PT)
                  Call ScreenToClient(hTVDW, PT)
                  tvhti.PT = PT
                  Call TreeView_HitTest(hTVDW, tvhti)
                  If (tvhti.Flags And TVHT_ONITEMLINE) Then
            
                    ' Get the item's ITEMDATA struct
                    tVI.Mask = TVIF_PARAM
                    tVI.hItem = tvhti.hItem
                    If TreeView_GetItem(hTVDW, tVI) Then
                        TreeView_SelectItem hTVD, tVI.hItem
                        RaiseEvent ItemClick(TVEntries(tVI.lParam).sName, TVEntries(tVI.lParam).sFullPath, TVEntries(tVI.lParam).bFolder, MK_RBUTTON, tvhti.hItem)
                        RaiseEvent ItemClickByShellItem(siSelected, TVEntries(tVI.lParam).sName, TVEntries(tVI.lParam).sFullPath, TVEntries(tVI.lParam).bFolder, MK_RBUTTON, tvhti.hItem)

                         ShowShellContextMenu tVI.lParam
                          DebugAppend "rclick " & GetTVItemText(hTVDW, tvhti.hItem)
                          lReturn = 1
                          bHandled = True

                     End If
            
                  End If   ' (tvhti.flags And TVHT_ONITEMLINE)
                End If
            Case NM_CUSTOMDRAW
                'As in Explorer, show encrypted filenames in green, compressed in blue
                If nmtv.hdr.hwndFrom = hTVD Then
                    Dim NMTVCD As NMTVCUSTOMDRAW
                    CopyMemory NMTVCD, ByVal lParam, Len(NMTVCD)
                    With NMTVCD.NMCD
                        Select Case .dwDrawStage
                            Case CDDS_PREPAINT
                                lReturn = CDRF_NOTIFYITEMDRAW
                                bHandled = True
                            
                            Case CDDS_ITEMPREPAINT
                                NMTVCD.ClrText = TVItemGetDispColor(.lItemlParam)
                                CopyMemory ByVal lParam, NMTVCD, Len(NMTVCD)
    
                        End Select
                    End With
                End If

        End Select
 
    Case WM_INITMENUPOPUP, WM_DRAWITEM, WM_MEASUREITEM
      On Error Resume Next
      Dim lResult As Long
      If (ICtxMenu3 Is Nothing) = False Then
        Call ICtxMenu3.HandleMenuMsg2(uMsg, wParam, lParam, lResult)
      End If
      If (ICtxMenu2 Is Nothing) = False Then
        Call ICtxMenu2.HandleMenuMsg(uMsg, wParam, lParam)
      End If
End Select

'<EhFooter>
Exit Sub

e0:
    DebugAppend "ucShellTree.ucWndProc->Error: " & Err.Description & ", 0x" & Hex$(Err.Number)
    Resume Next
'</EhFooter>
End Sub 'ucWndProc
'========================================================================
'WARNING: DO NOT ADD ANY CODE AFTER ucWndProc
