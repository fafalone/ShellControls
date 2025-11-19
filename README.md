# ShellControls
## Shell Browser and Shell Tree Controls
## ucShellBrowse v12.2 and ucShellTree v2.9.3  **BETA**
### Updated 12 May 2024

### CRITICAL BUG FIX RELEASED - ucShellBrowse v12.2

This repository contains a twinBASIC x86/x64 compatible port of my shell controls, ucShellBrowse and ucShellTree.

See the original threads for a full list of all of the extensive features of these controls:

[[VB6] ucShellBrowse: A modern replacement for Drive/FileList w/ extensive features](https://www.vbforums.com/showthread.php?854147-VB6-ucShellBrowse-A-modern-replacement-for-Drive-FileList-w-extensive-features)

[[VB6] ucShellTree - Full-featured Shell Tree UserControl](https://www.vbforums.com/showthread.php?862137-VB6-ucShellTree-Full-featured-Shell-Tree-UserControl)

### These are beta versions. 

There's still some work to do in twinBASIC to complete user control support, but it's now far enough along even a massively complex control like ucShellBrowse can run. Notably, there's numerous issues with sizing and scaling. The demos included in this repository have worked around them as best I could. Also, I haven't exhaustively tested all features. Please don't hestitate to create an issue for any bugs you encounter.

### Using these controls in your project

**Requires [twinBASIC Beta 513](https://github.com/twinbasic/twinbasic/releases) or newer**

>[!IMPORTANT]
>Now requires twinBASIC 513 or newer.

The demos are all set to open and run, to set these up in your project:

These projects use WinDevLib, the x64-compatible successor to oleexp.tlb written in twinBASIC. First add a reference to 'Windows Development Library for twinBASIC' in Settings->COM Type Library / Active-X References by clicking TWINPACK PACKAGES and selecting it from the list, or manually downloading it from it's [repository](https://github.com/fafalone/WinDevLib).

----

Files:\
ShellControlsTB.twinproj - Project file for building OCX controls.

ShellControlsPackage.twinpack - Contains both controls as a tB Package and can be added via the same references location (Import the file, it's not on the package server yet). This reference must come before WinNativeForms; tbShellLib must still be added. Note that packages are read-only when added to a project.

tbShellBrowseDemo.twinproj - Main ucShellBrowse demo project.

tbShellTree.twinproj - Main ucShellTree demo project.

FileDialogDemo.twinproj - Demo of combining ucShellBrowse and ucShellTree controls to make a highly customizable Open File dialog. 

UCSBDemoVB.twinproj - Demonstrates use of ucShellBrowse configured as replacements for the built in VB/tB file controls (DriveListBox, DirListBox, and FileListBox). 

As an alternative, to have them in an editable form, for ucShellTree, you need to import ucShellTree.twin and ucShellTree.tbcontrol. For ucShellBrowse, import ucShellBrowse.twin and ucShellBrowse.tbcontrol.

### Update

**19 Nov 2025** All files updated for compatibility with twinBASIC 896+, to use the latest WinDevLib, and to fix a major memory leak.

**IMPORTANT:** These controls now require [twinBASIC Beta 896](https://github.com/twinbasic/twinbasic/releases). 

**Update highlights:**\
-Thumbnail view in ucShellBrowse is now fixed
-OCX builds now working in all hosts, tested in tB, VB6, Excel/Access 2021 x64. 
-New ucShellTree features

**ucShellBrowse v12.2.0 BETA Update - CRITICAL BUG FIXES**

```

'New in v12.2 BETA (Released 12 May 2024)
'
'-(Bug fix) Workaround for Details or List view to Small Icon view freeze issue.
'
'-(Bug fix) PathCompactPath Unicode mismatch caused crashing on Windows 11.
'
'-(Bug fix) m_GroupSubset and m_GroupSubsetLinkText weren't initialized.
```


**ucShellBrowse v12.1.0 BETA Update**

```
'New in v12.1 BETA (Released 18 April 2024)
'
'-Made LVS_EX_DOUBLEBUFFER an optional style (ListViewDoubleBuffer, default True)
'
'-Updated IWebBrowser[App] implementation to use Boolean for VARIANT_BOOL instead
' of Integer (along with updating tbShellLib version which made this change).
'
'-Pointer math operations are now large-address safe, allowing safe use in apps
' or as an OCX with LARGEADDRESSAWARE enabled.
'
'-Errors loading a single file are no longer fatal to the entire folder; also
' generally improved error handling in load folder routines.
'
'-(Bug fix) Thumbnail view broken in all modes except 32bit compiled.
```

**ucShellTree v2.9.2 BETA Update**

```
'v2.9.2 
'
'Special thanks to VBForums user Mith for his incredible work helping with new features
'and bug fixes! Credit goes to him for most of the changes in 2.9.2 alpha and final.
'
'-Added ShowOnlyFileCheckboxes option: When Checkboxes = True and ShowFiles = True, will 
' show checkboxes only for files.
'
'-Try to load cached icons first
'
'-Minor sizing adjustment to eliminate 4px border.
'
'-(Bug fix) Default share/shortcut overlays not showing.
'
'-(Bug fix) New icon methods didn't properly support running without ComCtl6.
'
'-(Bug fix) Always Show extensions wasn't working.
'
'-(Bug fix) Disabling of Wow64 redirection wasn't reverted on terminate.
'
'-(Bug fix) If DPI awareness was on, IconSize would continuously grow.
'
'v2.9.2 (ALPHA, 29 Jan 2024)
'-New FileExtensions property to choose between following Explorer's setting for hiding
' known extensions, over forcing them to always show.
'
'-Added property CheckedPathCount.
'
'-Added UserOption mNeverExpandZip, to prevent expansion of .zip when ShowFiles = True.
'
'-(Bug fix) Border property not restored correctly due to re-reading it with the wrong
'           name from the property bag.
'
'-(Bug fix) Opening to a custom root would add the root as a child of itself.
```

**ucShellTree v2.9.1 Update**
```
-(Bug fix) AutoCheck = False not respected when expanding a checked folder (thanks to 
            VBForums user Mith for report and fix) 
 -(Bug fix) When AutoCheck = False and CheckBoxes = True is set at runtime, checkboxes 
            improperly cycled through partial checks. This is currently fixed as a work- 
            around that will not work properly if ExclusionChecks = True (and AutoCheck 
            = False), but it may be some time before I Can run down a proper fix.
```

**ucShellTree v2.9 Update**\
![image](https://github.com/fafalone/ShellControls/assets/7834493/24301b7e-ea8d-4ab6-83a1-09f70b964288)

I have not updated the joint projects yet-- the controls package or the FileDialogDemo. **Only** tbShellTree.twinproj and ucShellTree.twin/tbcontrol have been updated to 2.9. I'll look at the others next week after the new tB beta is out.\
ucShellTree v2.9 now supports icons of any size, instead of the fairly small fixed sized ones.


R1 of ucShellBrowse and R2 of ucShellTree update tbShellLib to 2.6.62 to correct improperly defined hex literals, and correct a few local ones as well.

**ucShellBrowse v12.0 Changelog**

```
-There's now a RegisterWindow option that, if True, will register this control as
 a shell window. It will be seen by the system as if it were an Explorer window,
 and can be interacted with via IShellView, IShellBrowser, IFolderView, etc, which
 includes responding to APIs such as SHOpenFolderAndSelectItems. 
 For instance, if you use your web browser to download to C:\download, and this
 control is open to C:\download, and you click the 'show in folder' button, it
 will select the file in this control rather than open an actual Explorer window.

 Programs like my 'List open windows and properties' demo will also be able to
 obtain all that information from this control. See the following:
 https://www.vbforums.com/showthread.php?818959
 All members that are applicable to this control have been implemented, with the
 exception I don't intend to implement the FolderItem, Folder, etc interfaces
 for use with IShellFolderViewDual/shell.application.

 Some features must be implemented by the host form/application. For instance,
 in order to respond to requests to close the window, see e.g.
 https://www.vbforums.com/showthread.php?t=898235 the control will raise the
 new RequestExit event, letting your app know the window should close.

 This option is False by default. Not available in DrivesOnly mode, DirOnly
 mode, or DirOnlyWithCtls mode.
 NOTE: If you have multiple ShellBrowse controls showing the same path, it's
       strongly advised only one registers as a shell window.

-Added LinkshellTree sub to store hwnd for associated ucShellTree; for 
 IShellBrowser.GetControlWindow only for now, but may do more in the future.

-Added DropdownShowFullPath option to show the full path in the path dropdown.

-Added ListViewAlphaShadow option, an undocumented feature that applies an alpha
 shadow to ListView labels when in Icon modes (Medium Icon, Large Icon, XL Icon,
 Thumbnail, and Custom, in this control).

-Added NoScroll (LVS_NOSCROLL) style option.

-(Bug fix) Column header text could sometimes become corrupted due to early
           release of temporary strings.

-(Bug fix) This control is labeled as supporting Windows Vista, but some
           ListView features that were available on Vista didn't work as the
           control only used the Windows 7+ version of IListView. This has been
           corrected and all features supported on Vista are available.
```


### The Demos

Here's the included demos in this repo:

#### Basic Shell Tree Demo - tbShellTree.twinproj:

![image](https://user-images.githubusercontent.com/7834493/208004027-283c2d98-aee1-4da8-8fd2-ffebd676414e.png)

#### Main Shell Browser Demo - DemoMain.twinproj

![image](https://user-images.githubusercontent.com/7834493/213609557-64e74258-66f1-41c3-806a-8e1126d21546.png)

#### ucShellBrowse as an upgrade to native VB/TB controls - UCSBDemoVB.twinproj

![image](https://user-images.githubusercontent.com/7834493/213373444-cfdd0e7d-74cc-48c6-95dc-63dd8beb4f25.png)

#### Fully customized file open dialog - FileDialogDemo.twinproj

Shows combining ucShellBrowse and ucShellTree:

![image](https://user-images.githubusercontent.com/7834493/213373633-e539fc13-0287-496e-9d69-a3518a3d6327.png)

## Older Updates

(Jan 22 2023) ucShellBrowse R2: Previously described error handling for categorizer errors had mistkanely been removed.

(Jan 20 2023) All code updated to R1: IPAO hooking code has been integrated into ucShellBrowse and ucShellTree, rather than as a separate module. As a separate module it couldn't be in both, there would be a conflict or it would be missing. But twinBASIC's support for AddressOf on class members makes having it as a standard module unneccessary. ucShellBrowse now loads the navigation sound even when not set by Explorer, the default in Win10, with a user option to disable this behavior if not desired. Additionally, the property descriptions from the VB version have been added to ucShellTree.

----

I'll add more in the future. Note that while due to the difference in typelib and IPAO-hooking methods, you can't directly import VB projects with these controls, you should be able to recreate any of the other screen shots.

![image](https://github.com/user-attachments/assets/6d0995e6-0759-4e19-9251-cbb7ee4e2e9c)

![image](https://github.com/user-attachments/assets/f2376a89-15e1-4676-9be0-48ef7daaba64)

You can open special locations like 'Devices and Printers' and 'Programs and Features', and use the right-click menu to carry out their custom actions like Uninstall a program, or Set As Default for a printer.

![image](https://github.com/user-attachments/assets/3f9326b3-95ac-4fe3-9994-9eac78aa2a20)

The search feature. You can start a quick search by just typing a string in the box on the form and pressing enter, or double-click the box to bring up the popup with additional options. You can optionally have the control box disabled, but have a menu item on right-click that brings up the popup (with additional box to type text in).

![image](https://github.com/user-attachments/assets/15491d4f-5907-4678-b15d-f06cd2bf0780)

You can configure it to be a simple thumbnail list, which includes thumbnails of embedded mp3 album art as shown on the right. Note that Windows doesn't come with a built-in FLAC handler, but if you install a 32-bit one, the control will be able to read/write those too.

![image](https://github.com/user-attachments/assets/bc6b8708-7b38-4822-a8ae-170b91c4447f)
![image](https://github.com/user-attachments/assets/ba075ca6-a4b5-48ed-92c4-c7f4849999bb)

Any property that's able to be edited can be set through the control, including through a popup date/time control or dropdown list where needed.


