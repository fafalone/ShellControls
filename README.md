# ShellControls
## Shell Browser and Shell Tree Controls
## ucShellBrowse v12.0 R1 and ucShellTree v2.9.1
### Updated 03 Feb 2023 / 27 Jan 2024

This repository contains a twinBASIC x86/x64 compatible port of my shell controls, ucShellBrowse and ucShellTree.

See the original threads for a full list of all of the extensive features of these controls:

[[VB6] ucShellBrowse: A modern replacement for Drive/FileList w/ extensive features](https://www.vbforums.com/showthread.php?854147-VB6-ucShellBrowse-A-modern-replacement-for-Drive-FileList-w-extensive-features)

[[VB6] ucShellTree - Full-featured Shell Tree UserControl](https://www.vbforums.com/showthread.php?862137-VB6-ucShellTree-Full-featured-Shell-Tree-UserControl)

### These are beta versions. 

There's still some work to do in twinBASIC to complete user control support, but it's now far enough along even a massively complex control like ucShellBrowse can run. Notably, there's numerous issues with sizing and scaling. The demos included in this repository have worked around them as best I could. Also, I haven't exhaustively tested all features. Please don't hestitate to create an issue for any bugs you encounter.

### Using these controls in your project

**Requires [twinBASIC Beta 432](https://github.com/twinbasic/twinbasic/releases) or newer**

>[!IMPORTANT]
>Now requires twinBASIC 432 or newer; it was broken in 424-431. You can still use it in 239-423 as well.

The demos are all set to open and run, to set these up in your project:

These projects use tbShellLib, the x64-compatible successor to oleexp.tlb written in twinBASIC. First add a reference to 'twinBASIC Shell Library' in Settings->COM Type Library / Active-X References by clicking TWINPACK PACKAGES and selecting it from the list, or manually downloading it from it's [repository](https://github.com/fafalone/tbShellLib).

ShellControls.twinpack contains both controls as a tB Package and can be added via the same references location (Import the file, it's not on the package server yet). This reference must come before WinNativeForms; tbShellLib must still be added. Note that packages are read-only when added to a project. As of twinBASIC Beta 239, you can now experiment with building this as an Active-X DLL. You can then use the controls in VB6; however they're not working in VBA yet.

As an alternative, to have them in an editable form, for ucShellTree, you need to import ucShellTree.twin and ucShellTree.tbcontrol. For ucShellBrowse, import ucShellBrowse.twin and ucShellBrowse.tbcontrol.

### Update

**IMPORTANT:** These controls now require [twinBASIC Beta 432](https://github.com/twinbasic/twinbasic/releases). The changes in/for this release allow building them as Active-X DLLs that can be used in VB6; however they're not working in VBA yet. They also require WinDevLib 7.0 or newer, if you're adding them to your own project.

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
