# ShellControls
## Shell Browser and Shell Tree Controls
## ucShellBrowse v11.3 R1 and ucShellTree v2.18 R1

This repository contains a twinBASIC x86/x64 compatible port of my shell controls, ucShellBrowse and ucShellTree.

See the original threads for a full list of all of the extensive features of these controls:

[[VB6] ucShellBrowse: A modern replacement for Drive/FileList w/ extensive features](https://www.vbforums.com/showthread.php?854147-VB6-ucShellBrowse-A-modern-replacement-for-Drive-FileList-w-extensive-features)

[[VB6] ucShellTree - Full-featured Shell Tree UserControl](https://www.vbforums.com/showthread.php?862137-VB6-ucShellTree-Full-featured-Shell-Tree-UserControl)

### These are beta versions. 

There's still some work to do in twinBASIC to complete user control support, but it's now far enough along even a massively complex control like ucShellBrowse can run. Notably, there's numerous issues with sizing and scaling. The demos included in this repository have worked around them as best I could. Also, I haven't exhaustively tested all features. Please don't hestitate to create an issue for any bugs you encounter.

### Using these controls in your project

**Requires [twinBASIC Beta 236](https://github.com/twinbasic/twinbasic/releases) or newer**

The demos are all set to open and run, to set these up in your project:

These projects use tbShellLib, the x64-compatible successor to oleexp.tlb written in twinBASIC. First add a reference to 'twinBASIC Shell Library' in Settings->COM Type Library / Active-X References by clicking TWINPACK PACKAGES and selecting it from the list, or manually downloading it from it's [repository](https://github.com/fafalone/tbShellLib).

ShellControls.twinpack contains both controls as a tB Package and can be added via the same references location (Import the file, it's not on the package server yet). This reference must come before WinNativeForms; tbShellLib must still be added. Note that packages are read-only when added to a project.

As an alternative, to have them in an editable form, for ucShellTree, you need to import ucShellTree.twin and ucShellTree.tbcontrol. For ucShellBrowse, import ucShellBrowse.twin and ucShellBrowse.tbcontrol.

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

## Updates

(Jan 20 2023) All code updated to R1: IPAO hooking code has been integrated into ucShellBrowse and ucShellTree, rather than as a separate module. As a separate module it couldn't be in both, there would be a conflict or it would be missing. But twinBASIC's support for AddressOf on class members makes having it as a standard module unneccessary. ucShellBrowse now loads the navigation sound even when not set by Explorer, the default in Win10, with a user option to disable this behavior if not desired. Additionally, the property descriptions from the VB version have been added to ucShellTree.

----

I'll add more in the future. Note that while due to the difference in typelib and IPAO-hooking methods, you can't directly import VB projects with these controls, you should be able to recreate any of the other screen shots.
