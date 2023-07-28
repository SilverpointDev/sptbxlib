# SpTBXLib

SpTBXLib is an add on package for TB2K components, it adds the following features:
- Unicode support
- Custom Skins engine
- Delphi VCL Styles support
- High DPI support
- Custom painting events
- Toolbar Customizer with Drag & Drop
- Custom item size
- Anchored items
- Right aligned items
- Accel char handling
- Extended button controls: Button, Label, Checkbox, RadioButton
- Extended editors: Edit, ButtonEdit, SpinEdit, ColorEdit, ComboBox, FontComboBox, ListBox, CheckListBox, ColorListBox
- Extended ProgressBar and TrackBar controls
- Panel and GroupBox with transparency support
- TabSet and TabControl with TB2K items support
- DockablePanel with TB2K items support
- Titlebar with TB2K items support
- StatusBar with TB2K items support
- Form Popup components
- Skin Editor

For more info go to:
http://www.silverpointdevelopment.com

# License

Use and/or distribution of the files requires compliance with the
SpTBXLib License, found in SpTBXLib-LICENSE.txt or at:

  http://www.silverpointdevelopment.com/sptbxlib/SpTBXLib-LICENSE.htm

Alternatively, at your option, the files may be used and/or distributed under
the terms of the Mozilla Public License Version 1.1, found in MPL-LICENSE.txt or at:

  http://www.mozilla.org/MPL

# Installation

### Requirements:
- RAD Studio XE2 or newer:
- Jordan Russell's Toolbar 2000
  http://www.jrsoftware.org

### Installing with Silverpoint MultiInstaller (http://www.silverpointdevelopment.com/multiinstaller/index.htm):
- Create a new folder for the installation.
- Download all the component zips to a folder: SpTBXLib + TB2K
- Download Silverpoint MultiInstaller and the Setup.Ini, extract them to the folder:

The installation folder will end up with this files: 
C:\MyInstall
       |-  SpTBXLib.zip
       |-  tb2k-2.2.2.zip
       |-  MultiInstaller.exe
       |-  Setup.ini 

You are ready to install the component packages, just run MultiInstaller, select the destination folder, and all the components will be unziped, patched, compiled and installed on the Delphi IDE.

### To install SpTBXLib manually:

First you need to apply the TB2K patch:
- Extract TB2K to a folder
- Extract SpTBXLib to a folder
- Copy the contents of SpTBXLib\TB2K Patch folder to TB2K\Source folder
- Run tb2kpatch.bat

Add the Source directories to the Library Path
- Add 'TB2K\Source' directory to Tools->Options->Language->Delphi->Library->Library Path
- Add 'SpTBXLib\Source' directory to Tools->Options->Language->Delphi->Library->Library Path

Compile and install the components
- If you have a previous version of TB2K installed in the IDE remove it from Component->Install Packages, select TB2K from the list and press the Remove button.
- If you have a previous version of SpTBXLib installed in the IDE remove it from Component->Install Packages, select SpTBXLib from the list and press the Remove button.
- Open the TB2K design package corresponding to the IDE version (tb2kdsgn_d12), press Compile and then press Install, close the package window (don't save the changes).
- Open the SpTBXLib design package corresponding to the IDE version (SpTBXLibDsgn_*.dpk), press Compile and then press Install, close the package window (don't save the changes).

For more info go to:
http://www.silverpointdevelopment.com/sptbxlib/support/index.htm