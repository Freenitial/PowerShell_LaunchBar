# PowerShell_LaunchBar

**Toolbar for launch anything as user/admin without asking credential everytime**

--------------------

### Features ✨ 

- 🎯 **Drag & Drop support for easy shortcuts creation & sorting**
- 🔄 **Customizable toolbar position (Top/Bottom)**
- 🎨 **Light and Dark themes**
- 📏 **Adjustable thickness (Small/Medium/Large)**
- 🔑 **Run as administrator option for each shortcut (only 1 UAC request)**
- 🔒 **Securised Pipeline instance for admin launch**
- 📝 **Customizable shortcut titles**
- ↔️ **Left/Right alignment options for each shortcut**
- 💾 **Import/Export shortcuts configuration**
- 🖥️ **Real-time updates and changes, DPI Aware**
- 🚀 **Just a lite PowerShell script providing full interface**

--------------------

![image](https://github.com/user-attachments/assets/a80468f3-a77c-4b53-9ffc-5122dcc06efb)

--------------------

### Usage 📝

1. **Add Shortcuts**:
   - Drag & Drop files or folders onto the toolbar
   - Right-click shortcuts for additional options

2. **Customize Appearance**:
   - Click the settings gear icon
   - Adjust toolbar position, size, and theme
   - Configure default behavior for new shortcuts

3. **Manage Shortcuts**:
   - Click and drag to reorder
   - Right-click for rename/remove options
   - Import/Export configurations in settings menu

--------------------

### Installation 🔧

_Requirement : Windows 10 build 1607 +_

1. Download **PowerShell_LaunchBar.bat**
2. Double-click to run

Configuration stored in `%LOCALAPPDATA%\Powershell_Toolbar`

--------------------

### Command line 💉 :

To start from other batch or cmd without exit, launch like this :  
```
start "" /d "FOLDER\CONTAINING_batchfile" PowerShell_LaunchBar
```

To start from other batch or cmd without exit **+ FORCE IMPORT SHORTCUTS FILE**, launch like this :  
```
start "" /d "FOLDER\CONTAINING_batchfile" PowerShell_LaunchBar "FULLPATH\TO_IMPORT\SHORTCUT.INI" /f
```

--------------------
