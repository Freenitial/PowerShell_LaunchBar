# PowerShell_LaunchBar

**Launch anything as user/admin without asking credential everytime**

--------------------

### Features ✨ 

- 🎯 **Drag & Drop support for easy shortcuts creation**
- 🔄 **Customizable toolbar position (Top/Bottom)**
- 🎨 **Light and Dark themes**
- 📏 **Adjustable thickness (Small/Medium/Large)**
- 🚀 **Run as administrator option for each shortcut (single UAC request)**
- 📝 **Customizable text display**
- ↔️ **Left/Right alignment options for each shortcut**
- 💾 **Import/Export shortcuts configuration**
- 🔄 **Real-time updates and changes**
- 🎯 **Smart positioning with Windows workspace**

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

1. Download **PowerShell_LaunchBar.bat**
2. Double-click to run

Settings are stored in `%LOCALAPPDATA%\Powershell_Toolbar`

--------------------

### To improve (feel free to submit pull request)

- DPI aware is very poorly implemented. 
The scaling detected when the program is opened determines that the toolbar will be dirty-scaled itself the next time Windows parameters are changed.
Instead, the toolbar should be rebuilt as a fresh script opening on a given Windows scale.
- No overflow management when too many shortcuts are present
