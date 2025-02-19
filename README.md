# PowerShell_LaunchBar

**Toolbar for launch anything as user/admin without asking credential everytime**

--------------------

### Features ✨ 

- 🎯 **Drag & Drop support for easy shortcuts creation & sorting**
- 🔄 **Customizable toolbar position (Top/Bottom)**
- 🎨 **Light and Dark themes**
- 📏 **Adjustable thickness (Small/Medium/Large)**
- 🔑 **Run as administrator option for each shortcut (only 1 UAC request)**
- 🔒 **Securised pipe for admin launch**
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
start "" /d "FOLDER_PATH" PowerShell_LaunchBar "FULLPATH\SHORTCUTS.INI" "FULLPATH\SETTINGS.INI" -force -showdebug
```

  
Multi-line example :
```
start "" /d "FOLDER_PATH"   PowerShell_LaunchBar ^
                            "FULLPATH\SHORTCUTS.INI" ^
                            "FULLPATH\SETTINGS.INI" ^
                            -force -showdebug
```

--------------------

### Admin Pipe Security 🔒

An admin powershell opens and stay in the background after the first shortcut has been launched as admin.  
The next times, the password is not requested.  
The host script (launched as user) communicate with admin script when admin launch is needed for specified shortcuts.  
This implementation uses multiple layers of security to ensure that elevated commands are executed only by the host :

```
Randomized Named Pipe Names:

The command and token pipes are assigned names based on GUIDs (e.g., PSAdminHelperPipe_<GUID>).
This randomness makes it nearly impossible for an attacker to guess the pipe names and establish an unauthorized connection.
```
```
Secure Transmission of Pipe Configuration:

The parent process writes the pipe names into a temporary configuration file.
The helper script, running elevated, reads the file at startup and then deletes it immediately to minimize exposure.
```
```
Dedicated Token Exchange:

The helper generates a secret token (a random GUID) and sends it over a dedicated token pipe.
The client retrieves this token securely without exposing it via command-line arguments or logs.
```
```
Challenge–Response Protocol with HMAC-SHA256:

For every command connection, the helper sends a randomly generated challenge (another GUID).
The client must compute an HMAC-SHA256 using the secret token and the challenge.
The computed HMAC is sent along with the command.
The helper computes its own HMAC for the challenge using the same secret token and compares it to the client's HMAC.
Only if they match is the command executed.
```

--------------------
