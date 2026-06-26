# BUDDY AI Assistant — Setup Guide

## Files in This Folder
```
BUDDY_System/
├── buddy.py            ← Python backend (OS control engine)
├── buddy_ui.html       ← JARVIS-style visual interface
├── setup_and_run.bat   ← First-time setup + launch
└── run_buddy.bat       ← Launch after setup is done
```

---

## Quick Start (First Time)

1. Put all 4 files in the **same folder** (e.g. `C:\BUDDY\`)
2. Double-click **`setup_and_run.bat`**
3. It will install dependencies and launch BUDDY automatically
4. The JARVIS UI opens in your browser
5. Click **"Initialize Mic Access"** → allow microphone
6. Say **"Wake up Buddy"** → Buddy greets you!

---

## After First Setup

Just double-click **`run_buddy.bat`** every time.

---

## Voice Commands

### System
| Say | What Happens |
|-----|-------------|
| "Wake up Buddy" | Activates Buddy with intro message |
| "System status" | Reports CPU, RAM, battery |
| "What time is it" | Tells the time |
| "What's the date" | Tells today's date |
| "Battery status" | Reports battery level |
| "Go to sleep" | Puts Buddy in standby |
| "Shut down my PC" | Shuts down Windows in 5s |
| "Restart my PC" | Restarts Windows in 5s |
| "Sleep my PC" | Hibernates the computer |

### Open Apps
| Say | What Happens |
|-----|-------------|
| "Open VS Code" | Launches VS Code |
| "Open VS Code file hello.py" | Opens hello.py in VS Code |
| "Open MS Word" | Launches Microsoft Word |
| "Open Excel" | Launches Excel |
| "Open PowerPoint" | Launches PowerPoint |
| "Open Notepad" | Launches Notepad |
| "Open Notepad file myfile.txt" | Opens file in Notepad |
| "Open File Explorer" | Opens File Explorer |
| "Open File Explorer at C:\Users" | Opens Explorer at path |
| "Open Calculator" | Opens Calculator |
| "Open Task Manager" | Opens Task Manager |
| "Open Command Prompt" | Opens CMD |
| "Open PowerShell" | Opens PowerShell |
| "Open Paint" | Opens MS Paint |
| "Open Chrome" | Opens Google Chrome |
| "Open Spotify" | Opens Spotify |
| "Open Discord" | Opens Discord |

### Create Documents
| Say | What Happens |
|-----|-------------|
| "Write a letter" | Creates a Word letter template |
| "Write a resume" | Creates a Word CV template |
| "Write a report" | Creates a Word report template |
| "Write an email" | Creates a Word email draft |
| "Create a Python file named hello" | Creates hello.py in VS Code |
| "Create an HTML file named index" | Creates index.html in VS Code |
| "Write a JavaScript script named app" | Creates app.js in VS Code |

### Web & Media
| Say | What Happens |
|-----|-------------|
| "Open YouTube" | Opens YouTube |
| "Play Believer on YouTube" | Searches YouTube for the song |
| "Search for Python tutorials" | Google search |
| "Open Google" | Opens Google |
| "Open Gmail" | Opens Gmail |
| "Open GitHub" | Opens GitHub |
| "Open Netflix" | Opens Netflix |

### System Control
| Say | What Happens |
|-----|-------------|
| "Volume 50" | Sets volume to 50% |
| "Mute" | Mutes audio |
| "Volume up" | Increases volume |
| "Volume down" | Decreases volume |
| "Take a screenshot" | Saves screenshot to Pictures |

---

## Customizing App Paths

Open `buddy.py` and find the `APP_PATHS` section (around line 60).
Edit paths to match where YOUR apps are installed:

```python
APP_PATHS = {
    "vscode": r"C:\Users\YourName\AppData\Local\Programs\Microsoft VS Code\Code.exe",
    "word":   r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
    # etc...
}
```

---

## How It Works

```
You speak → Browser captures voice → 
Speech Recognition converts to text →
HTML UI sends text to Python backend (localhost:8765) →
Python executes OS command (open app, create file, etc.) →
Python speaks response via pyttsx3 →
UI displays response with typing animation
```

The Python backend runs a tiny local web server on port 8765.
The HTML UI connects to it automatically.
If the backend is offline, the UI falls back to browser-only mode
(web commands work, but OS commands like opening apps won't).

---

## Requirements
- Windows 10 / 11
- Python 3.8+
- Google Chrome or Microsoft Edge (for the UI + voice recognition)
- Microphone
