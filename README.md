# Outlook-TrackAndBlock

![🇩🇪 German version of this file](/README_de.md)

PowerShell tool with both **GUI** and **CLI** that creates/extends private "tracking" appointments in Outlook.  
This automatically sets **Microsoft Teams** presence to **"Busy"** — perfect as a Stream Deck action.

![Track & Block — Screenshot](/assets/screenshot.png?raw=true)

## ✨ Features
- 🗓️ **Start new focus blocks** (30/60/90/120 min) — F1–F4
- ➕ **Extend the current appointment** (+30/+60/+90/+120 min) — F5–F8
- 🔒 Appointments are ~~**private**~~ and categorized **"Tracking"**
- 🖥️ **Dark-ish** WinForms dialog, DPI-aware, focus fix (AttachThreadInput)
- 🧰 **CLI mode** for direct use without GUI (e.g., Stream Deck)
- 🪟 Console is **hidden**; start with `-WindowStyle Hidden`

## ⚙️ Requirements
- Windows 10/11
- Outlook Desktop (Microsoft 365 / Office)
- PowerShell 5.1 **or** 7.x (WinForms available)

## 🚀 Quickstart
1. Save the script `Outlook_Timetracker.ps1` from `/scripts`.
2. Test:

~~~powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\Path\Outlook_Timetracker.ps1"
~~~

3. Press F1/F2/F3/F4 to start a new block · F5–F8 to extend the running appointment.

## 🎛️ Stream Deck Integration
- **Action:** System → *Open*
- **Program:** `C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe`
- **Arguments (GUI):**

~~~powershell
-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\Path\Outlook_Timetracker.ps1"
~~~

- **Arguments (CLI — no GUI, 90 min "Deep Work"):**

~~~powershell
-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\Path\Outlook_Timetracker.ps1" -Subject "Deep Work" -StartMinutes 90
~~~

- **Arguments (extend only, +30 min):**

~~~powershell
-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\Path\Outlook_Timetracker.ps1" -ExtendMinutes 30
~~~

## 🔧 Configuration (at the top of the script)
- `$CategoryName = "Tracking"` (created automatically if missing)
- `$DurationsStart / $DurationsExtend` — button minutes
- `$BtnWidth / $BtnHeight` — button sizes
- Theme colors (dark/subtle) are defined as variables
- Optional: `$SilentExtendDefault = $true` (disable MessageBox after "Extend")

## 🧪 CLI Parameters (optional)
~~~powershell
-Subject <string>        # Task/subject name
-StartMinutes <int>      # Start a block immediately (skips GUI)
-ExtendMinutes <int>     # Extend the currently running appointment
~~~

## ❓ FAQ
**Does this set Teams "Do Not Disturb (DND)"?**  
No — Outlook calendar sets Teams to **"Busy"**. For true Teams DND use separate measures (e.g., Windows Focus Assist or UI automation).

**Why the US date format internally?**  
Outlook's `Restrict` API requires `MM/dd/yyyy HH:mm`. The script handles this for you.

## 🛠️ Troubleshooting
- **ExecutionPolicy:** Start with `-ExecutionPolicy Bypass`.
- **Category not visible:** In Outlook calendar view, enable the "Categories" column.
- **No running appointment detected:** Ensure there's an event with Start ≤ now < End (recurrences supported).
- **Dialog not focused:** The foreground fix is included; if desktop policies are strict, try running Stream Deck "as Administrator".

## 🔐 Privacy
- Appointments are created **locally** via Outlook COM (no cloud API).
- **No data leaves your machine.**

## 📦 Structure
~~~
/scripts/Outlook_Timetracker.ps1
/assets/screenshot.png
/LICENSE
/README.md
~~~

## 🏷️ Topics / Tags
`powershell`, `outlook`, `microsoft-teams`, `time-tracking`, `stream-deck`, `calendar`, `windows`, `productivity`, `winforms`, `com-interop`, `focus-time`

## 📜 License
MIT — see `LICENSE`.

## 🤝 Contributing
Issues and PRs welcome! For PRs, please:
- keep commits compact (Conventional Commits optional)
- add short comments for Outlook interop or UI changes
- briefly describe how you tested
