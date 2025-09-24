# Outlook-TrackAndBlock

[🇩🇪 German version of this file](/README_de.md)

PowerShell tool with both **GUI** and **CLI** that creates/extends private "tracking" appointments in Outlook.  
This automatically sets **Microsoft Teams** presence to **"Busy"** — perfect as a Stream Deck action.

![Track & Block — Screenshot](/assets/screenshot.png?raw=true)

## Features

- **Start new focus blocks** (30/60/90/120 min) — F1-F4
- **Aligned start times** — new blocks snap to configurable minute slots (default 00/15/30/45) and respect nearby bookings
- **Extend the current appointment** (+30/+60/+90/+120 min) — F5-F8
- Appointments can be marked private and categorized **"Tracking"**
- **Dark-ish** WinForms dialog, DPI-aware, focus fix (AttachThreadInput)
- **CLI mode** for direct use without GUI (e.g., Stream Deck)
- Console is **hidden**; start with `-WindowStyle Hidden`

## Requirements

- Windows 10/11
- Outlook Desktop (Microsoft 365 / Office)
- PowerShell 5.1 **or** 7.x (WinForms available)

## Quickstart

1. Save the script `Outlook_Timetracker.ps1` from `/scripts`.
2. Test:

~~~powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\Path\Outlook_Timetracker.ps1"
~~~

3. Press F1/F2/F3/F4 to start a new block · F5-F8 to extend the running appointment.

## Stream Deck Integration

- **Action:** System → *Open*
- **Program (Windows PowerShell 5.1):** `C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe`
- **Program (PowerShell 7+):** `C:\Program Files\PowerShell\7\pwsh.exe`
- **Arguments (GUI):**

~~~powershell
-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\Path\Outlook_Timetracker.ps1"
~~~

- **Arguments (CLI — no GUI, 90 min "Deep Work"):**

~~~powershell
-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\Path\Outlook_Timetracker.ps1" -Subject 'Deep Work' -StartMinutes 90
~~~

- **Arguments (extend only, +30 min):**

~~~powershell
-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\Path\Outlook_Timetracker.ps1" -ExtendMinutes 30
~~~

## Configuration (at the top of the script)

- `$CategoryName = "Tracking"` (created automatically if missing)
- `$DurationsStart / $DurationsExtend` — button minutes
- `$AllowedStartMinutes` — minute marks for start alignment (e.g. `@(0,15,30,45)`); valid minutes are 0-59; duplicates ignored; use `@()` to disable rounding  
- `$AlignmentLookAroundMinutes` — window (in minutes) to treat appointments that just ended or end soon as "nearby" (default: 10)
- `$BtnWidth / $BtnHeight` — button sizes
- Theme colors (dark/subtle) are defined as variables
- Optional: `$SilentExtendDefault = $true` (disable MessageBox after "Extend")

When alignment is active, the script checks for nearby appointments that just ended or end within `$LookAroundMinutes` and starts the new block right after them; otherwise it rounds to the closest allowed slot.

## CLI Parameters (optional)

~~~powershell
-Subject <string>        # Task/subject name
-StartMinutes <int>      # Start a block immediately with this duration in minutes (skips GUI)
-ExtendMinutes <int>     # Extend the currently running appointment
-Private                 # Switch — mark the next block as private (also pre-checks the GUI)
~~~

`-StartMinutes` and `-ExtendMinutes` are mutually exclusive; provide a positive value for exactly one of them. `-Subject` can be
used on its own to pre-fill the GUI.

## FAQ

**Does this set Teams "Do Not Disturb (DND)"?**  
No — Outlook calendar sets Teams to **"Busy"**. For true Teams DND use separate measures (e.g., Windows Focus Assist or UI automation).

**Why the US date format internally?**  
Outlook's `Restrict` API requires `MM/dd/yyyy HH:mm`. The script handles this for you.

## Troubleshooting

- **ExecutionPolicy:** Start with `-ExecutionPolicy Bypass`.
- **Category not visible:** In Outlook calendar view, enable the "Categories" column.
- **No running appointment detected:** Ensure there's an event with Start ≤ now < End (recurrences supported).
- **Dialog not focused:** The foreground fix is included; if desktop policies are strict, try running Stream Deck "as Administrator".

## Architecture

```mermaid
sequenceDiagram
  autonumber
  participant U as User / CLI
  participant PS as Outlook_Timetracker.ps1
  participant OL as Outlook Session
  participant CAL as Calendar Items

  Note over PS: v1.2 — start-time alignment flow
  U->>PS: Request new tracking appointment
  PS->>PS: Load/normalize AllowedStartMinutes (e.g., 0,15,30,45)
  alt AllowedStartMinutes empty
    PS->>PS: start = Now
  else AllowedStartMinutes set
    PS->>PS: candidate = Get-ClosestAllowedStart(Now, AllowedStartMinutes)
    PS->>OL: Open session
    PS->>CAL: Query items within AlignmentLookAroundMinutes
    alt Nearby event ended/ends soon
      PS->>PS: start = Get-AlignedStartTime(..., favor immediate-after)
    else No nearby items
      PS->>PS: start = candidate
    end
  end
  PS->>PS: end = start + duration (or ExtendMinutes)
  PS->>OL: Create appointment (start, end)
  OL-->>U: Appointment created
```

## Privacy

- Appointments are created **locally** via Outlook COM (no cloud API).
- **No data leaves your machine.**

## Topics / Tags

`powershell`, `outlook`, `microsoft-teams`, `time-tracking`, `stream-deck`, `calendar`, `windows`, `productivity`, `winforms`, `com-interop`, `focus-time`

## Contributing

Issues and PRs welcome! For PRs, please:

- keep commits compact (Conventional Commits optional)
- add short comments for Outlook interop or UI changes
