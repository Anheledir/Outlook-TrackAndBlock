param(
    [string]$Subject,
    [ValidateRange(1, 1440)]
    [int]$StartMinutes,
    [ValidateRange(1, 1440)]
    [int]$ExtendMinutes,
    [switch]$Private
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------- Robust Foreground/Focus Interop ----------
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class FocusHack {
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern uint  GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
  [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
  [DllImport("user32.dll")] public static extern bool  AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
  [DllImport("user32.dll")] public static extern bool  SetForegroundWindow(IntPtr hWnd);
}
"@

# ---------- Hide/Minimize Console ----------
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class Win32 {
  [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
  [DllImport("user32.dll")]   public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@
$hWnd = [Win32]::GetConsoleWindow()
if ($hWnd -ne [IntPtr]::Zero) { [Win32]::ShowWindow($hWnd, 0) } # 0 = SW_HIDE

# ------------------ Configuration ------------------
$CategoryName    = "Tracking" # TODO: Additional categories, e.g. for different projects?
$CategoryColor   = 6
$MinDuration     = 15
$DurationsStart  = @(30,60,90,120)    # F1-F4
$DurationsExtend = @(30,60,90,120)    # F5-F8
$AllowedStartMinutes = @(0,15,30,45)  # Allowed minute marks for new appointments; set @() to disable alignment
$AlignmentLookAroundMinutes = 10      # Window (± minutes) to align after nearby endings
$SilentExtendDefault = $false         # Set to $true to suppress MessageBox after Extend
$AppDir          = Join-Path $env:LOCALAPPDATA "OutlookTimeTracker"
$LastSubjectFile = Join-Path $AppDir "lastsubject.txt"
$LastPrivateFile = Join-Path $AppDir "lastprivate.txt"
if (-not (Test-Path $AppDir)) { New-Item -ItemType Directory -Path $AppDir -Force | Out-Null }

$StoredSubject = $null
try {
    if (Test-Path $LastSubjectFile) {
        $StoredSubject = Get-Content $LastSubjectFile -Encoding UTF8 -EA SilentlyContinue | Select-Object -First 1
    }
} catch {}
if ($StoredSubject -ne $null) { $StoredSubject = $StoredSubject.Trim() }
$normalizedSlots = [System.Collections.Generic.List[int]]::new()
foreach ($slot in $AllowedStartMinutes) {
    $parsed = 0
    if ([int]::TryParse("$slot", [ref]$parsed)) {
        if ($parsed -ge 0 -and $parsed -lt 60 -and -not $normalizedSlots.Contains($parsed)) {
            [void]$normalizedSlots.Add($parsed)
        }
    }
}
$normalizedSlots.Sort()
$AllowedStartMinutes = $normalizedSlots.ToArray()

$DefaultPrivate = if ($PSBoundParameters.ContainsKey('Private')) {
    [bool]$Private
} elseif (Test-Path $LastPrivateFile) {
    (Get-Content $LastPrivateFile -EA SilentlyContinue | Select-Object -First 1) -eq '1'
} else {
    $false
}

$ResolvedSubject = $null
if ($PSBoundParameters.ContainsKey('Subject')) {
    $ResolvedSubject = $Subject
} elseif (-not [string]::IsNullOrWhiteSpace($StoredSubject)) {
    $ResolvedSubject = $StoredSubject
}
if ($ResolvedSubject -ne $null) { $ResolvedSubject = $ResolvedSubject.Trim() }

$ScriptVersion = '1.2'
$ProjectOwner  = 'Anheledir'
$ProjectRepo   = 'Outlook-TrackAndBlock'
$ProjectUrl    = "https://github.com/$ProjectOwner/$ProjectRepo"

# UI fine-tuning
$BtnWidth        = 150
$BtnHeight       = 42
$Gap             = 12

# ------------------ Theme ------------------
$ClrBg        = [System.Drawing.Color]::FromArgb(36, 38, 41)
$ClrText      = [System.Drawing.Color]::FromArgb(230, 232, 235)
$ClrMutedText = [System.Drawing.Color]::FromArgb(180, 184, 188)
$ClrCard      = [System.Drawing.Color]::FromArgb(47, 50, 54)
$ClrCard2     = [System.Drawing.Color]::FromArgb(52, 56, 60)

$ClrA1 = [System.Drawing.Color]::FromArgb(88, 148, 196)
$ClrA2 = [System.Drawing.Color]::FromArgb(114, 168, 125)
$ClrA3 = [System.Drawing.Color]::FromArgb(168, 142, 214)
$ClrA4 = [System.Drawing.Color]::FromArgb(196, 160, 112)
$ClrPlus   = [System.Drawing.Color]::FromArgb(128, 176, 216)
$ClrCancel = [System.Drawing.Color]::FromArgb(180, 88, 88)

function Lighten([System.Drawing.Color]$c, [int]$d=16) {
    [System.Drawing.Color]::FromArgb(
        [Math]::Min($c.R + $d,255),
        [Math]::Min($c.G + $d,255),
        [Math]::Min($c.B + $d,255)
    )
}

# ------------------ Outlook Helper ------------------
function Get-OutlookApp { try { [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application") } catch { New-Object -ComObject Outlook.Application } }
function Ensure-Category { param([object]$Session) try { $null = $Session.Categories.Item($CategoryName) } catch { $null = $Session.Categories.Add($CategoryName, $CategoryColor) } }

# ------------------ Start Alignment Helpers ------------------
function Get-NextAllowedStartOnOrAfter {
    param(
        [datetime]$Reference,
        [int[]]$AllowedMinutes
    )
    if (-not $AllowedMinutes -or $AllowedMinutes.Count -eq 0) { return $Reference }
    # Ensure ascending, distinct minutes even if caller passes arbitrary input
    $mins = @($AllowedMinutes | Sort-Object -Unique)

    $hourBase = $Reference.Date.AddHours($Reference.Hour)
    foreach ($minute in $mins) {
        $candidate = $hourBase.AddMinutes($minute)
        if ($candidate -ge $Reference) { return $candidate }
    }

    $nextHourBase = $hourBase.AddHours(1)
    $nextMinute   = $mins[0]
    $nextHourBase.AddMinutes($nextMinute)
}

function Get-ClosestAllowedStart {
    param(
        [datetime]$Reference,
        [int[]]$AllowedMinutes
    )
    if (-not $AllowedMinutes -or $AllowedMinutes.Count -eq 0) { return $Reference }

    $candidates = @()
    for ($offset = -1; $offset -le 1; $offset++) {
        $base = $Reference.Date.AddHours($Reference.Hour + $offset)
        foreach ($minute in $AllowedMinutes) {
            $candidate = $base.AddMinutes($minute)
            $candidates += [PSCustomObject]@{
                Time   = $candidate
                Diff   = [math]::Abs(($candidate - $Reference).TotalMinutes)
                Future = ($candidate -ge $Reference)
            }
        }
    }

    $ordered = $candidates | Sort-Object -Property @{Expression={$_.Diff}}, @{Expression={ if ($_.Future) { 0 } else { 1 } }}
    if (-not $ordered -or $ordered.Count -eq 0) { return $Reference }
    $ordered[0].Time
}

function Get-AlignedStartTime {
    param(
        [object]$Session,
        [datetime]$Reference,
        [int[]]$AllowedMinutes,
        [int]$LookAroundMinutes = 10
    )

    if (-not $AllowedMinutes -or $AllowedMinutes.Count -eq 0) { return $Reference }

    $calendar = $Session.GetDefaultFolder(9)
    $items = $calendar.Items
    # Required to expand recurrences reliably in Restrict/Find scenarios
    $items.Sort("[Start]")
    $items.IncludeRecurrences = $true

    # Narrow to a tight window around the reference to avoid scanning entire mailbox
    $lower = $Reference.AddMinutes(-$LookAroundMinutes)
    $upper = $Reference.AddMinutes( $LookAroundMinutes)
    # Outlook Restrict parsing can be picky across locales; using invariant culture and seconds avoids false negatives around minute edges.
    $ci  = [System.Globalization.CultureInfo]::InvariantCulture
    $fmt = "MM/dd/yyyy HH:mm:ss"
    $lowerStr = $lower.ToString($fmt, $ci)
    $upperStr = $upper.ToString($fmt, $ci)
    $filter = "[End] >= '$lowerStr' AND [End] <= '$upperStr'"

    $windowItems = $items.Restrict($filter)

    $nearestFutureEnd = $null
    $nearestFutureDiff = [double]::PositiveInfinity
    $nearestPastEnd = $null
    $nearestPastDiff = [double]::PositiveInfinity

    $item = $windowItems.GetFirst()
    while ($item) {
        try {
            if ($item.MessageClass -like "IPM.Appointment*") {
                $end = $item.End
                $diffMinutes = ($end - $Reference).TotalMinutes
                if ($diffMinutes -ge 0) {
                    if ($diffMinutes -le $LookAroundMinutes -and $diffMinutes -lt $nearestFutureDiff) {
                        $nearestFutureEnd = $end
                        $nearestFutureDiff = $diffMinutes
                    }
                } else {
                    $abs = [math]::Abs($diffMinutes)
                    if ($abs -le $LookAroundMinutes -and $abs -lt $nearestPastDiff) {
                        $nearestPastEnd = $end
                        $nearestPastDiff = $abs
                    }
                }
            }
        } catch {}
        $item = $windowItems.GetNext()
    }

    if ($nearestFutureEnd) { return Get-NextAllowedStartOnOrAfter -Reference $nearestFutureEnd -AllowedMinutes $AllowedMinutes }
    if ($nearestPastEnd)   { return Get-NextAllowedStartOnOrAfter -Reference $nearestPastEnd   -AllowedMinutes $AllowedMinutes }

    Get-ClosestAllowedStart -Reference $Reference -AllowedMinutes $AllowedMinutes
}

function New-TrackingAppointment {
    param([string]$Subject,[int]$DurationMinutes = 30,[bool]$Private=$false)
    $outlook = Get-OutlookApp
    if (-not $outlook) {
        [System.Windows.Forms.MessageBox]::Show("Could not start Outlook.","Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null; return
    }
    $session = $outlook.Session; Ensure-Category -Session $session
    $now  = Get-Date; $mins = [math]::Max($MinDuration, [int]$DurationMinutes)
    $start = $now
    if ($AllowedStartMinutes.Count -gt 0) {
        $start = Get-AlignedStartTime -Session $session -Reference $now -AllowedMinutes $AllowedStartMinutes -LookAroundMinutes $AlignmentLookAroundMinutes
    }
    $finalSubject = if ([string]::IsNullOrWhiteSpace($Subject)) { $CategoryName } else { $Subject.Trim() }
    $appt = $outlook.CreateItem(1)
    $appt.Subject=$finalSubject; $appt.Start=$start; $appt.End=$start.AddMinutes($mins)
    $appt.Categories=$CategoryName; $appt.BusyStatus=2; $appt.ReminderSet=$false
    if ($Private) { $appt.Sensitivity=2 } else { $appt.Sensitivity=0 }
    $appt.Body = "Automatically created via script on $($now.ToString('yyyy-MM-dd HH:mm'))."
    $appt.Save()
    try { Set-Content -Path $LastSubjectFile -Value $finalSubject -Encoding UTF8 -Force } catch {}
}

function Get-CurrentAppointment {
    param([bool]$PreferTracking = $true)
    $outlook = Get-OutlookApp; if (-not $outlook) { return $null }
    $ns=$outlook.Session; $cal=$ns.GetDefaultFolder(9); $items=$cal.Items
    $items.IncludeRecurrences=$true
    $now = Get-Date
    $currentItems=@()
    foreach ($it in $items) {
        try {
            if ($it.MessageClass -like "IPM.Appointment*" -and $it.Start -le $now -and $it.End -gt $now) {
                $currentItems += $it
            }
        } catch {}
    }
    if ($currentItems.Count -gt 0) {
        if ($PreferTracking) {
            foreach ($it in $currentItems) {
                try {
                    if ($it.Categories -and ($it.Categories -split ';' | % { $_.Trim() }) -contains $CategoryName) { return $it }
                } catch {}
            }
        }
        return $currentItems[0]
    }
    $null
}

function Extend-CurrentAppointment {
    param([int]$AddMinutes, [switch]$Silent)
    $appt = Get-CurrentAppointment -PreferTracking:$true
    if (-not $appt) {
        if (-not $Silent) { [System.Windows.Forms.MessageBox]::Show("No running appointment found.","Nothing to extend",
            [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null }
        return
    }
    $appt.End = $appt.End.AddMinutes([int]$AddMinutes)
    $append = "`r`n[+$AddMinutes min on $((Get-Date).ToString('yyyy-MM-dd HH:mm'))]"
    try { if ([string]::IsNullOrEmpty($appt.Body)) { $appt.Body=$append.Trim() } else { $appt.Body += $append } } catch {}
    $appt.Save()
    if (-not $Silent) {
        [System.Windows.Forms.MessageBox]::Show("Appointment extended until $($appt.End.ToString('HH:mm')).","Extended (+$AddMinutes)",
            [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    }
}

function Check-ForUpdate {
    param([string]$CurrentVersion)
    try {
        $tags = Invoke-RestMethod -Uri "https://api.github.com/repos/$ProjectOwner/$ProjectRepo/tags" -UseBasicParsing -Headers @{ 'User-Agent' = 'OutlookTimeTracker' }
        if ($tags) {
            $latest = ($tags | Sort-Object { [version]$_.name } -Descending | Select-Object -First 1).name
            if ([version]$latest -gt [version]$CurrentVersion) { return $latest }
        }
    } catch {}
    $null
}

# ------------------ UI Helpers ------------------
function Round-Control($ctrl, [int]$radius = 8) {
    $gp = New-Object System.Drawing.Drawing2D.GraphicsPath
    $r = $radius; $rect = New-Object System.Drawing.Rectangle(0,0,$ctrl.Width,$ctrl.Height)
    $gp.AddArc($rect.X,$rect.Y,$r,$r,180,90); $gp.AddArc($rect.Right-$r,$rect.Y,$r,$r,270,90)
    $gp.AddArc($rect.Right-$r,$rect.Bottom-$r,$r,$r,0,90); $gp.AddArc($rect.X,$rect.Bottom-$r,$r,$r,90,90)
    $gp.CloseAllFigures(); $ctrl.Region = New-Object System.Drawing.Region($gp)
}

function New-NiceButton {
    param(
        [string]$Text,[int]$TagValue,
        [System.Drawing.Color]$Back,[System.Drawing.Color]$Fore,[System.Drawing.Color]$Hover
    )
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text=$Text; $btn.Tag=$TagValue
    $btn.AutoSize = $false
    $btn.Width    = $BtnWidth
    $btn.Height   = $BtnHeight
    $btn.Margin   = New-Object System.Windows.Forms.Padding(0,0,$Gap,$Gap)
    $btn.Padding  = New-Object System.Windows.Forms.Padding(10,6,10,6)
    $btn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btn.FlatAppearance.BorderSize = 0
    $btn.BackColor = $Back; $btn.ForeColor=$Fore
    $btn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btn.Add_MouseEnter({ param($s,$e) $s.BackColor = $Hover })
    $btn.Add_MouseLeave({ param($s,$e) $s.BackColor = $Back })
    $btn.Add_Resize({ param($s,$e) Round-Control $s 10 })
    $btn
}

# ---------- Foreground-Fix Helper (Timer + AttachThreadInput) ----------
function Invoke-ForegroundFix {
    param([System.Windows.Forms.Form]$Form, [System.Windows.Forms.Control]$FocusControl)

    $script:__ffTry = 0
    $script:__ffDone = $false
    $attempts = 5
    $t = New-Object System.Windows.Forms.Timer
    $t.Interval = 60
    $t.Add_Tick({
        if ($script:__ffDone) { $t.Stop(); $t.Dispose(); return }

        $fg = [FocusHack]::GetForegroundWindow()
        [uint32]$fgPid = 0
        $fgTid = [FocusHack]::GetWindowThreadProcessId($fg, [ref]$fgPid)
        $curTid = [FocusHack]::GetCurrentThreadId()

        if ($fgTid -ne 0 -and $curTid -ne 0) {
            [FocusHack]::AttachThreadInput($curTid, $fgTid, $true) | Out-Null
            $null = $Form.Activate()
            $Form.BringToFront()
            [FocusHack]::SetForegroundWindow($Form.Handle) | Out-Null
            [FocusHack]::AttachThreadInput($curTid, $fgTid, $false) | Out-Null
        } else {
            $null = $Form.Activate()
            $Form.BringToFront()
        }

        if ($FocusControl) { $FocusControl.Focus(); $FocusControl.SelectAll() }

        $script:__ffTry++
        if ($script:__ffTry -ge $attempts) {
            $script:__ffDone = $true
            $t.Stop(); $t.Dispose()
        }
    })
    $t.Start()
}

# ------------------ CLI Mode ------------------
$cliStartRequested  = $PSBoundParameters.ContainsKey('StartMinutes')
$cliExtendRequested = $PSBoundParameters.ContainsKey('ExtendMinutes')

if ($cliStartRequested -and $cliExtendRequested) {
    [System.Windows.Forms.MessageBox]::Show("Use -StartMinutes or -ExtendMinutes, not both.","Invalid parameters",
        [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
    Write-Error "Use -StartMinutes or -ExtendMinutes, not both."
    $global:LASTEXITCODE = 1
    return
}

if ($cliStartRequested) {
    if ($StartMinutes -le 0) {
        [System.Windows.Forms.MessageBox]::Show("StartMinutes must be a positive number of minutes.","Invalid StartMinutes",
            [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        Write-Error "StartMinutes must be a positive integer."
        $global:LASTEXITCODE = 1
        return
    }

    New-TrackingAppointment -Subject $ResolvedSubject -DurationMinutes $StartMinutes -Private:$DefaultPrivate
    try { Set-Content -Path $LastPrivateFile -Value (if($DefaultPrivate){'1'}else{'0'}) -Encoding UTF8 -Force } catch {}
    return
}

if ($cliExtendRequested) {
    if ($ExtendMinutes -le 0) {
        [System.Windows.Forms.MessageBox]::Show("ExtendMinutes must be a positive number of minutes.","Invalid ExtendMinutes",
            [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        Write-Error "ExtendMinutes must be a positive integer."
        $global:LASTEXITCODE = 1
        return
    }

    Extend-CurrentAppointment -AddMinutes $ExtendMinutes -Silent:$SilentExtendDefault
    return
}

# ------------------ GUI ------------------
$form                        = New-Object System.Windows.Forms.Form
$form.Text                   = "Track & Block"
$form.StartPosition          = "CenterScreen"
$form.Size                   = New-Object System.Drawing.Size(780, 560)
$form.MinimumSize            = New-Object System.Drawing.Size(760, 520)
$form.TopMost                = $true
$form.KeyPreview             = $true
$form.AutoScaleMode          = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.Font                   = New-Object System.Drawing.Font("Segoe UI", 10.5)
$form.BackColor              = $ClrBg
$form.ForeColor              = $ClrText

# Root table
$root = New-Object System.Windows.Forms.TableLayoutPanel
$root.Dock = 'Fill'
$root.ColumnCount = 1
$root.RowCount    = 6
$root.Padding     = New-Object System.Windows.Forms.Padding($Gap,$Gap,$Gap,$Gap)
$root.BackColor   = $ClrBg
$root.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) ) | Out-Null
$root.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) ) | Out-Null
$root.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) ) | Out-Null
$root.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) ) | Out-Null
$root.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) ) | Out-Null
$root.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) ) | Out-Null

# Header + task input
$labelTask = New-Object System.Windows.Forms.Label
$labelTask.Text = "Task / Topic"
$labelTask.AutoSize = $true
$labelTask.Margin = New-Object System.Windows.Forms.Padding(0,0,0,6)
$labelTask.Font = New-Object System.Drawing.Font($form.Font, [System.Drawing.FontStyle]::Bold)
$labelTask.ForeColor = $ClrText

$textTask = New-Object System.Windows.Forms.TextBox
$textTask.BorderStyle=[System.Windows.Forms.BorderStyle]::FixedSingle
$textTask.BackColor=[System.Drawing.Color]::FromArgb(58,61,66)
$textTask.ForeColor=$ClrText
$textTask.Margin = New-Object System.Windows.Forms.Padding(0,0,0,$Gap)
$textTask.Anchor = 'Left,Right'
$textTask.Width  = 700
if ($ResolvedSubject -ne $null) { $textTask.Text = $ResolvedSubject }

$chkPrivate = New-Object System.Windows.Forms.CheckBox
$chkPrivate.Text = "Private appointment"
$chkPrivate.AutoSize = $true
$chkPrivate.Checked = $DefaultPrivate
$chkPrivate.Margin = New-Object System.Windows.Forms.Padding(0,0,0,$Gap)
$chkPrivate.ForeColor = $ClrText

function New-Card([string]$title, [System.Drawing.Color]$bg) {
    $card = New-Object System.Windows.Forms.TableLayoutPanel
    $card.BackColor = $bg
    $card.AutoSize  = $true
    $card.Margin    = New-Object System.Windows.Forms.Padding(0,0,0,$Gap)
    $card.Padding   = New-Object System.Windows.Forms.Padding(10,10,10,10)
    $card.ColumnCount = 1; $card.RowCount=2
    $card.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) ) | Out-Null
    $card.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) ) | Out-Null
    $card.Anchor = 'Left,Right'

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $title
    $lbl.AutoSize = $true
    $lbl.Margin = New-Object System.Windows.Forms.Padding(4,0,0,8)
    $lbl.ForeColor = $ClrMutedText

    $row = New-Object System.Windows.Forms.FlowLayoutPanel
    $row.FlowDirection = 'LeftToRight'
    $row.WrapContents  = $true
    $row.AutoSize      = $true
    $row.Margin        = New-Object System.Windows.Forms.Padding(2,0,0,0)
    $row.BackColor     = $bg

    $card.Controls.Add($lbl, 0,0)
    $card.Controls.Add($row, 0,1)
    $card.Add_Resize({ param($s,$e) Round-Control $s 10 })
    return @($card,$row)
}

# Cards
$cardStart,  $panelStart  = New-Card "Start new (sets Teams to 'Busy')" $ClrCard
$cardExtend, $panelExtend = New-Card "Extend running appointment" $ClrCard2

# Tooltips
$tt = New-Object System.Windows.Forms.ToolTip
$tt.AutoPopDelay=6000; $tt.InitialDelay=250; $tt.ReshowDelay=100
$tt.SetToolTip($chkPrivate, "Mark appointment as private")

# Start buttons
$startButtons=@(); $accents=@($ClrA1,$ClrA2,$ClrA3,$ClrA4)
for ($i=0; $i -lt $DurationsStart.Count; $i++) {
    $mins=$DurationsStart[$i]; $base=$accents[$i]
    $btn = New-NiceButton "⏱  $mins min  (F$($i+1))" $mins $base $ClrBg (Lighten $base 18)
    $btn.Add_Click([System.EventHandler]{ param($sender,$e)
        New-TrackingAppointment -Subject $textTask.Text -DurationMinutes ([int]$sender.Tag) -Private:$chkPrivate.Checked
        $form.Close()
    })
    $tt.SetToolTip($btn, "Start a $mins-minute block (Teams = Busy)")
    $panelStart.Controls.Add($btn) | Out-Null
    $startButtons += $btn
}
$form.AcceptButton = $startButtons[0]

# Extend buttons
$extendButtons=@()
for ($i=0; $i -lt $DurationsExtend.Count; $i++) {
    $mins=$DurationsExtend[$i]
    $btn = New-NiceButton "＋  +$mins min  (F$($i+5))" $mins $ClrPlus $ClrBg (Lighten $ClrPlus 18)
    $btn.Add_Click([System.EventHandler]{ param($sender,$e) Extend-CurrentAppointment -AddMinutes ([int]$sender.Tag) -Silent:$SilentExtendDefault })
    $tt.SetToolTip($btn, "Extend current appointment by $mins minutes")
    $panelExtend.Controls.Add($btn) | Out-Null
    $extendButtons += $btn
}

# Bottom row: Cancel on the right
$bottomRow = New-Object System.Windows.Forms.TableLayoutPanel
$bottomRow.ColumnCount=2; $bottomRow.RowCount=1
$bottomRow.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)) ) | Out-Null
$bottomRow.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)) ) | Out-Null
$bottomRow.Dock='Top'; $bottomRow.AutoSize=$true
$bottomRow.Margin = New-Object System.Windows.Forms.Padding(0,0,0,0)
$bottomRow.Anchor = 'Left,Right'

$infoHost = New-Object System.Windows.Forms.FlowLayoutPanel
$infoHost.FlowDirection='LeftToRight'
$infoHost.WrapContents=$false
$infoHost.AutoSize=$true
$infoHost.Dock='Top'
$infoHost.Margin = New-Object System.Windows.Forms.Padding(0,0,0,0)

$lnkGit = New-Object System.Windows.Forms.LinkLabel
$lnkGit.Text='GitHub'
$lnkGit.AutoSize=$true
$lnkGit.LinkColor=$ClrMutedText
$lnkGit.ActiveLinkColor=(Lighten $ClrMutedText 40)
$lnkGit.VisitedLinkColor=$ClrMutedText
$lnkGit.Margin = New-Object System.Windows.Forms.Padding(0,12,8,0)
$lnkGit.Links.Add(0,$lnkGit.Text.Length,$ProjectUrl)
$lnkGit.Add_LinkClicked({ param($s,$e) Start-Process $e.Link.LinkData })
$infoHost.Controls.Add($lnkGit) | Out-Null

$labelUpdate = New-Object System.Windows.Forms.LinkLabel
$labelUpdate.AutoSize=$true
$labelUpdate.LinkColor=$ClrMutedText
$labelUpdate.ActiveLinkColor=(Lighten $ClrMutedText 40)
$labelUpdate.VisitedLinkColor=$ClrMutedText
$labelUpdate.Margin = New-Object System.Windows.Forms.Padding(0,12,0,0)
$labelUpdate.Visible=$false
$infoHost.Controls.Add($labelUpdate) | Out-Null

$cancelHost = New-Object System.Windows.Forms.FlowLayoutPanel
$cancelHost.FlowDirection='RightToLeft'
$cancelHost.WrapContents=$false
$cancelHost.AutoSize=$true
$cancelHost.Dock='Top'
$cancelHost.Margin = New-Object System.Windows.Forms.Padding(0,0,0,0)

$btnCancel = New-NiceButton "✖  Cancel  (Esc)" 0 $ClrCancel $ClrBg (Lighten $ClrCancel 18)
$btnCancel.Add_Click({ $form.Close() })
$cancelHost.Controls.Add($btnCancel) | Out-Null

$bottomRow.Controls.Add($infoHost,0,0)
$bottomRow.Controls.Add($cancelHost,1,0)

$latest = Check-ForUpdate -CurrentVersion $ScriptVersion
if ($latest) {
    $labelUpdate.Text = "Update $latest available"
    $labelUpdate.Links.Add(0,$labelUpdate.Text.Length,"$ProjectUrl/releases/tag/$latest")
    $labelUpdate.Visible = $true
}

$form.Add_FormClosing({
    try {
        Set-Content -Path $LastPrivateFile -Value (if($chkPrivate.Checked){'1'}else{'0'}) -Encoding UTF8 -Force
    } catch {}
})

# Key bindings
$form.Add_KeyDown({
    if ($_.Alt) {
        if ($_.KeyCode -eq 'F4') { $form.Close() }
        return
    }
    switch ($_.KeyCode) {
        'F1' { $startButtons[0].PerformClick(); break }
        'F2' { $startButtons[1].PerformClick(); break }
        'F3' { $startButtons[2].PerformClick(); break }
        'F4' { $startButtons[3].PerformClick(); break }
        'F5' { $extendButtons[0].PerformClick(); break }
        'F6' { $extendButtons[1].PerformClick(); break }
        'F7' { $extendButtons[2].PerformClick(); break }
        'F8' { $extendButtons[3].PerformClick(); break }
        'Escape' { $btnCancel.PerformClick(); break }
    }
})

# Assembly
$root.Controls.Add($labelTask,  0,0)
$root.Controls.Add($textTask,   0,1)
$root.Controls.Add($chkPrivate, 0,2)
$root.Controls.Add($cardStart,  0,3)
$root.Controls.Add($cardExtend, 0,4)
$root.Controls.Add($bottomRow,  0,5)
$form.Controls.Add($root)

# keep right edge aligned
$form.Add_Resize({
    $textTask.Width = $root.ClientSize.Width - $root.Padding.Left - $root.Padding.Right
    $cardStart.Width = $textTask.Width
    $cardExtend.Width = $textTask.Width
})

# --- Reliable focus/foreground (variant A) ---
$form.Add_Shown({
    $form.TopMost = $true
    [System.Windows.Forms.Application]::DoEvents()
    Invoke-ForegroundFix -Form $form -FocusControl $textTask
})

# Dialog
[void]$form.ShowDialog()
