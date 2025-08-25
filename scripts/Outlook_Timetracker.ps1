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

# ---------- Konsole verstecken/minimieren ----------
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

# ------------------ Konfiguration ------------------
$CategoryName    = "Tracking" # TODO: Zusätzliche Kategorien, bsp. für verschiedene Projekte?
$CategoryColor   = 6
$MinDuration     = 15
$DurationsStart  = @(30,60,90,120)  # F1-F4
$DurationsExtend = @(30,60,90,120)  # F5-F8
$AppDir          = Join-Path $env:LOCALAPPDATA "OutlookTimeTracker"
$LastSubjectFile = Join-Path $AppDir "lastsubject.txt"
if (-not (Test-Path $AppDir)) { New-Item -ItemType Directory -Path $AppDir -Force | Out-Null }

# UI-Feintuning
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

function New-TrackingAppointment {
    param([string]$Subject,[int]$DurationMinutes = 30)
    $outlook = Get-OutlookApp
    if (-not $outlook) {
        [System.Windows.Forms.MessageBox]::Show("Outlook konnte nicht gestartet werden.","Fehler",
            [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null; return
    }
    $session = $outlook.Session; Ensure-Category -Session $session
    $now  = Get-Date; $mins = [math]::Max($MinDuration, [int]$DurationMinutes)
    $finalSubject = if ([string]::IsNullOrWhiteSpace($Subject)) { $CategoryName } else { $Subject.Trim() }
    $appt = $outlook.CreateItem(1)
    $appt.Subject=$finalSubject; $appt.Start=$now; $appt.End=$now.AddMinutes($mins)
    # Markiert den Eintrag als "Privat", TODO: Switch einbauen
    # $appt.Categories=$CategoryName; $appt.BusyStatus=2; $appt.Sensitivity=2; $appt.ReminderSet=$false
    $appt.Categories=$CategoryName; $appt.BusyStatus=2; $appt.ReminderSet=$false
    $appt.Body = "Automatisch angelegt via Stream Deck / Script am $($now.ToString('yyyy-MM-dd HH:mm'))."
    $appt.Save()
    try { Set-Content -Path $LastSubjectFile -Value $finalSubject -Encoding UTF8 -Force } catch {}
}

function Get-CurrentAppointment {
    param([bool]$PreferTracking = $true)
    $outlook = Get-OutlookApp; if (-not $outlook) { return $null }
    $ns=$outlook.Session; $cal=$ns.GetDefaultFolder(9); $items=$cal.Items
    $items.IncludeRecurrences=$true; $items.Sort("[Start]")
    $stamp=(Get-Date).ToString("MM/dd/yyyy HH:mm")
    $current=$items.Restrict("[Start] <= '$stamp' AND [End] > '$stamp'")
    if ($current -and $current.Count -gt 0) {
        if ($PreferTracking) {
            foreach ($it in $current) {
                try {
                    if ($it.MessageClass -like "IPM.Appointment*" -and
                        ($it.Categories -and ($it.Categories -split ';' | % { $_.Trim() }) -contains $CategoryName)) { return $it }
                } catch {}
            }
        }
        foreach ($it in $current) { try { if ($it.MessageClass -like "IPM.Appointment*") { return $it } } catch {} }
    }
    $null
}

function Extend-CurrentAppointment {
    param([int]$AddMinutes, [switch]$Silent)
    $appt = Get-CurrentAppointment -PreferTracking:$true
    if (-not $appt) {
        if (-not $Silent) { [System.Windows.Forms.MessageBox]::Show("Kein laufender Termin gefunden.","Nichts zu verlängern",
            [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null }
        return
    }
    $appt.End = $appt.End.AddMinutes([int]$AddMinutes)
    $append = "`r`n[+$AddMinutes min am $((Get-Date).ToString('yyyy-MM-dd HH:mm'))]"
    try { if ([string]::IsNullOrEmpty($appt.Body)) { $appt.Body=$append.Trim() } else { $appt.Body += $append } } catch {}
    $appt.Save()
    if (-not $Silent) {
        [System.Windows.Forms.MessageBox]::Show("Termin bis $($appt.End.ToString('HH:mm')) verlängert.","Fortgesetzt (+$AddMinutes)",
            [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    }
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

# Haupt-Table
$root = New-Object System.Windows.Forms.TableLayoutPanel
$root.Dock = 'Fill'
$root.ColumnCount = 1
$root.RowCount    = 5
$root.Padding     = New-Object System.Windows.Forms.Padding($Gap,$Gap,$Gap,$Gap)
$root.BackColor   = $ClrBg
$root.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) ) | Out-Null
$root.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) ) | Out-Null
$root.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) ) | Out-Null
$root.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) ) | Out-Null
$root.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) ) | Out-Null

# Überschrift + Text
$labelTask = New-Object System.Windows.Forms.Label
$labelTask.Text = "Aufgabe / Thema"
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
try { if (Test-Path $LastSubjectFile) { $last = Get-Content $LastSubjectFile -EA SilentlyContinue | Select-Object -First 1; if ($last){ $textTask.Text = $last } } } catch {}

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

# Karten
$cardStart,  $panelStart  = New-Card "Neu starten (setzt Teams auf 'Beschäftigt')" $ClrCard
$cardExtend, $panelExtend = New-Card "Laufenden Termin fortsetzen" $ClrCard2

# ToolTips
$tt = New-Object System.Windows.Forms.ToolTip
$tt.AutoPopDelay=6000; $tt.InitialDelay=250; $tt.ReshowDelay=100

# Start-Buttons
$startButtons=@(); $accents=@($ClrA1,$ClrA2,$ClrA3,$ClrA4)
for ($i=0; $i -lt $DurationsStart.Count; $i++) {
    $mins=$DurationsStart[$i]; $base=$accents[$i]
    $btn = New-NiceButton "⏱  $mins Min  (F$($i+1))" $mins $base $ClrBg (Lighten $base 18)
    $btn.Add_Click([System.EventHandler]{ param($sender,$e)
        New-TrackingAppointment -Subject $textTask.Text -DurationMinutes ([int]$sender.Tag)
        $form.Close()
    })
    $tt.SetToolTip($btn, "Neuen $mins-Minuten-Block starten (Teams = Beschäftigt)")
    $panelStart.Controls.Add($btn) | Out-Null
    $startButtons += $btn
}
$form.AcceptButton = $startButtons[0]

# Extend-Buttons
$extendButtons=@()
for ($i=0; $i -lt $DurationsExtend.Count; $i++) {
    $mins=$DurationsExtend[$i]
    $btn = New-NiceButton "＋  +$mins Min  (F$($i+5))" $mins $ClrPlus $ClrBg (Lighten $ClrPlus 18)
    $btn.Add_Click([System.EventHandler]{ param($sender,$e) Extend-CurrentAppointment -AddMinutes ([int]$sender.Tag) })
    $tt.SetToolTip($btn, "Laufenden Termin um $mins Minuten verlängern")
    $panelExtend.Controls.Add($btn) | Out-Null
    $extendButtons += $btn
}

# Bottom Row: Abbrechen rechts
$bottomRow = New-Object System.Windows.Forms.TableLayoutPanel
$bottomRow.ColumnCount=1; $bottomRow.RowCount=1
$bottomRow.Dock='Top'; $bottomRow.AutoSize=$true
$bottomRow.Margin = New-Object System.Windows.Forms.Padding(0,0,0,0)
$bottomRow.Anchor = 'Left,Right'

$cancelHost = New-Object System.Windows.Forms.FlowLayoutPanel
$cancelHost.FlowDirection='RightToLeft'
$cancelHost.WrapContents=$false
$cancelHost.AutoSize=$true
$cancelHost.Dock='Top'
$cancelHost.Margin = New-Object System.Windows.Forms.Padding(0,0,0,0)

$btnCancel = New-NiceButton "✖  Abbrechen  (Esc)" 0 $ClrCancel $ClrBg (Lighten $ClrCancel 18)
$btnCancel.Add_Click({ $form.Close() })
$cancelHost.Controls.Add($btnCancel) | Out-Null
$bottomRow.Controls.Add($cancelHost)

# Keybindings
$form.Add_KeyDown({
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

# Zusammenbau
$root.Controls.Add($labelTask,  0,0)
$root.Controls.Add($textTask,   0,1)
$root.Controls.Add($cardStart,  0,2)
$root.Controls.Add($cardExtend, 0,3)
$root.Controls.Add($bottomRow,  0,4)
$form.Controls.Add($root)

# rechte Flucht synchron halten
$form.Add_Resize({
    $textTask.Width = $root.ClientSize.Width - $root.Padding.Left - $root.Padding.Right
    $cardStart.Width = $textTask.Width
    $cardExtend.Width = $textTask.Width
})

# --- Fokus/Foreground zuverlässig setzen (Variante A) ---
$form.Add_Shown({
    $form.TopMost = $true
    [System.Windows.Forms.Application]::DoEvents()
    Invoke-ForegroundFix -Form $form -FocusControl $textTask
})

# Dialog
[void]$form.ShowDialog()
