Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

# ---- Form ----------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = 'D365 Time Entry Automation'
$form.Size = New-Object System.Drawing.Size(820, 640)
$form.StartPosition = 'CenterScreen'
$form.MinimumSize = New-Object System.Drawing.Size(820, 200)
$form.MaximumSize = New-Object System.Drawing.Size(820, 2000)
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

# ---- Excel file picker --------------------------------------------
$lblFile = New-Object System.Windows.Forms.Label
$lblFile.Text = 'Excel file:'
$lblFile.Location = New-Object System.Drawing.Point(15, 20)
$lblFile.Size = New-Object System.Drawing.Size(80, 22)
$form.Controls.Add($lblFile)

$txtFile = New-Object System.Windows.Forms.TextBox
$txtFile.Location = New-Object System.Drawing.Point(100, 18)
$txtFile.Size = New-Object System.Drawing.Size(580, 22)
$txtFile.Text = Join-Path $scriptDir 'data\timesheet.xlsx'
$form.Controls.Add($txtFile)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = 'Browse...'
$btnBrowse.Location = New-Object System.Drawing.Point(690, 17)
$btnBrowse.Size = New-Object System.Drawing.Size(100, 24)
$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*'
    $dlg.InitialDirectory = Split-Path -Parent $txtFile.Text
    if ($dlg.ShowDialog() -eq 'OK') {
        $txtFile.Text = $dlg.FileName
        Show-TaskList
    }
})
$form.Controls.Add($btnBrowse)

# ---- Week filter --------------------------------------------------
$lblWeek = New-Object System.Windows.Forms.Label
$lblWeek.Text = 'Week (optional):'
$lblWeek.Location = New-Object System.Drawing.Point(15, 55)
$lblWeek.Size = New-Object System.Drawing.Size(110, 22)
$form.Controls.Add($lblWeek)

$txtWeek = New-Object System.Windows.Forms.TextBox
$txtWeek.Location = New-Object System.Drawing.Point(130, 53)
$txtWeek.Size = New-Object System.Drawing.Size(150, 22)
$txtWeek.Text = ''
$form.Controls.Add($txtWeek)

$lblWeekHint = New-Object System.Windows.Forms.Label
$lblWeekHint.Text = 'YYYY-MM-DD (Monday). Leave empty to process all weeks.'
$lblWeekHint.Location = New-Object System.Drawing.Point(290, 55)
$lblWeekHint.Size = New-Object System.Drawing.Size(490, 22)
$lblWeekHint.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($lblWeekHint)

# ---- Row 1: Test + Copy Logs --------------------------------------
$btnTest = New-Object System.Windows.Forms.Button
$btnTest.Text = 'Test Connection'
$btnTest.Location = New-Object System.Drawing.Point(15, 95)
$btnTest.Size = New-Object System.Drawing.Size(140, 38)
$btnTest.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
$btnTest.ForeColor = [System.Drawing.Color]::White
$btnTest.FlatStyle = 'Flat'
$form.Controls.Add($btnTest)

$btnSyncProfile = New-Object System.Windows.Forms.Button
$btnSyncProfile.Text = 'Sync Chrome Profile'
$btnSyncProfile.Location = New-Object System.Drawing.Point(165, 95)
$btnSyncProfile.Size = New-Object System.Drawing.Size(155, 38)
$btnSyncProfile.BackColor = [System.Drawing.Color]::FromArgb(80, 50, 120)
$btnSyncProfile.ForeColor = [System.Drawing.Color]::White
$btnSyncProfile.FlatStyle = 'Flat'
$form.Controls.Add($btnSyncProfile)

$btnLaunchDebug = New-Object System.Windows.Forms.Button
$btnLaunchDebug.Text = 'Launch Debug Chrome'
$btnLaunchDebug.Location = New-Object System.Drawing.Point(330, 95)
$btnLaunchDebug.Size = New-Object System.Drawing.Size(165, 38)
$btnLaunchDebug.BackColor = [System.Drawing.Color]::FromArgb(30, 100, 60)
$btnLaunchDebug.ForeColor = [System.Drawing.Color]::White
$btnLaunchDebug.FlatStyle = 'Flat'
$form.Controls.Add($btnLaunchDebug)

$btnCopyLogs = New-Object System.Windows.Forms.Button
$btnCopyLogs.Text = 'Copy Logs'
$btnCopyLogs.Location = New-Object System.Drawing.Point(660, 95)
$btnCopyLogs.Size = New-Object System.Drawing.Size(130, 38)
$btnCopyLogs.BackColor = [System.Drawing.Color]::FromArgb(40, 80, 120)
$btnCopyLogs.ForeColor = [System.Drawing.Color]::White
$btnCopyLogs.FlatStyle = 'Flat'
$form.Controls.Add($btnCopyLogs)

# ---- Row 2: automation buttons ------------------------------------
$btnRunAll = New-Object System.Windows.Forms.Button
$btnRunAll.Text = 'Run All (Tasks + Time Entries)'
$btnRunAll.Location = New-Object System.Drawing.Point(15, 143)
$btnRunAll.Size = New-Object System.Drawing.Size(270, 34)
$btnRunAll.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$btnRunAll.ForeColor = [System.Drawing.Color]::White
$btnRunAll.FlatStyle = 'Flat'
$form.Controls.Add($btnRunAll)

$btnStage1 = New-Object System.Windows.Forms.Button
$btnStage1.Text = 'Stage 1: Create tasks'
$btnStage1.Location = New-Object System.Drawing.Point(295, 143)
$btnStage1.Size = New-Object System.Drawing.Size(240, 34)
$form.Controls.Add($btnStage1)

$btnStage2 = New-Object System.Windows.Forms.Button
$btnStage2.Text = 'Stage 2: Submit entries'
$btnStage2.Location = New-Object System.Drawing.Point(545, 143)
$btnStage2.Size = New-Object System.Drawing.Size(245, 34)
$form.Controls.Add($btnStage2)

$btnContinue = New-Object System.Windows.Forms.Button
$btnContinue.Text = "I'm logged in - Continue"
$btnContinue.Location = New-Object System.Drawing.Point(545, 143)
$btnContinue.Size = New-Object System.Drawing.Size(245, 34)
$btnContinue.BackColor = [System.Drawing.Color]::FromArgb(40, 100, 40)
$btnContinue.ForeColor = [System.Drawing.Color]::White
$btnContinue.FlatStyle = 'Flat'
$btnContinue.Visible = $false
$form.Controls.Add($btnContinue)
$btnContinue.BringToFront()

$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Text = 'Stop'
$btnStop.Location = New-Object System.Drawing.Point(690, 143)
$btnStop.Size = New-Object System.Drawing.Size(100, 34)
$btnStop.Enabled = $false
$form.Controls.Add($btnStop)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = 'Ready'
$lblStatus.Location = New-Object System.Drawing.Point(15, 185)
$lblStatus.Size = New-Object System.Drawing.Size(775, 22)
$lblStatus.ForeColor = [System.Drawing.Color]::DarkGreen
$form.Controls.Add($lblStatus)

# ---- Task list panel (auto-sizes to content after file open) ------
$lstTasks = New-Object System.Windows.Forms.ListBox
$lstTasks.Location = New-Object System.Drawing.Point(15, 210)
$lstTasks.Size = New-Object System.Drawing.Size(775, 0)
$lstTasks.BorderStyle = 'FixedSingle'
$lstTasks.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$lstTasks.ForeColor = [System.Drawing.Color]::FromArgb(160, 210, 160)
$lstTasks.Font = New-Object System.Drawing.Font('Consolas', 9)
$lstTasks.SelectionMode = 'None'
$lstTasks.Visible = $false
$form.Controls.Add($lstTasks)

# ---- Log output ---------------------------------------------------
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(15, 210)
$txtLog.Size = New-Object System.Drawing.Size(775, 385)
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$txtLog.ForeColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
$txtLog.Font = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtLog)

# ---- Thread-safe log queue ----------------------------------------
# Register-ObjectEvent action blocks run in a separate runspace and cannot
# access script-scope variables. Instead we use a ConcurrentQueue that any
# runspace can Enqueue into, and a WinForms Timer that drains it on the UI thread.
$script:logQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
$script:currentProcess = $null
$script:isListingTasks = $false
$script:taskLinePrefix = '  ' + [char]0x2022 + ' '

$logTimer = New-Object System.Windows.Forms.Timer
$logTimer.Interval = 80
$logTimer.Add_Tick({
    $line = $null
    $anyLines = $false
    while ($script:logQueue.TryDequeue([ref]$line)) {
        if ($null -eq $line) { continue }
        if ($line.StartsWith('__EXIT__:')) {
            $code = [int]($line.Substring(9))
            if ($script:isListingTasks) {
                $script:isListingTasks = $false
                if ($lstTasks.Items.Count -gt 0) {
                    $lstTasks.Height = [Math]::Min($lstTasks.Items.Count, 10) * $lstTasks.ItemHeight + 4
                    $lstTasks.Visible = $true
                    Reposition-Log
                }
                $n = $lstTasks.Items.Count
                if ($code -eq 0) {
                    $lblStatus.Text = if ($n -gt 0) { "$n task(s) found — click Run All to submit" } else { 'No tasks found in file' }
                    $lblStatus.ForeColor = [System.Drawing.Color]::DarkGreen
                } else {
                    $lblStatus.Text = 'Failed to read tasks from file'
                    $lblStatus.ForeColor = [System.Drawing.Color]::DarkRed
                }
            } elseif ($code -eq 0) {
                $lblStatus.Text = 'Done'
                $lblStatus.ForeColor = [System.Drawing.Color]::DarkGreen
            } else {
                $lblStatus.Text = "Failed (exit $code)"
                $lblStatus.ForeColor = [System.Drawing.Color]::DarkRed
            }
            $btnContinue.Visible = $false
            Set-Running $false
            $script:currentProcess = $null
        } elseif ($line -eq '__AWAIT_LOGIN__') {
            $btnContinue.Visible = $true
            $lblStatus.Text = 'Sign in to D365 in the Chrome window, then click "I''m logged in"'
            $lblStatus.ForeColor = [System.Drawing.Color]::DarkBlue
        } elseif ($script:isListingTasks -and $line.StartsWith($script:taskLinePrefix)) {
            [void]$lstTasks.Items.Add($line.Substring($script:taskLinePrefix.Length))
        } elseif (-not $script:isListingTasks) {
            $txtLog.AppendText($line + [Environment]::NewLine)
            $anyLines = $true
        }
    }
    if ($anyLines) {
        $txtLog.SelectionStart = $txtLog.Text.Length
        $txtLog.ScrollToCaret()
    }
})
$logTimer.Start()

# ---- Helpers ------------------------------------------------------
function Append-Log {
    param([string]$text)
    $txtLog.AppendText($text + [Environment]::NewLine)
    $txtLog.SelectionStart = $txtLog.Text.Length
    $txtLog.ScrollToCaret()
}

function Set-Running {
    param([bool]$running)
    $btnRunAll.Enabled = -not $running
    $btnStage1.Enabled = -not $running
    $btnStage2.Enabled = -not $running
    $btnTest.Enabled        = -not $running
    $btnSyncProfile.Enabled  = -not $running
    $btnLaunchDebug.Enabled  = -not $running
    $btnBrowse.Enabled = -not $running
    $txtFile.Enabled   = -not $running
    $txtWeek.Enabled   = -not $running
    $btnStop.Enabled   = $running
    if (-not $running -and $lblStatus.Text -match 'Running|Testing') {
        $lblStatus.Text = 'Ready'
        $lblStatus.ForeColor = [System.Drawing.Color]::DarkGreen
    }
}

function Reposition-Log {
    $top = if ($lstTasks.Visible) { $lstTasks.Bottom + 4 } else { $lstTasks.Top }
    $h = $form.ClientSize.Height - $top - 8
    if ($h -lt 50) { $h = 50 }
    $txtLog.Location = New-Object System.Drawing.Point(15, $top)
    $txtLog.Size = New-Object System.Drawing.Size(775, $h)
}

function Start-NodeProcess {
    param([string[]]$argList, [string]$statusLabel)

    if (-not (Test-Path $txtFile.Text)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Excel file not found:`n$($txtFile.Text)", 'File not found', 'OK', 'Warning') | Out-Null
        return
    }

    # Fresh queue for each run
    $script:logQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
    $txtLog.Clear()
    $btnContinue.Visible = $false

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName  = 'node'
    $psi.Arguments = ($argList | ForEach-Object {
        if ($_ -match '\s') { '"' + $_ + '"' } else { $_ }
    }) -join ' '
    $psi.WorkingDirectory        = $scriptDir
    $psi.RedirectStandardOutput  = $true
    $psi.RedirectStandardError   = $true
    $psi.RedirectStandardInput   = $true
    $psi.UseShellExecute         = $false
    $psi.CreateNoWindow          = $true
    $psi.StandardOutputEncoding  = [System.Text.Encoding]::UTF8
    $psi.StandardErrorEncoding   = [System.Text.Encoding]::UTF8

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    [void]$proc.Start()
    $script:currentProcess = $proc

    Set-Running $true
    $lblStatus.Text     = $statusLabel
    $lblStatus.ForeColor = [System.Drawing.Color]::DarkOrange

    # Capture queue ref — runspaces receive it as a plain .NET object (no PS scope issues)
    $qRef = $script:logQueue

    # Stdout reader — runs in its own runspace, enqueues every line
    $rsOut = [powershell]::Create()
    [void]$rsOut.AddScript({
        param($reader, $queue)
        try {
            $line = $reader.ReadLine()
            while ($null -ne $line) {
                $queue.Enqueue($line)
                $line = $reader.ReadLine()
            }
        } catch {}
    }).AddArgument($proc.StandardOutput).AddArgument($qRef)
    $rsOutHandle = $rsOut.BeginInvoke()

    # Stderr reader
    $rsErr = [powershell]::Create()
    [void]$rsErr.AddScript({
        param($reader, $queue)
        try {
            $line = $reader.ReadLine()
            while ($null -ne $line) {
                $queue.Enqueue("[err] $line")
                $line = $reader.ReadLine()
            }
        } catch {}
    }).AddArgument($proc.StandardError).AddArgument($qRef)
    $rsErrHandle = $rsErr.BeginInvoke()

    # Watcher — waits for process + both readers, then pushes exit sentinel
    $rsWatch = [powershell]::Create()
    [void]$rsWatch.AddScript({
        param($p, $queue, $ro, $roh, $re, $reh)
        try { $p.WaitForExit() }      catch {}
        try { $ro.EndInvoke($roh) }   catch {}
        try { $re.EndInvoke($reh) }   catch {}
        $code = try { $p.ExitCode } catch { 1 }
        if ($code -eq 0) {
            $queue.Enqueue('')
            $queue.Enqueue('[OK] Finished successfully')
        } else {
            $queue.Enqueue("[FAIL] Exited with code $code")
        }
        $queue.Enqueue("__EXIT__:$code")
    })
    [void]$rsWatch.AddArgument($proc)
    [void]$rsWatch.AddArgument($qRef)
    [void]$rsWatch.AddArgument($rsOut)
    [void]$rsWatch.AddArgument($rsOutHandle)
    [void]$rsWatch.AddArgument($rsErr)
    [void]$rsWatch.AddArgument($rsErrHandle)
    [void]$rsWatch.BeginInvoke()
}

function Run-Automation {
    param([string]$mode)
    $argList = @('--require', 'ts-node/register', 'src/index.ts', '--file', $txtFile.Text)
    if ($txtWeek.Text) { $argList += @('--week', $txtWeek.Text) }
    if ($mode -eq 'stage1') { $argList += '--stage1-only' }
    if ($mode -eq 'stage2') { $argList += '--stage2-only' }
    Start-NodeProcess -argList $argList -statusLabel "Running ($mode)..."
}

function Run-Preflight {
    $argList = @('--require', 'ts-node/register', 'src/index.ts', '--file', $txtFile.Text, '--preflight-only')
    Start-NodeProcess -argList $argList -statusLabel 'Testing connection...'
}

function Show-TaskList {
    $lstTasks.Items.Clear()
    $lstTasks.Visible = $false
    $lstTasks.Height = 0
    Reposition-Log
    $script:isListingTasks = $true
    $argList = @('--require', 'ts-node/register', 'src/index.ts', '--file', $txtFile.Text, '--list-tasks')
    if ($txtWeek.Text) { $argList += @('--week', $txtWeek.Text) }
    Start-NodeProcess -argList $argList -statusLabel 'Reading tasks...'
}

# ---- Button wiring ------------------------------------------------
$btnCopyLogs.Add_Click({
    if ([string]::IsNullOrEmpty($txtLog.Text)) { return }
    [System.Windows.Forms.Clipboard]::SetText($txtLog.Text)
    $prevText = $btnCopyLogs.Text
    $btnCopyLogs.Text = 'Copied!'
    $resetTimer = New-Object System.Windows.Forms.Timer
    $resetTimer.Interval = 1200
    $resetTimer.Add_Tick({
        $btnCopyLogs.Text = $prevText
        $resetTimer.Stop()
        $resetTimer.Dispose()
    })
    $resetTimer.Start()
})

$btnLaunchDebug.Add_Click({
    $chromeExe = 'C:\Program Files\Google\Chrome\Application\chrome.exe'
    if (-not (Test-Path $chromeExe)) {
        $chromeExe = Join-Path $env:LOCALAPPDATA 'Google\Chrome\Application\chrome.exe'
    }
    if (-not (Test-Path $chromeExe)) {
        [System.Windows.Forms.MessageBox]::Show('chrome.exe not found', 'Error', 'OK', 'Error') | Out-Null
        return
    }

    # Kill any existing Chrome with debug port
    $existing = Get-Process -Name chrome -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -like '*remote-debugging-port*' }
    if ($existing) { $existing | Stop-Process -Force }

    $realProfile = Join-Path $env:LOCALAPPDATA 'Google\Chrome\User Data'
    Start-Process $chromeExe -ArgumentList "--remote-debugging-port=9222", "--remote-allow-origins=*", "--user-data-dir=$realProfile"
    Append-Log ''
    Append-Log '[CDP] Chrome launched with debug port 9222'
    Append-Log '[CDP] Navigate to D365, sign in, then click Run All'
    $lblStatus.Text      = 'Chrome ready - sign in then Run All'
    $lblStatus.ForeColor = [System.Drawing.Color]::DarkBlue
})

$btnSyncProfile.Add_Click({
    $result = [System.Windows.Forms.MessageBox]::Show(
        "This will copy cookies and auth data from your real Chrome to browser-profile\.`n`nPlease CLOSE Chrome completely first, then click OK.",
        'Sync Chrome Profile', 'OKCancel', 'Warning')
    if ($result -ne 'OK') { return }

    $chromeSrc = Join-Path $env:LOCALAPPDATA 'Google\Chrome\User Data'
    if (-not (Test-Path $chromeSrc)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Chrome profile not found:`n$chromeSrc", 'Error', 'OK', 'Error') | Out-Null
        return
    }

    $dst    = Join-Path $scriptDir 'browser-profile'
    $srcDef = Join-Path $chromeSrc 'Default'
    $dstDef = Join-Path $dst 'Default'
    if (-not (Test-Path $dstDef)) { New-Item -ItemType Directory -Path $dstDef | Out-Null }

    $txtLog.Clear()
    Append-Log '[Sync] Starting...'

    try {
        $ls = Join-Path $chromeSrc 'Local State'
        if (Test-Path $ls) {
            Copy-Item $ls (Join-Path $dst 'Local State') -Force
            Append-Log '[Sync] OK: Local State'
        }

        @('Cookies', 'Login Data', 'Web Data', 'Preferences', 'Secure Preferences') | ForEach-Object {
            $src = Join-Path $srcDef $_
            if (Test-Path $src) {
                Copy-Item $src (Join-Path $dstDef $_) -Force
                Append-Log "[Sync] OK: Default\$_"
            }
        }

        $netCookies = Join-Path $srcDef 'Network\Cookies'
        if (Test-Path $netCookies) {
            $dstNet = Join-Path $dstDef 'Network'
            if (-not (Test-Path $dstNet)) { New-Item -ItemType Directory -Path $dstNet | Out-Null }
            Copy-Item $netCookies (Join-Path $dstNet 'Cookies') -Force
            Append-Log '[Sync] OK: Default\Network\Cookies'
        }

        # Copy Microsoft SSO extension (required by AMC Bridge policy)
        $extId  = 'ppnbnpeolgkicgegkbkbjmhlideopiji'
        $srcExt = Join-Path $srcDef "Extensions\$extId"
        if (Test-Path $srcExt) {
            $dstExt = Join-Path $dstDef "Extensions\$extId"
            if (Test-Path $dstExt) { Remove-Item $dstExt -Recurse -Force }
            Copy-Item $srcExt $dstExt -Recurse -Force
            Append-Log '[Sync] OK: Microsoft SSO extension'
        } else {
            Append-Log '[Sync] WARN: Microsoft SSO extension not found in real Chrome profile'
        }

        Append-Log ''
        Append-Log '[Sync] Done - click Run All to start.'
        $lblStatus.Text      = 'Profile synced — ready to run'
        $lblStatus.ForeColor = [System.Drawing.Color]::DarkGreen
    } catch {
        Append-Log "[Sync] ERROR: $_"
        $lblStatus.Text      = 'Sync failed — see log'
        $lblStatus.ForeColor = [System.Drawing.Color]::DarkRed
    }
})

$btnTest.Add_Click(   { Run-Preflight })
$btnRunAll.Add_Click( { Run-Automation 'all' })
$btnStage1.Add_Click( { Run-Automation 'stage1' })
$btnStage2.Add_Click( { Run-Automation 'stage2' })

$btnContinue.Add_Click({
    if ($script:currentProcess -and -not $script:currentProcess.HasExited) {
        try {
            $script:currentProcess.StandardInput.WriteLine('CONTINUE')
            $script:currentProcess.StandardInput.Flush()
            $btnContinue.Visible = $false
            $lblStatus.Text = 'Continuing...'
            $lblStatus.ForeColor = [System.Drawing.Color]::DarkOrange
        } catch {
            $script:logQueue.Enqueue("[err] Failed to send continue signal: $_")
        }
    }
})

$btnStop.Add_Click({
    if ($script:currentProcess -and -not $script:currentProcess.HasExited) {
        try {
            $script:currentProcess.Kill($true)
            Append-Log ''
            Append-Log '[STOPPED] by user'
        } catch { Append-Log "Stop failed: $_" }
    }
})

$form.Add_FormClosing({
    $logTimer.Stop()
    if ($script:currentProcess -and -not $script:currentProcess.HasExited) {
        try { $script:currentProcess.Kill($true) } catch {}
    }
})

$form.Add_Resize({ Reposition-Log })

[void]$form.ShowDialog()
