Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir
$envFile = Join-Path $scriptDir '.env'

function Read-EnvFile {
    $h = @{}
    if (Test-Path $envFile) {
        Get-Content $envFile | ForEach-Object {
            if ($_ -match '^([^#=\s][^=]*)=(.*)$') { $h[$matches[1].Trim()] = $matches[2].Trim() }
        }
    }
    return $h
}

function Save-JiraEnv($url, $email, $token, $project, $accountId, $excelFile, $sprintAnchor) {
    $keys = @('JIRA_URL','JIRA_EMAIL','JIRA_API_TOKEN','JIRA_PROJECT','JIRA_ACCOUNT_ID','EXCEL_FILE','SPRINT_ANCHOR')
    $vals = @{ JIRA_URL=$url; JIRA_EMAIL=$email; JIRA_API_TOKEN=$token; JIRA_PROJECT=$project; JIRA_ACCOUNT_ID=$accountId; EXCEL_FILE=$excelFile; SPRINT_ANCHOR=$sprintAnchor }
    $lines = @(); $written = @{}
    if (Test-Path $envFile) {
        Get-Content $envFile | ForEach-Object {
            $m = $false
            foreach ($k in $keys) {
                if ($_ -match "^$k\s*=") { $lines += "$k=$($vals[$k])"; $written[$k]=$true; $m=$true; break }
            }
            if (-not $m) { $lines += $_ }
        }
    }
    foreach ($k in $keys) { if (-not $written[$k]) { $lines += "$k=$($vals[$k])" } }
    $lines | Set-Content $envFile -Encoding utf8
}

$envData = Read-EnvFile

# ============================================================
# Main form
# ============================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Jira -> Timesheet Sync'
$form.Size = New-Object System.Drawing.Size(1060, 680)
$form.MinimumSize = New-Object System.Drawing.Size(1060, 400)
$form.MaximumSize = New-Object System.Drawing.Size(1060, 2000)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$form.BackColor = [System.Drawing.Color]::FromArgb(242, 242, 245)

function New-Btn($parent, $text, $x, $y, $w, $h, $r, $g, $b) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $text; $btn.Location = New-Object System.Drawing.Point($x,$y)
    $btn.Size = New-Object System.Drawing.Size($w,$h)
    $btn.BackColor = [System.Drawing.Color]::FromArgb(255,[int]$r,[int]$g,[int]$b)
    $btn.ForeColor = [System.Drawing.Color]::White
    $btn.FlatStyle = 'Flat'; $btn.FlatAppearance.BorderSize = 0
    $btn.Font = New-Object System.Drawing.Font('Segoe UI',9,[System.Drawing.FontStyle]::Bold)
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    [void]$parent.Controls.Add($btn); return $btn
}

# ---- Top bar: Settings + connection dot + Excel picker --------------------
$pnlTop = New-Object System.Windows.Forms.Panel
$pnlTop.Location = New-Object System.Drawing.Point(0,0)
$pnlTop.Size = New-Object System.Drawing.Size(1060,46)
$pnlTop.BackColor = [System.Drawing.Color]::FromArgb(255,50,50,60)
[void]$form.Controls.Add($pnlTop)

$btnSettings = New-Btn $pnlTop 'Settings' 8 8 90 30 70 70 85
$btnSettings.Font = New-Object System.Drawing.Font('Segoe UI',9)

$lblDot = New-Object System.Windows.Forms.Label
$lblDot.Text = 'l'; $lblDot.ForeColor = [System.Drawing.Color]::Gray
$lblDot.Font = New-Object System.Drawing.Font('Segoe UI',18,[System.Drawing.FontStyle]::Bold)
$lblDot.Location = New-Object System.Drawing.Point(106,6)
$lblDot.Size = New-Object System.Drawing.Size(22,30)
$lblDot.TextAlign = 'MiddleCenter'
[void]$pnlTop.Controls.Add($lblDot)

$lblConnStatus = New-Object System.Windows.Forms.Label
$lblConnStatus.Text = 'not tested'
$lblConnStatus.ForeColor = [System.Drawing.Color]::FromArgb(255,180,180,180)
$lblConnStatus.Font = New-Object System.Drawing.Font('Segoe UI',9)
$lblConnStatus.Location = New-Object System.Drawing.Point(130,15)
$lblConnStatus.Size = New-Object System.Drawing.Size(160,18)
[void]$pnlTop.Controls.Add($lblConnStatus)

# Excel path lives in Settings only (persisted to .env as EXCEL_FILE). We keep
# $txtFile as an off-screen holder so the rest of the UI (Read/Submit/Fill) can
# still read $txtFile.Text without threading the path through every handler.
$txtFile = New-Object System.Windows.Forms.TextBox
$txtFile.Text = if ($envData['EXCEL_FILE']) { $envData['EXCEL_FILE'] } else { Join-Path $scriptDir 'data\timesheet.xlsx' }

# ---- Button strip ----------------------------------------------------------
$pnlBtns = New-Object System.Windows.Forms.Panel
$pnlBtns.Location = New-Object System.Drawing.Point(0,46)
$pnlBtns.Size = New-Object System.Drawing.Size(1060,52)
$pnlBtns.BackColor = [System.Drawing.Color]::FromArgb(255,60,60,72)
[void]$form.Controls.Add($pnlBtns)

$btnTest   = New-Btn $pnlBtns 'Test'           8   8 80 36 70  70  90
$btnRead   = New-Btn $pnlBtns 'Read File'      96  8 110 36 40  100 140
$btnOpen   = New-Btn $pnlBtns 'Open File'      214 8 100 36 40  110 60

# separator (visual gap)
$sep = New-Object System.Windows.Forms.Label
$sep.Location = New-Object System.Drawing.Point(322,10)
$sep.Size = New-Object System.Drawing.Size(2,30)
$sep.BackColor = [System.Drawing.Color]::FromArgb(255,100,100,115)
[void]$pnlBtns.Controls.Add($sep)

$btnSync   = New-Btn $pnlBtns 'Sync Jira -> Excel'  332 8 170 36 0   122 200
$btnSubmit = New-Btn $pnlBtns 'Submit Tasks -> D365' 510 8 185 36 180 80  0
$btnFill   = New-Btn $pnlBtns 'Fill Times -> D365'   703 8 155 36 80  80  80
$btnFill.Enabled = $false
$btnFill.ForeColor = [System.Drawing.Color]::FromArgb(255,140,140,140)

$btnFillStd = New-Btn $pnlBtns 'Standard -> Excel' 866 8 150 36 120 80 160
$btnStop    = New-Btn $pnlBtns 'Stop' 1024 8 22 36 140 30 30
$btnStop.Text = [char]9632  # stop square
$btnStop.Enabled = $false

# ---- Weeks strip (compact, checkbox-only) ----------------------------------
$pnlWeeks = New-Object System.Windows.Forms.Panel
$pnlWeeks.Location = New-Object System.Drawing.Point(0,98)
$pnlWeeks.Size = New-Object System.Drawing.Size(1060,40)
$pnlWeeks.BackColor = [System.Drawing.Color]::White
[void]$form.Controls.Add($pnlWeeks)

$lblWeeksTitle = New-Object System.Windows.Forms.Label
$lblWeeksTitle.Text = 'Weeks:'
$lblWeeksTitle.Font = New-Object System.Drawing.Font('Segoe UI',9)
$lblWeeksTitle.Location = New-Object System.Drawing.Point(8,12)
$lblWeeksTitle.Size = New-Object System.Drawing.Size(48,18)
$lblWeeksTitle.ForeColor = [System.Drawing.Color]::FromArgb(255,80,80,100)
[void]$pnlWeeks.Controls.Add($lblWeeksTitle)

$lnkSelectAll = New-Object System.Windows.Forms.LinkLabel
$lnkSelectAll.Text = 'All'
$lnkSelectAll.Location = New-Object System.Drawing.Point(58,12)
$lnkSelectAll.Size = New-Object System.Drawing.Size(22,18)
$lnkSelectAll.Font = New-Object System.Drawing.Font('Segoe UI',9)
$lnkSelectAll.Visible = $false
[void]$pnlWeeks.Controls.Add($lnkSelectAll)

$lnkNone = New-Object System.Windows.Forms.LinkLabel
$lnkNone.Text = 'None'
$lnkNone.Location = New-Object System.Drawing.Point(82,12)
$lnkNone.Size = New-Object System.Drawing.Size(34,18)
$lnkNone.Font = New-Object System.Drawing.Font('Segoe UI',9)
$lnkNone.Visible = $false
[void]$pnlWeeks.Controls.Add($lnkNone)

$lblMonth = New-Object System.Windows.Forms.Label
$lblMonth.Text = 'Month:'
$lblMonth.Location = New-Object System.Drawing.Point(122,13)
$lblMonth.Size = New-Object System.Drawing.Size(44,18)
$lblMonth.Font = New-Object System.Drawing.Font('Segoe UI',9)
$lblMonth.ForeColor = [System.Drawing.Color]::FromArgb(255,80,80,100)
[void]$pnlWeeks.Controls.Add($lblMonth)

$cmbMonth = New-Object System.Windows.Forms.ComboBox
$cmbMonth.Location = New-Object System.Drawing.Point(166,10)
$cmbMonth.Size = New-Object System.Drawing.Size(128,22)
$cmbMonth.DropDownStyle = 'DropDownList'
$cmbMonth.Font = New-Object System.Drawing.Font('Segoe UI',9)
[void]$cmbMonth.Items.Add('All')
$cmbMonth.SelectedIndex = 0
[void]$pnlWeeks.Controls.Add($cmbMonth)

$script:monthCodes = @{}

$script:weekCheckboxes = @()

function Populate-Weeks($weeksJson) {
    $toRemove = @($pnlWeeks.Controls | Where-Object { $_ -is [System.Windows.Forms.CheckBox] })
    foreach ($c in $toRemove) { $pnlWeeks.Controls.Remove($c) }
    $script:weekCheckboxes = @()
    $lnkSelectAll.Visible = $false; $lnkNone.Visible = $false
    if (-not $weeksJson -or $weeksJson.Count -eq 0) { return }

    # Rebuild month dropdown from the months each week touches (start + end).
    $cmbMonth.Items.Clear()
    [void]$cmbMonth.Items.Add('All')
    $script:monthCodes = @{}
    $codes = New-Object 'System.Collections.Generic.SortedSet[string]'
    foreach ($w in $weeksJson) {
        [void]$codes.Add($w.start.Substring(0,7))
        [void]$codes.Add($w.end.Substring(0,7))
    }
    foreach ($code in $codes) {
        $display = [datetime]::ParseExact("$code-01",'yyyy-MM-dd',$null).ToString('MMMM yyyy',[System.Globalization.CultureInfo]::InvariantCulture)
        $script:monthCodes[$display] = $code
        [void]$cmbMonth.Items.Add($display)
    }
    $cmbMonth.SelectedIndex = 0

    # Lay out checkboxes in rows of 5; increase panel height per extra row
    $colW = 142; $startX = 306; $cols = [Math]::Floor((1060 - $startX) / $colW)
    $rowsNeeded = [Math]::Ceiling($weeksJson.Count / $cols)
    $panelH = 40 + ([Math]::Max($rowsNeeded - 1, 0) * 22)
    $pnlWeeks.Height = $panelH
    $divider.Top    = 98 + $panelH
    $pnlStatus.Top  = 98 + $panelH + 1
    $txtLog.Top     = 98 + $panelH + 29
    $txtLog.Height  = [Math]::Max($form.ClientSize.Height - $txtLog.Top, 50)

    $i = 0
    foreach ($w in $weeksJson) {
        $col = $i % $cols
        $row = [Math]::Floor($i / $cols)
        $chk = New-Object System.Windows.Forms.CheckBox
        $chk.Text = $w.start.Substring(5) + ' – ' + $w.end.Substring(5)
        $chk.Location = New-Object System.Drawing.Point(($startX + $col * $colW), (11 + $row * 22))
        $chk.Size = New-Object System.Drawing.Size($colW, 18)
        $chk.Checked = $true
        $chk.Font = New-Object System.Drawing.Font('Segoe UI',8.5)
        $chk.Tag = $w.start
        [void]$pnlWeeks.Controls.Add($chk)
        $script:weekCheckboxes += $chk
        $i++
    }
    $lnkSelectAll.Visible = $true; $lnkNone.Visible = $true
}

$lnkSelectAll.Add_LinkClicked({ foreach ($c in $script:weekCheckboxes) { $c.Checked = $true } })
$lnkNone.Add_LinkClicked({      foreach ($c in $script:weekCheckboxes) { $c.Checked = $false } })

# ---- Divider ---------------------------------------------------------------
$divider = New-Object System.Windows.Forms.Label
$divider.Location = New-Object System.Drawing.Point(0,138)
$divider.Size = New-Object System.Drawing.Size(1060,1)
$divider.BackColor = [System.Drawing.Color]::FromArgb(255,200,200,210)
[void]$form.Controls.Add($divider)

# ---- Status bar ------------------------------------------------------------
$pnlStatus = New-Object System.Windows.Forms.Panel
$pnlStatus.Location = New-Object System.Drawing.Point(0,139)
$pnlStatus.Size = New-Object System.Drawing.Size(1060,28)
$pnlStatus.BackColor = [System.Drawing.Color]::FromArgb(255,235,235,240)
[void]$form.Controls.Add($pnlStatus)

$lblStatusDot = New-Object System.Windows.Forms.Label
$lblStatusDot.Text = 'l'
$lblStatusDot.Font = New-Object System.Drawing.Font('Segoe UI',16,[System.Drawing.FontStyle]::Bold)
$lblStatusDot.Location = New-Object System.Drawing.Point(8,-4)
$lblStatusDot.Size = New-Object System.Drawing.Size(20,30)
$lblStatusDot.ForeColor = [System.Drawing.Color]::LimeGreen
[void]$pnlStatus.Controls.Add($lblStatusDot)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = 'Ready'
$lblStatus.Location = New-Object System.Drawing.Point(30,6)
$lblStatus.Size = New-Object System.Drawing.Size(710,18)
$lblStatus.Font = New-Object System.Drawing.Font('Segoe UI',9)
[void]$pnlStatus.Controls.Add($lblStatus)

$btnCopyLog = New-Object System.Windows.Forms.Button
$btnCopyLog.Text = 'Copy log'
$btnCopyLog.Location = New-Object System.Drawing.Point(728, 3)
$btnCopyLog.Size = New-Object System.Drawing.Size(164, 22)
$btnCopyLog.FlatStyle = 'Flat'
$btnCopyLog.FlatAppearance.BorderSize = 1
$btnCopyLog.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(255,200,200,210)
$btnCopyLog.BackColor = [System.Drawing.Color]::FromArgb(255,245,245,250)
$btnCopyLog.ForeColor = [System.Drawing.Color]::FromArgb(255,80,80,100)
$btnCopyLog.Font = New-Object System.Drawing.Font('Segoe UI',8.5)
$btnCopyLog.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnCopyLog.Add_Click({
    if ([string]::IsNullOrEmpty($txtLog.Text)) { return }
    [System.Windows.Forms.Clipboard]::SetText($txtLog.Text)
    $prevText = $btnCopyLog.Text
    $btnCopyLog.Text = 'Copied!'
    $btnCopyLog.ForeColor = [System.Drawing.Color]::DarkGreen
    $script:copyLogTimer = New-Object System.Windows.Forms.Timer
    $script:copyLogTimer.Interval = 1500
    $script:copyLogTimer.Add_Tick({
        $btnCopyLog.Text = $prevText
        $btnCopyLog.ForeColor = [System.Drawing.Color]::FromArgb(255,80,80,100)
        $script:copyLogTimer.Stop(); $script:copyLogTimer.Dispose()
    })
    $script:copyLogTimer.Start()
})
[void]$pnlStatus.Controls.Add($btnCopyLog)

# ---- Log -------------------------------------------------------------------
$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location = New-Object System.Drawing.Point(0,167)
$txtLog.Size = New-Object System.Drawing.Size(900,472)
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.Color]::FromArgb(255,20,20,26)
$txtLog.ForeColor = [System.Drawing.Color]::FromArgb(255,200,200,200)
$txtLog.Font = New-Object System.Drawing.Font('Consolas',9)
$txtLog.BorderStyle = 'None'
$txtLog.ScrollBars = 'Vertical'
[void]$form.Controls.Add($txtLog)

$form.Add_Resize({
    $h = $form.ClientSize.Height
    $txtLog.Height = [Math]::Max($h - $txtLog.Top, 50)
})

# ============================================================
# Settings dialog
# ============================================================
function Show-Settings {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Settings'
    $dlg.Size = New-Object System.Drawing.Size(560,470)
    $dlg.StartPosition = 'CenterParent'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font('Segoe UI',9)
    $dlg.BackColor = [System.Drawing.Color]::FromArgb(255,248,248,252)

    function Add-Row($parent, $label, $y, $pw = $false) {
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $label; $lbl.Location = New-Object System.Drawing.Point(16,$($y+3))
        $lbl.Size = New-Object System.Drawing.Size(90,20); $lbl.TextAlign = 'MiddleRight'
        [void]$parent.Controls.Add($lbl)
        $txt = New-Object System.Windows.Forms.TextBox
        $txt.Location = New-Object System.Drawing.Point(114,$y)
        $txt.Size = New-Object System.Drawing.Size(390,24)
        if ($pw) { $txt.UseSystemPasswordChar = $true }
        [void]$parent.Controls.Add($txt)
        return $txt
    }

    $d = Read-EnvFile
    $tUrl  = Add-Row $dlg 'Jira URL:'    20
    $tMail = Add-Row $dlg 'Email:'       58
    $tTok  = Add-Row $dlg 'API Token:'   96 $true
    $tProj = Add-Row $dlg 'Project:'    134
    $tAcct = Add-Row $dlg 'Account ID:' 172

    $tUrl.Text  = if ($d['JIRA_URL'])        { $d['JIRA_URL'] }        else { 'https://amcbridge.atlassian.net' }
    $tMail.Text = if ($d['JIRA_EMAIL'])       { $d['JIRA_EMAIL'] }       else { '' }
    $tTok.Text  = if ($d['JIRA_API_TOKEN'])   { $d['JIRA_API_TOKEN'] }   else { '' }
    $tProj.Text = if ($d['JIRA_PROJECT'])     { $d['JIRA_PROJECT'] }     else { 'GT2' }
    $tAcct.Text = if ($d['JIRA_ACCOUNT_ID'])  { $d['JIRA_ACCOUNT_ID'] }  else { '' }

    # Hint: which Jira user the activity queries (2,3,4,5,6) filter on. Empty = token owner.
    $lblAcctHint = New-Object System.Windows.Forms.Label
    $lblAcctHint.Text = 'Leave empty to use the API-token owner (/myself)'
    $lblAcctHint.Location = New-Object System.Drawing.Point(114,197)
    $lblAcctHint.Size = New-Object System.Drawing.Size(390,16)
    $lblAcctHint.ForeColor = [System.Drawing.Color]::Gray
    $lblAcctHint.Font = New-Object System.Drawing.Font('Segoe UI',8)
    [void]$dlg.Controls.Add($lblAcctHint)

    $tAnchor = Add-Row $dlg 'Sprint end:' 219
    $tAnchor.Text = if ($d['SPRINT_ANCHOR']) { $d['SPRINT_ANCHOR'] } else { '' }
    $lblAnchorHint = New-Object System.Windows.Forms.Label
    $lblAnchorHint.Text = 'Any sprint-end Friday (YYYY-MM-DD) - 2-week cycles counted from here'
    $lblAnchorHint.Location = New-Object System.Drawing.Point(114,244)
    $lblAnchorHint.Size = New-Object System.Drawing.Size(390,16)
    $lblAnchorHint.ForeColor = [System.Drawing.Color]::Gray
    $lblAnchorHint.Font = New-Object System.Drawing.Font('Segoe UI',8)
    [void]$dlg.Controls.Add($lblAnchorHint)

    # ---- Excel file row ----
    $lblEx = New-Object System.Windows.Forms.Label
    $lblEx.Text = 'Excel file:'; $lblEx.Location = New-Object System.Drawing.Point(16,269)
    $lblEx.Size = New-Object System.Drawing.Size(90,20); $lblEx.TextAlign = 'MiddleRight'
    [void]$dlg.Controls.Add($lblEx)
    $tExcel = New-Object System.Windows.Forms.TextBox
    $tExcel.Location = New-Object System.Drawing.Point(114,266)
    $tExcel.Size = New-Object System.Drawing.Size(354,24)
    $tExcel.Text = $txtFile.Text
    [void]$dlg.Controls.Add($tExcel)
    $btnExBrowse = New-Object System.Windows.Forms.Button
    $btnExBrowse.Text = '...'; $btnExBrowse.Location = New-Object System.Drawing.Point(474,266)
    $btnExBrowse.Size = New-Object System.Drawing.Size(30,24); $btnExBrowse.FlatStyle = 'Flat'
    $btnExBrowse.Add_Click({
        $fd = New-Object System.Windows.Forms.OpenFileDialog
        $fd.Filter = 'Excel files (*.xlsx)|*.xlsx'
        $fd.InitialDirectory = Split-Path $tExcel.Text -Parent
        if ($fd.ShowDialog() -eq 'OK') { $tExcel.Text = $fd.FileName }
    })
    [void]$dlg.Controls.Add($btnExBrowse)

    $chkSh = New-Object System.Windows.Forms.CheckBox
    $chkSh.Text = 'Show token'; $chkSh.Location = New-Object System.Drawing.Point(114,301)
    $chkSh.Size = New-Object System.Drawing.Size(100,22)
    $chkSh.Add_CheckedChanged({ $tTok.UseSystemPasswordChar = -not $chkSh.Checked })
    [void]$dlg.Controls.Add($chkSh)

    $lblTest = New-Object System.Windows.Forms.Label
    $lblTest.Location = New-Object System.Drawing.Point(16,304)
    $lblTest.Size = New-Object System.Drawing.Size(500,20)
    $lblTest.ForeColor = [System.Drawing.Color]::Gray
    [void]$dlg.Controls.Add($lblTest)

    $btnT = New-Object System.Windows.Forms.Button
    $btnT.Text = 'Test Connection'
    $btnT.Location = New-Object System.Drawing.Point(16,363)
    $btnT.Size = New-Object System.Drawing.Size(140,32)
    $btnT.Add_Click({
        $lblTest.Text = 'Testing...'; $lblTest.ForeColor = [System.Drawing.Color]::DodgerBlue
        $dlg.Refresh()
        Save-JiraEnv $tUrl.Text $tMail.Text $tTok.Text $tProj.Text $tAcct.Text $tExcel.Text $tAnchor.Text
        $psi2 = New-Object System.Diagnostics.ProcessStartInfo
        $psi2.FileName = 'python'; $psi2.Arguments = "scripts\jira-sync.py --command test --file `"$($tExcel.Text)`""
        $psi2.WorkingDirectory = $scriptDir; $psi2.UseShellExecute = $false
        $psi2.RedirectStandardOutput = $true; $psi2.RedirectStandardError = $true
        $psi2.CreateNoWindow = $true
        $psi2.StandardOutputEncoding = [System.Text.Encoding]::UTF8
        $psi2.StandardErrorEncoding  = [System.Text.Encoding]::UTF8
        $p2 = [System.Diagnostics.Process]::Start($psi2)
        $out = $p2.StandardOutput.ReadToEnd(); $p2.WaitForExit()
        if ($p2.ExitCode -eq 0) {
            $lblTest.Text = 'Connected OK'; $lblTest.ForeColor = [System.Drawing.Color]::DarkGreen
            $lblDot.ForeColor = [System.Drawing.Color]::LimeGreen
            $lblConnStatus.Text = 'Connected'
            $lblConnStatus.ForeColor = [System.Drawing.Color]::FromArgb(255,100,220,100)
        } else {
            $lblTest.Text = 'Connection failed - check credentials'; $lblTest.ForeColor = [System.Drawing.Color]::Red
            $lblDot.ForeColor = [System.Drawing.Color]::OrangeRed
            $lblConnStatus.Text = 'Not connected'
            $lblConnStatus.ForeColor = [System.Drawing.Color]::OrangeRed
        }
    })
    [void]$dlg.Controls.Add($btnT)

    $btnSv = New-Object System.Windows.Forms.Button
    $btnSv.Text = 'Save & Close'; $btnSv.DialogResult = 'OK'
    $btnSv.Location = New-Object System.Drawing.Point(410,363)
    $btnSv.Size = New-Object System.Drawing.Size(120,32)
    $btnSv.BackColor = [System.Drawing.Color]::FromArgb(255,0,122,200)
    $btnSv.ForeColor = [System.Drawing.Color]::White; $btnSv.FlatStyle = 'Flat'
    $btnSv.Add_Click({
        Save-JiraEnv $tUrl.Text $tMail.Text $tTok.Text $tProj.Text $tAcct.Text $tExcel.Text $tAnchor.Text
        $txtFile.Text = $tExcel.Text
    })
    [void]$dlg.Controls.Add($btnSv)

    [void]$dlg.ShowDialog($form)
}

$btnSettings.Add_Click({ Show-Settings })

# ============================================================
# Process runner
# ============================================================
$script:proc = $null
$script:pollTimer = $null

function Set-Status($text, $color) {
    $lblStatus.Text = $text; $lblStatusDot.ForeColor = $color
}

function Append-Log($msg) {
    if     ($msg -match '^\[ERR')            { $txtLog.SelectionColor = [System.Drawing.Color]::FromArgb(255,255,90,90) }
    elseif ($msg -match '^\[OK\]')           { $txtLog.SelectionColor = [System.Drawing.Color]::FromArgb(255,90,220,90) }
    elseif ($msg -match '^\[WARN')           { $txtLog.SelectionColor = [System.Drawing.Color]::FromArgb(255,255,200,60) }
    elseif ($msg -match '^\[INFO')           { $txtLog.SelectionColor = [System.Drawing.Color]::FromArgb(255,110,170,255) }
    elseif ($msg -match '^\[SKIP')           { $txtLog.SelectionColor = [System.Drawing.Color]::FromArgb(255,160,160,160) }
    elseif ($msg -match '^\s+\[\d+/6\]')    { $txtLog.SelectionColor = [System.Drawing.Color]::FromArgb(255,160,200,255) }
    elseif ($msg.Contains([char]0x2192))     { $txtLog.SelectionColor = [System.Drawing.Color]::FromArgb(255,140,230,140) }
    elseif ($msg -match '^\s+\+\s')         { $txtLog.SelectionColor = [System.Drawing.Color]::FromArgb(255,160,255,160) }
    else                                     { $txtLog.SelectionColor = [System.Drawing.Color]::FromArgb(255,200,200,200) }
    $txtLog.AppendText($msg + [System.Environment]::NewLine)
    $txtLog.ScrollToCaret()
}

function Start-PyProc($args_, $onDone, $exe = 'python') {
    if ($script:proc -and -not $script:proc.HasExited) { return }
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $exe; $psi.Arguments = $args_
    $psi.WorkingDirectory = $scriptDir; $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
    $psi.StandardErrorEncoding  = [System.Text.Encoding]::UTF8
    $script:proc = [System.Diagnostics.Process]::Start($psi)
    $script:onDone = $onDone
    $btnStop.Enabled = $true

    # Thread-safe queue: background runspaces enqueue lines; UI timer drains on the main thread.
    # StreamReader.Peek() blocks on Windows named pipes when no data is available — running
    # ReadLine() on a background runspace avoids freezing the WinForms message loop.
    $script:pyQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
    $qRef = $script:pyQueue

    $rsOut = [powershell]::Create()
    [void]$rsOut.AddScript({
        param($reader, $queue)
        try {
            $line = $reader.ReadLine()
            while ($null -ne $line) { $queue.Enqueue($line); $line = $reader.ReadLine() }
        } catch {}
    }).AddArgument($script:proc.StandardOutput).AddArgument($qRef)
    $rsOutHandle = $rsOut.BeginInvoke()

    $rsErr = [powershell]::Create()
    [void]$rsErr.AddScript({
        param($reader, $queue)
        try {
            $line = $reader.ReadLine()
            while ($null -ne $line) { $queue.Enqueue('[ERR] ' + $line); $line = $reader.ReadLine() }
        } catch {}
    }).AddArgument($script:proc.StandardError).AddArgument($qRef)
    $rsErrHandle = $rsErr.BeginInvoke()

    $rsWatch = [powershell]::Create()
    [void]$rsWatch.AddScript({
        param($p, $queue, $ro, $roh, $re, $reh)
        try { $p.WaitForExit() }    catch {}
        try { $ro.EndInvoke($roh) } catch {}
        try { $re.EndInvoke($reh) } catch {}
        $code = try { $p.ExitCode } catch { 1 }
        $queue.Enqueue("__EXIT__:$code")
    })
    [void]$rsWatch.AddArgument($script:proc)
    [void]$rsWatch.AddArgument($qRef)
    [void]$rsWatch.AddArgument($rsOut)
    [void]$rsWatch.AddArgument($rsOutHandle)
    [void]$rsWatch.AddArgument($rsErr)
    [void]$rsWatch.AddArgument($rsErrHandle)
    [void]$rsWatch.BeginInvoke()

    $t = New-Object System.Windows.Forms.Timer; $t.Interval = 100
    $t.Add_Tick({
        $line = $null
        while ($script:pyQueue.TryDequeue([ref]$line)) {
            if ($null -eq $line) { continue }
            if ($line.StartsWith('__EXIT__:')) {
                $code = [int]($line.Substring(9))
                $script:pollTimer.Stop(); $btnStop.Enabled = $false
                & $script:onDone $code
            } else {
                Append-Log $line
            }
        }
    })
    $script:pollTimer = $t; $t.Start()
}

# ---- Button actions --------------------------------------------------------
$btnOpen.Add_Click({ Start-Process $txtFile.Text })

$btnTest.Add_Click({
    $txtLog.Clear()
    Set-Status 'Testing connection...' ([System.Drawing.Color]::DodgerBlue)
    $fileArg = if ($txtFile.Text) { '--file "' + $txtFile.Text + '"' } else { '' }
    Start-PyProc "scripts\jira-sync.py --command test $fileArg" {
        param($code)
        if ($code -eq 0) {
            Set-Status 'Connection OK' ([System.Drawing.Color]::LimeGreen)
            $lblDot.ForeColor = [System.Drawing.Color]::LimeGreen
            $lblConnStatus.Text = 'Connected'; $lblConnStatus.ForeColor = [System.Drawing.Color]::FromArgb(255,100,220,100)
        } else {
            Set-Status 'Connection failed' ([System.Drawing.Color]::OrangeRed)
            $lblDot.ForeColor = [System.Drawing.Color]::OrangeRed
            $lblConnStatus.Text = 'Not connected'; $lblConnStatus.ForeColor = [System.Drawing.Color]::OrangeRed
        }
    }
})

$btnRead.Add_Click({
    if (-not (Test-Path $txtFile.Text)) {
        [void][System.Windows.Forms.MessageBox]::Show('Excel file not found: ' + $txtFile.Text, 'Error')
        return
    }
    $txtLog.Clear()
    Set-Status 'Reading file...' ([System.Drawing.Color]::DodgerBlue)
    $psiR = New-Object System.Diagnostics.ProcessStartInfo
    $psiR.FileName = 'python'
    $psiR.Arguments = 'scripts\jira-sync.py --command read-weeks --file "' + $txtFile.Text + '"'
    $psiR.WorkingDirectory = $scriptDir; $psiR.UseShellExecute = $false
    $psiR.RedirectStandardOutput = $true; $psiR.RedirectStandardError = $true; $psiR.CreateNoWindow = $true
    $psiR.StandardOutputEncoding = [System.Text.Encoding]::UTF8
    $psiR.StandardErrorEncoding  = [System.Text.Encoding]::UTF8
    $pR = [System.Diagnostics.Process]::Start($psiR)
    $out = $pR.StandardOutput.ReadToEnd(); $pR.WaitForExit()
    if ($pR.ExitCode -eq 0) {
        try {
            $weeks = $out | ConvertFrom-Json
            Populate-Weeks $weeks
            Set-Status ("File loaded: " + $weeks.Count + " week(s) found") ([System.Drawing.Color]::LimeGreen)
            Append-Log ('[INFO] Loaded ' + $weeks.Count + ' week(s) from: ' + $txtFile.Text)
            foreach ($w in $weeks) {
                Append-Log ('')
                Append-Log ("[INFO] $($w.start)  –  $($w.end)  ($($w.tasks.Count) task(s))")
                if ($w.tasks.Count -eq 0) {
                    Append-Log ('    (no tasks yet)')
                } else {
                    foreach ($task in $w.tasks) {
                        Append-Log ("    $([char]0x2022) $task")
                    }
                }
            }
        } catch {
            Set-Status 'Failed to parse weeks' ([System.Drawing.Color]::OrangeRed)
            Append-Log ('[ERR] ' + $out)
        }
    } else {
        Set-Status 'Read failed' ([System.Drawing.Color]::OrangeRed)
        Append-Log ('[ERR] ' + $pR.StandardError.ReadToEnd())
    }
})

$btnSync.Add_Click({
    $monthCode = $null
    if ($cmbMonth.SelectedIndex -gt 0) { $monthCode = $script:monthCodes[[string]$cmbMonth.SelectedItem] }

    $selected = @($script:weekCheckboxes | Where-Object { $_.Checked } | ForEach-Object { $_.Tag })
    if (-not $monthCode -and $script:weekCheckboxes.Count -gt 0 -and $selected.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show('No weeks selected.', 'Sync')
        return
    }
    $txtLog.Clear()
    Set-Status 'Syncing from Jira...' ([System.Drawing.Color]::DodgerBlue)

    # A chosen month drives week selection itself — the week checkboxes are ignored for that run.
    $extraArgs = ''
    if ($monthCode) {
        $extraArgs = ' --month ' + $monthCode
    } elseif ($selected.Count -gt 0) {
        $extraArgs = ' --weeks ' + ($selected -join ' ')
    }
    Start-PyProc ('scripts\jira-sync.py --command sync --file "' + $txtFile.Text + '"' + $extraArgs) {
        param($code)
        if ($code -eq 0) { Set-Status 'Sync complete' ([System.Drawing.Color]::LimeGreen) }
        else              { Set-Status ('Sync failed (exit '+$code+')') ([System.Drawing.Color]::OrangeRed) }
    }
})

$btnFillStd.Add_Click({
    if (-not (Test-Path $txtFile.Text)) {
        [void][System.Windows.Forms.MessageBox]::Show('Excel file not found: ' + $txtFile.Text, 'Standard')
        return
    }
    $monthCode = $null
    if ($cmbMonth.SelectedIndex -gt 0) { $monthCode = $script:monthCodes[[string]$cmbMonth.SelectedItem] }
    $selected = @($script:weekCheckboxes | Where-Object { $_.Checked } | ForEach-Object { $_.Tag })
    if (-not $monthCode -and $selected.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show('Select a month or check at least one week.', 'Standard')
        return
    }
    $txtLog.Clear()
    Set-Status 'Filling standard tasks...' ([System.Drawing.Color]::DodgerBlue)
    $extraArgs = ''
    if ($monthCode)            { $extraArgs += ' --month ' + $monthCode }
    if ($selected.Count -gt 0) { $extraArgs += ' --weeks ' + ($selected -join ' ') }
    Start-PyProc ('scripts\jira-sync.py --command fill-standard --file "' + $txtFile.Text + '"' + $extraArgs) {
        param($code)
        if ($code -eq 0) { Set-Status 'Standard tasks filled' ([System.Drawing.Color]::LimeGreen) }
        else             { Set-Status ('Fill failed (exit ' + $code + ')') ([System.Drawing.Color]::OrangeRed) }
    }
})

$btnSubmit.Add_Click({
    if (-not (Test-Path $txtFile.Text)) {
        [System.Windows.Forms.MessageBox]::Show('Excel file not found: ' + $txtFile.Text, 'Error', 'OK', 'Warning') | Out-Null
        return
    }
    # @() ensures result is always an array even when pipeline yields a single item.
    # Without @(), a single string is returned as a scalar and $selected[0] gives
    # the first CHARACTER instead of the first element — e.g. "2026-06-01"[0] = "2".
    $selected = @($script:weekCheckboxes | Where-Object { $_.Checked } | ForEach-Object { $_.Tag })
    if ($script:weekCheckboxes.Count -gt 0 -and $selected.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show('No weeks selected. Check at least one week.', 'D365 Submit', 'OK', 'Warning')
        return
    }
    # D365 CLI supports one week at a time (--week <YYYY-MM-DD>).
    # If a specific subset is selected, use the first checked week.
    # For remaining weeks, re-run with the next week checked.
    $weekArg = ''
    if ($selected.Count -gt 0 -and $selected.Count -lt $script:weekCheckboxes.Count) {
        $weekArg = ' --week ' + $selected[0]
        if ($selected.Count -gt 1) {
            $msg = "$($selected.Count) weeks selected. D365 submit processes one week at a time.`n`nWill submit: $($selected[0])`nRe-run for remaining week(s)."
            [void][System.Windows.Forms.MessageBox]::Show($msg, 'D365 Submit', 'OK', 'Information')
        }
    }
    $nodeArgs = '--require ts-node/register/transpile-only src/index.ts --file "' + $txtFile.Text + '"' + $weekArg
    $txtLog.Clear()
    Set-Status 'Running D365 automation...' ([System.Drawing.Color]::DodgerBlue)
    Start-PyProc $nodeArgs {
        param($exitCode)
        if ($exitCode -eq 0) { Set-Status 'D365 submission done' ([System.Drawing.Color]::LimeGreen) }
        else                 { Set-Status "D365 submission failed (exit $exitCode)" ([System.Drawing.Color]::OrangeRed) }
    } 'node'
})

$btnStop.Add_Click({
    if ($script:proc -and -not $script:proc.HasExited) { $script:proc.Kill() }
    if ($script:pollTimer) { $script:pollTimer.Stop() }
    $btnStop.Enabled = $false
    Set-Status 'Stopped' ([System.Drawing.Color]::Orange)
})

$form.Add_FormClosing({
    if ($script:proc -and -not $script:proc.HasExited) { $script:proc.Kill() }
})

[System.Windows.Forms.Application]::Run($form)