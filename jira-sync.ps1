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

function Save-JiraEnv($url, $email, $token, $project, $accountId, $excelFile, $sprintAnchor, $sprintLengthWeeks) {
    $keys = @('JIRA_URL','JIRA_EMAIL','JIRA_API_TOKEN','JIRA_PROJECT','JIRA_ACCOUNT_ID','EXCEL_FILE','SPRINT_ANCHOR','SPRINT_LENGTH_WEEKS')
    $vals = @{ JIRA_URL=$url; JIRA_EMAIL=$email; JIRA_API_TOKEN=$token; JIRA_PROJECT=$project; JIRA_ACCOUNT_ID=$accountId; EXCEL_FILE=$excelFile; SPRINT_ANCHOR=$sprintAnchor; SPRINT_LENGTH_WEEKS=$sprintLengthWeeks }
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

$configDir = Join-Path $scriptDir 'config'

function Read-JsonConfig($name) {
    $p = Join-Path $configDir $name
    if (Test-Path $p) {
        try { return (Get-Content $p -Raw -Encoding utf8 | ConvertFrom-Json) } catch { return $null }
    }
    return $null
}

function Write-JsonConfig($name, $obj) {
    if (-not (Test-Path $configDir)) { New-Item -ItemType Directory -Path $configDir | Out-Null }
    $json = $obj | ConvertTo-Json -Depth 6
    [System.IO.File]::WriteAllText((Join-Path $configDir $name), $json, (New-Object System.Text.UTF8Encoding($false)))
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

# ---- Toolbar row 1: connection status + setup actions ---------------------
$pnlTop = New-Object System.Windows.Forms.Panel
$pnlTop.Location = New-Object System.Drawing.Point(0,0)
$pnlTop.Size = New-Object System.Drawing.Size(1060,46)
$pnlTop.BackColor = [System.Drawing.Color]::FromArgb(255,50,50,60)
[void]$form.Controls.Add($pnlTop)

$lblDot = New-Object System.Windows.Forms.Label
$lblDot.Text = 'l'; $lblDot.ForeColor = [System.Drawing.Color]::Gray
$lblDot.Font = New-Object System.Drawing.Font('Segoe UI',18,[System.Drawing.FontStyle]::Bold)
$lblDot.Location = New-Object System.Drawing.Point(8,6)
$lblDot.Size = New-Object System.Drawing.Size(22,30)
$lblDot.TextAlign = 'MiddleCenter'
[void]$pnlTop.Controls.Add($lblDot)

$lblConnStatus = New-Object System.Windows.Forms.Label
$lblConnStatus.Text = 'not tested'
$lblConnStatus.ForeColor = [System.Drawing.Color]::FromArgb(255,180,180,180)
$lblConnStatus.Font = New-Object System.Drawing.Font('Segoe UI',9)
$lblConnStatus.Location = New-Object System.Drawing.Point(32,15)
$lblConnStatus.Size = New-Object System.Drawing.Size(110,18)
[void]$pnlTop.Controls.Add($lblConnStatus)

$btnTest     = New-Btn $pnlTop '? Test'      150 6 74  34 70  70  85
$btnSettings = New-Btn $pnlTop 'Settings'    230 6 100 34 70  70  85
$btnRead     = New-Btn $pnlTop 'Read File'   336 6 120 34 40  100 140
$btnOpen     = New-Btn $pnlTop 'Open File'   462 6 120 34 40  110 60

# Excel path lives in Settings only (persisted to .env as EXCEL_FILE). We keep
# $txtFile as an off-screen holder so the rest of the UI (Read/Submit/Fill) can
# still read $txtFile.Text without threading the path through every handler.
$txtFile = New-Object System.Windows.Forms.TextBox
$txtFile.Text = if ($envData['EXCEL_FILE']) { $envData['EXCEL_FILE'] } else { Join-Path $scriptDir 'data\timesheet.xlsx' }

# ---- Toolbar row 2: actions that do work -----------------------------------
$pnlBtns = New-Object System.Windows.Forms.Panel
$pnlBtns.Location = New-Object System.Drawing.Point(0,46)
$pnlBtns.Size = New-Object System.Drawing.Size(1060,52)
$pnlBtns.BackColor = [System.Drawing.Color]::FromArgb(255,60,60,72)
[void]$form.Controls.Add($pnlBtns)

$btnFillStd = New-Btn $pnlBtns 'Standard ↓'         8   8 150 36 120 80  160
$btnSync    = New-Btn $pnlBtns 'Jira ↓'             166 8 150 36 0   122 200
$btnSubmit  = New-Btn $pnlBtns 'Submit tasks →'      324 8 185 36 180 80  0
$btnFill    = New-Btn $pnlBtns 'Fill days ↑↑↑↑↑'    517 8 220 36 80  80  80
$btnFill.Enabled = $false
$btnFill.ForeColor = [System.Drawing.Color]::FromArgb(255,140,140,140)

$btnStop = New-Btn $pnlBtns '⛔' 1024 8 22 36 140 30 30
$btnStop.Enabled = $false

# ---- Right-side Weeks sidebar (fixed width, full remaining height) ---------
$pnlWeeksSidebar = New-Object System.Windows.Forms.Panel
$pnlWeeksSidebar.Location = New-Object System.Drawing.Point(800,128)
$pnlWeeksSidebar.Size = New-Object System.Drawing.Size(260,524)
$pnlWeeksSidebar.BackColor = [System.Drawing.Color]::White
[void]$form.Controls.Add($pnlWeeksSidebar)

$pnlWeeksList = New-Object System.Windows.Forms.Panel
$pnlWeeksList.Dock = 'Fill'
$pnlWeeksList.AutoScroll = $true
$pnlWeeksList.BackColor = [System.Drawing.Color]::White

$pnlWeeksTop = New-Object System.Windows.Forms.Panel
$pnlWeeksTop.Dock = 'Top'; $pnlWeeksTop.Height = 64
$pnlWeeksTop.BackColor = [System.Drawing.Color]::White

$lblMonth = New-Object System.Windows.Forms.Label
$lblMonth.Text = 'Month:'
$lblMonth.Location = New-Object System.Drawing.Point(6,8)
$lblMonth.Size = New-Object System.Drawing.Size(44,18)
$lblMonth.Font = New-Object System.Drawing.Font('Segoe UI',9)
$lblMonth.ForeColor = [System.Drawing.Color]::FromArgb(255,80,80,100)
[void]$pnlWeeksTop.Controls.Add($lblMonth)

$cmbMonth = New-Object System.Windows.Forms.ComboBox
$cmbMonth.Location = New-Object System.Drawing.Point(54,6)
$cmbMonth.Size = New-Object System.Drawing.Size(190,22)
$cmbMonth.DropDownStyle = 'DropDownList'
$cmbMonth.Font = New-Object System.Drawing.Font('Segoe UI',9)
[void]$cmbMonth.Items.Add('All')
$cmbMonth.SelectedIndex = 0
[void]$pnlWeeksTop.Controls.Add($cmbMonth)

$chkWeeksAll = New-Object System.Windows.Forms.CheckBox
$chkWeeksAll.Text = 'Select all'
$chkWeeksAll.ThreeState = $true
$chkWeeksAll.Location = New-Object System.Drawing.Point(6,34)
$chkWeeksAll.Size = New-Object System.Drawing.Size(244,20)
$chkWeeksAll.Font = New-Object System.Drawing.Font('Segoe UI',9,[System.Drawing.FontStyle]::Bold)
$chkWeeksAll.Visible = $false
[void]$pnlWeeksTop.Controls.Add($chkWeeksAll)

$pnlWeeksBottom = New-Object System.Windows.Forms.Panel
$pnlWeeksBottom.Dock = 'Bottom'; $pnlWeeksBottom.Height = 64
$pnlWeeksBottom.BackColor = [System.Drawing.Color]::White

$lblWeeksSummary = New-Object System.Windows.Forms.Label
$lblWeeksSummary.Text = ''
$lblWeeksSummary.Location = New-Object System.Drawing.Point(6,6)
$lblWeeksSummary.Size = New-Object System.Drawing.Size(244,20)
$lblWeeksSummary.Font = New-Object System.Drawing.Font('Segoe UI',8.5)
$lblWeeksSummary.ForeColor = [System.Drawing.Color]::FromArgb(255,80,80,100)
[void]$pnlWeeksBottom.Controls.Add($lblWeeksSummary)

$btnCopyLog = New-Object System.Windows.Forms.Button
$btnCopyLog.Text = 'Copy log'
$btnCopyLog.Location = New-Object System.Drawing.Point(6,32)
$btnCopyLog.Size = New-Object System.Drawing.Size(244,24)
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
[void]$pnlWeeksBottom.Controls.Add($btnCopyLog)

# Add order matters: Fill panel first, then Top/Bottom strips after — each
# claims its own edge without covering the Fill panel (same pattern already
# proven safe for the Jira-queries tab's FlowLayoutPanel + add-strip).
[void]$pnlWeeksSidebar.Controls.Add($pnlWeeksList)
[void]$pnlWeeksSidebar.Controls.Add($pnlWeeksTop)
[void]$pnlWeeksSidebar.Controls.Add($pnlWeeksBottom)

$script:monthCodes = @{}
$script:weekCheckboxes = @()

function Update-MasterCheckboxState {
    if ($script:weekCheckboxes.Count -eq 0) { $chkWeeksAll.CheckState = 'Unchecked'; return }
    $checkedCount = @($script:weekCheckboxes | Where-Object { $_.Checked }).Count
    if ($checkedCount -eq 0) { $chkWeeksAll.CheckState = 'Unchecked' }
    elseif ($checkedCount -eq $script:weekCheckboxes.Count) { $chkWeeksAll.CheckState = 'Checked' }
    else { $chkWeeksAll.CheckState = 'Indeterminate' }
}

$chkWeeksAll.Add_Click({
    # WinForms auto-cycles a ThreeState CheckBox's CheckState on click, in order
    # Unchecked -> Checked -> Indeterminate -> Unchecked, BEFORE this handler runs.
    # So by the time we read $chkWeeksAll.CheckState here, it already reflects the
    # post-click auto-cycled value:
    #   prior Unchecked    -> now Checked        -> already means "select all", keep it
    #   prior Checked      -> now Indeterminate  -> override to mean "deselect all"
    #   prior Indeterminate -> now Unchecked     -> override to mean "select all"
    #                          (a real click never leaves the box on Indeterminate;
    #                          only Update-MasterCheckboxState's recompute does that)
    if ($chkWeeksAll.CheckState -eq 'Indeterminate') {
        foreach ($c in $script:weekCheckboxes) { $c.Checked = $false }
        $chkWeeksAll.CheckState = 'Unchecked'
    } else {
        foreach ($c in $script:weekCheckboxes) { $c.Checked = $true }
        $chkWeeksAll.CheckState = 'Checked'
    }
})

function Populate-Weeks($weeksJson) {
    $toRemove = @($pnlWeeksList.Controls | Where-Object { $_ -is [System.Windows.Forms.CheckBox] })
    foreach ($c in $toRemove) { $pnlWeeksList.Controls.Remove($c) }
    $script:weekCheckboxes = @()
    $chkWeeksAll.Visible = $false
    if (-not $weeksJson -or $weeksJson.Count -eq 0) {
        $chkWeeksAll.CheckState = 'Unchecked'
        return
    }

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

    $i = 0
    foreach ($w in $weeksJson) {
        $chk = New-Object System.Windows.Forms.CheckBox
        $chk.Text = $w.start.Substring(5) + ' – ' + $w.end.Substring(5)
        $chk.Location = New-Object System.Drawing.Point(6, (4 + $i * 24))
        $chk.Size = New-Object System.Drawing.Size(230, 20)
        $chk.Checked = $true
        $chk.Font = New-Object System.Drawing.Font('Segoe UI',8.5)
        $chk.Tag = $w.start
        $chk.Add_CheckedChanged({ Update-MasterCheckboxState })
        [void]$pnlWeeksList.Controls.Add($chk)
        $script:weekCheckboxes += $chk
        $i++
    }
    $chkWeeksAll.Visible = $true
    Update-MasterCheckboxState
}

# ---- Divider ---------------------------------------------------------------
$divider = New-Object System.Windows.Forms.Label
$divider.Location = New-Object System.Drawing.Point(0,98)
$divider.Size = New-Object System.Drawing.Size(1060,1)
$divider.BackColor = [System.Drawing.Color]::FromArgb(255,200,200,210)
[void]$form.Controls.Add($divider)

# ---- Status bar ------------------------------------------------------------
$pnlStatus = New-Object System.Windows.Forms.Panel
$pnlStatus.Location = New-Object System.Drawing.Point(0,99)
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
$lblStatus.Size = New-Object System.Drawing.Size(1020,18)
$lblStatus.Font = New-Object System.Drawing.Font('Segoe UI',9)
[void]$pnlStatus.Controls.Add($lblStatus)

# ---- Log -------------------------------------------------------------------
$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Location = New-Object System.Drawing.Point(0,128)
$txtLog.Size = New-Object System.Drawing.Size(800,524)
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.Color]::FromArgb(255,20,20,26)
$txtLog.ForeColor = [System.Drawing.Color]::FromArgb(255,200,200,200)
$txtLog.Font = New-Object System.Drawing.Font('Consolas',9)
$txtLog.BorderStyle = 'None'
$txtLog.ScrollBars = 'Vertical'
[void]$form.Controls.Add($txtLog)

$form.Add_Resize({
    $h = $form.ClientSize.Height
    $newHeight = [Math]::Max($h - $txtLog.Top, 50)
    $txtLog.Height = $newHeight
    $pnlWeeksSidebar.Height = $newHeight
})

# ============================================================
# Settings dialog
# ============================================================
function Show-Settings {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Settings'
    $dlg.ClientSize = New-Object System.Drawing.Size(612,518)
    $dlg.StartPosition = 'CenterParent'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font('Segoe UI',9)
    $dlg.BackColor = [System.Drawing.Color]::FromArgb(255,248,248,252)

    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Location = New-Object System.Drawing.Point(8,8)
    $tabs.Size = New-Object System.Drawing.Size(596,430)
    [void]$dlg.Controls.Add($tabs)

    $tabConn = New-Object System.Windows.Forms.TabPage; $tabConn.Text = 'Connection'
    $tabStd  = New-Object System.Windows.Forms.TabPage; $tabStd.Text  = 'Standard tasks'
    $tabJql  = New-Object System.Windows.Forms.TabPage; $tabJql.Text  = 'Jira queries'
    $tabConn.BackColor = [System.Drawing.Color]::White
    $tabStd.BackColor  = [System.Drawing.Color]::White
    $tabJql.BackColor  = [System.Drawing.Color]::White
    $tabs.TabPages.AddRange(@($tabConn, $tabStd, $tabJql))

    $jqlSlots = @(
        @{ key='investigation';      label='1. Investigation issues' },
        @{ key='bug_verification';   label='2. Bug verification' },
        @{ key='story_creation';     label='3. User story creation' },
        @{ key='functional_testing'; label='4. Functional testing' },
        @{ key='regression_testing'; label='5. Regression testing' },
        @{ key='other_qa';           label='6. Other QA activities' }
    )
    $defaultJql = @{
        investigation      = 'project = {project} AND issuetype = Bug AND created >= "{ws}" AND created <= "{we}" AND (creator = {account_id} OR reporter = {account_id}) ORDER BY priority DESC, issuetype ASC, key ASC'
        bug_verification   = 'project = {project} AND issuetype = Bug AND status CHANGED TO "Done" BY {account_id} DURING ("{ws}","{we}") ORDER BY priority DESC, issuetype ASC, key ASC'
        story_creation     = 'project = {project} AND issuetype = Story AND created >= "{ws}" AND created <= "{we}" AND creator = {account_id} AND parent = {project}-80 ORDER BY status DESC, issuetype ASC, key ASC'
        functional_testing = 'project = {project} AND issuetype = Story AND (status CHANGED FROM "Ready for QA" BY {account_id} DURING ("{ws}", "{we}") OR status CHANGED FROM "IN QA" BY {account_id} DURING ("{ws}", "{we}")) ORDER BY key ASC'
        regression_testing = 'project = {project} AND status CHANGED TO Done BY {account_id} DURING ("{ws}", "{we}") AND parent = {project}-73 AND summary ~ "Regression" ORDER BY status DESC, issuetype ASC, key ASC'
        other_qa           = 'project = {project} AND status CHANGED TO Done BY {account_id} DURING ("{ws}", "{we}") AND parent = {project}-73 AND summary !~ "Smoke" AND summary !~ "Regression" AND summary !~ "Functional." ORDER BY status DESC, issuetype ASC, key ASC'
    }
    $jqlCfg = Read-JsonConfig 'jql_queries.json'
    $jqlValues = @{}
    foreach ($k in $defaultJql.Keys) { $jqlValues[$k] = $defaultJql[$k] }
    if ($jqlCfg -and $jqlCfg.queries) {
        foreach ($q in $jqlCfg.queries) { if ($jqlValues.ContainsKey($q.key) -and $q.jql) { $jqlValues[$q.key] = $q.jql } }
    }

    $panelJqlFlow = New-Object System.Windows.Forms.FlowLayoutPanel
    $panelJqlFlow.Location = New-Object System.Drawing.Point(0,0)
    $panelJqlFlow.Dock = 'Fill'
    $panelJqlFlow.FlowDirection = 'TopDown'
    $panelJqlFlow.WrapContents = $false
    $panelJqlFlow.AutoScroll = $true
    [void]$tabJql.Controls.Add($panelJqlFlow)

    function New-JqlRow($key, $titleText, $jqlText, $titleEditable) {
        $row = New-Object System.Windows.Forms.Panel
        $row.Size = New-Object System.Drawing.Size(566,74)
        $row.Margin = New-Object System.Windows.Forms.Padding(4,4,4,0)
        $row.Tag = $key

        if ($titleEditable) {
            $titleCtl = New-Object System.Windows.Forms.TextBox
        } else {
            $titleCtl = New-Object System.Windows.Forms.Label
        }
        $titleCtl.Text = $titleText; $titleCtl.Location = New-Object System.Drawing.Point(0,0)
        $titleCtl.Size = New-Object System.Drawing.Size(420,18)
        $titleCtl.Font = New-Object System.Drawing.Font('Segoe UI',9,[System.Drawing.FontStyle]::Bold)
        [void]$row.Controls.Add($titleCtl)

        $box = New-Object System.Windows.Forms.TextBox
        $box.Multiline = $true; $box.ScrollBars = 'Vertical'; $box.WordWrap = $true
        $box.Location = New-Object System.Drawing.Point(0,20); $box.Size = New-Object System.Drawing.Size(420,44)
        $box.Font = New-Object System.Drawing.Font('Consolas',8)
        $box.Text = $jqlText
        [void]$row.Controls.Add($box)

        $btnCopy = New-Object System.Windows.Forms.Button
        $btnCopy.Text = 'Copy'; $btnCopy.Location = New-Object System.Drawing.Point(424,20); $btnCopy.Size = New-Object System.Drawing.Size(60,21)
        $btnCopy.FlatStyle = 'Flat'
        $btnCopy.Add_Click({
            $tpl = [string]$this.Parent.Controls[1].Text
            $proj = if ($tProj.Text) { $tProj.Text } else { 'GT2' }
            $acct = $tAcct.Text
            $checked = @($script:weekCheckboxes | Where-Object { $_.Checked } | ForEach-Object { $_.Tag })
            if ($checked.Count -gt 0) {
                $ws = [string]$checked[0]
                $we = ([datetime]::ParseExact($ws,'yyyy-MM-dd',$null).AddDays(6)).ToString('yyyy-MM-dd')
            } else {
                $today = Get-Date
                $sunday = $today.AddDays(-[int]$today.DayOfWeek)
                $ws = $sunday.ToString('yyyy-MM-dd'); $we = $sunday.AddDays(6).ToString('yyyy-MM-dd')
            }
            $resolved = $tpl.Replace('{project}',$proj).Replace('{account_id}',$acct).Replace('{ws}',$ws).Replace('{we}',$we)
            [System.Windows.Forms.Clipboard]::SetText($resolved)
        })
        [void]$row.Controls.Add($btnCopy)

        $btnDel = New-Object System.Windows.Forms.Button
        $btnDel.Text = 'Delete'; $btnDel.Location = New-Object System.Drawing.Point(424,44); $btnDel.Size = New-Object System.Drawing.Size(60,20)
        $btnDel.FlatStyle = 'Flat'; $btnDel.ForeColor = [System.Drawing.Color]::FromArgb(255,180,40,40)
        $btnDel.Add_Click({
            $this.Parent.Parent.Controls.Remove($this.Parent)
        })
        [void]$row.Controls.Add($btnDel)

        [void]$panelJqlFlow.Controls.Add($row)
        return $row
    }

    $hiddenBuiltin = @{}
    if ($jqlCfg -and $jqlCfg.PSObject.Properties.Match('hidden_builtin').Count -and $jqlCfg.hidden_builtin) {
        foreach ($hk in $jqlCfg.hidden_builtin) { $hiddenBuiltin[$hk] = $true }
    }

    foreach ($slot in $jqlSlots) {
        if ($hiddenBuiltin.ContainsKey($slot.key)) { continue }
        New-JqlRow $slot.key $slot.label ([string]$jqlValues[$slot.key]) $false
    }

    if ($jqlCfg -and $jqlCfg.queries) {
        $fixedKeySet = @{}
        foreach ($s in $jqlSlots) { $fixedKeySet[$s.key] = $true }
        foreach ($q in $jqlCfg.queries) {
            if (-not $fixedKeySet.ContainsKey($q.key)) {
                New-JqlRow $q.key $q.label $q.jql $true
            }
        }
    }

    $panelJqlAddStrip = New-Object System.Windows.Forms.Panel
    $panelJqlAddStrip.Dock = 'Bottom'; $panelJqlAddStrip.Height = 32
    $btnAddJql = New-Object System.Windows.Forms.Button
    $btnAddJql.Text = '+ Add query'; $btnAddJql.Location = New-Object System.Drawing.Point(0,2)
    $btnAddJql.Size = New-Object System.Drawing.Size(110,26); $btnAddJql.FlatStyle = 'Flat'
    $btnAddJql.Add_Click({
        $newKey = 'extra_' + ([guid]::NewGuid().ToString('N').Substring(0,8))
        $newRow = New-JqlRow $newKey '' '' $true
        $panelJqlFlow.ScrollControlIntoView($newRow)
        $newRow.Controls[0].Focus()
    })
    [void]$panelJqlAddStrip.Controls.Add($btnAddJql)
    [void]$tabJql.Controls.Add($panelJqlAddStrip)

    $d = Read-EnvFile
    $lblSprintLen = New-Object System.Windows.Forms.Label
    $lblSprintLen.Text = 'Sprint length (weeks):'
    $lblSprintLen.Location = New-Object System.Drawing.Point(8,10)
    $lblSprintLen.Size = New-Object System.Drawing.Size(140,20)
    [void]$tabStd.Controls.Add($lblSprintLen)
    $tSprintLen = New-Object System.Windows.Forms.TextBox
    $tSprintLen.Location = New-Object System.Drawing.Point(152,8)
    $tSprintLen.Size = New-Object System.Drawing.Size(40,24)
    [void]$tabStd.Controls.Add($tSprintLen)

    $lblAnchorRow = New-Object System.Windows.Forms.Label
    $lblAnchorRow.Text = 'Sprint end:'
    $lblAnchorRow.Location = New-Object System.Drawing.Point(210,10)
    $lblAnchorRow.Size = New-Object System.Drawing.Size(70,20)
    [void]$tabStd.Controls.Add($lblAnchorRow)
    $tAnchor = New-Object System.Windows.Forms.TextBox
    $tAnchor.Location = New-Object System.Drawing.Point(284,8)
    $tAnchor.Size = New-Object System.Drawing.Size(110,24)
    [void]$tabStd.Controls.Add($tAnchor)

    $lblAnchorHint = New-Object System.Windows.Forms.Label
    $lblAnchorHint.Location = New-Object System.Drawing.Point(8,36)
    $lblAnchorHint.Size = New-Object System.Drawing.Size(572,16)
    $lblAnchorHint.ForeColor = [System.Drawing.Color]::Gray
    $lblAnchorHint.Font = New-Object System.Drawing.Font('Segoe UI',8)
    [void]$tabStd.Controls.Add($lblAnchorHint)

    function Update-AnchorHint {
        $weeks = 2
        $parsed = 0
        if ([int]::TryParse($tSprintLen.Text, [ref]$parsed) -and $parsed -gt 0) { $weeks = $parsed }
        $lblAnchorHint.Text = "Any sprint-end Friday (YYYY-MM-DD) - $weeks-week cycles counted from here"
    }
    $tSprintLen.Add_TextChanged({ Update-AnchorHint })

    $tAnchor.Text = if ($d['SPRINT_ANCHOR']) { $d['SPRINT_ANCHOR'] } else { '' }
    $tSprintLen.Text = if ($d['SPRINT_LENGTH_WEEKS']) { $d['SPRINT_LENGTH_WEEKS'] } else { '2' }
    Update-AnchorHint

    $gridStd = New-Object System.Windows.Forms.DataGridView
    $gridStd.Location = New-Object System.Drawing.Point(8,60)
    $gridStd.Size = New-Object System.Drawing.Size(572,304)
    $gridStd.AllowUserToAddRows = $true
    $gridStd.AllowUserToDeleteRows = $true
    $gridStd.AutoSizeColumnsMode = 'Fill'
    [void]$tabStd.Controls.Add($gridStd)
    $colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colName.HeaderText = 'Task'; $colName.FillWeight = 200
    [void]$gridStd.Columns.Add($colName)
    foreach ($dh in 'Mon','Tue','Wed','Thu','Fri') {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $col.HeaderText = $dh; $col.FillWeight = 45
        [void]$gridStd.Columns.Add($col)
    }
    $colFreq = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
    $colFreq.HeaderText = 'Frequency'; $colFreq.FillWeight = 90
    [void]$colFreq.Items.AddRange(@('weekly','sprint-end','placeholder'))
    [void]$gridStd.Columns.Add($colFreq)
    $colDel = New-Object System.Windows.Forms.DataGridViewButtonColumn
    $colDel.Name = 'Delete'; $colDel.HeaderText = ''
    $colDel.Text = 'Delete'; $colDel.UseColumnTextForButtonValue = $true
    $colDel.FillWeight = 55
    [void]$gridStd.Columns.Add($colDel)
    $gridStd.Add_CellContentClick({
        param($eventSender, $e)
        if ($e.RowIndex -ge 0 -and $gridStd.Columns[$e.ColumnIndex].Name -eq 'Delete') {
            if (-not $gridStd.Rows[$e.RowIndex].IsNewRow) { $gridStd.Rows.RemoveAt($e.RowIndex) }
        }
    })

    $lblStdHint = New-Object System.Windows.Forms.Label
    $lblStdHint.Text = 'One hours value per task (same each day it occurs); blank = not that day. Placeholders: name only.'
    $lblStdHint.Location = New-Object System.Drawing.Point(8,368); $lblStdHint.Size = New-Object System.Drawing.Size(572,18)
    $lblStdHint.ForeColor = [System.Drawing.Color]::Gray; $lblStdHint.Font = New-Object System.Drawing.Font('Segoe UI',8)
    [void]$tabStd.Controls.Add($lblStdHint)

    $stdCfg = Read-JsonConfig 'standard_tasks.json'
    $defaultSchedule = @(
        @{name='Internal Daily meeting'; hours=0.5; days=@('Mon','Tue','Wed','Thu','Fri'); freq='weekly'},
        @{name='Internal bug triage'; hours=0.5; days=@('Tue'); freq='weekly'},
        @{name='External customer meeting'; hours=1.0; days=@('Tue','Wed'); freq='weekly'},
        @{name='Weekly project report'; hours=1.0; days=@('Fri'); freq='weekly'},
        @{name='Internal sprint review'; hours=0.5; days=@('Fri'); freq='sprint-end'},
        @{name='External sprint review'; hours=1.0; days=@('Fri'); freq='sprint-end'},
        @{name='Summary report creation'; hours=2.0; days=@('Fri'); freq='sprint-end'}
    )
    $defaultPlaceholders = @('Bug verification','Functional testing','Automation test maintenance','Investigation issue')

    $schedule = if ($stdCfg -and $stdCfg.PSObject.Properties.Match('schedule').Count) { $stdCfg.schedule } else { $defaultSchedule }
    $placeholders = if ($stdCfg -and $stdCfg.PSObject.Properties.Match('placeholders').Count) { $stdCfg.placeholders } else { $defaultPlaceholders }

    $dayIndex = @{ Mon=1; Tue=2; Wed=3; Thu=4; Fri=5 }
    foreach ($t in $schedule) {
        $cells = @($t.name, '', '', '', '', '', $t.freq)
        foreach ($dn in $t.days) { $cells[$dayIndex[$dn]] = [string]$t.hours }
        [void]$gridStd.Rows.Add($cells)
    }
    foreach ($ph in $placeholders) {
        [void]$gridStd.Rows.Add(@($ph, '', '', '', '', '', 'placeholder'))
    }

    $script:SaveStandardFromGrid = {
        $sched = @(); $ph = @()
        foreach ($row in $gridStd.Rows) {
            if ($row.IsNewRow) { continue }
            $name = [string]$row.Cells[0].Value
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            $freq = [string]$row.Cells[6].Value
            if ($freq -eq 'placeholder') { $ph += $name.Trim(); continue }
            $days = @(); $hours = $null
            foreach ($dn in 'Mon','Tue','Wed','Thu','Fri') {
                $v = [string]$row.Cells[$dayIndex[$dn]].Value
                if (-not [string]::IsNullOrWhiteSpace($v)) { $days += $dn; $hours = [double]$v }
            }
            if ($days.Count -eq 0) { continue }
            $sched += @{ name=$name.Trim(); hours=$hours; days=$days; freq=$(if ($freq) { $freq } else { 'weekly' }) }
        }
        Write-JsonConfig 'standard_tasks.json' @{ schedule=$sched; placeholders=$ph }
    }

    $script:SaveJqlConfig = {
        $queries = @()
        $fixedKeys = @($jqlSlots | ForEach-Object { $_.key })
        $presentFixed = @{}
        foreach ($row in $panelJqlFlow.Controls) {
            $title = [string]$row.Controls[0].Text
            $jql   = [string]$row.Controls[1].Text
            if ($fixedKeys -contains $row.Tag) { $presentFixed[$row.Tag] = $true }
            if ([string]::IsNullOrWhiteSpace($title) -and [string]::IsNullOrWhiteSpace($jql)) { continue }
            $queries += @{ key=$row.Tag; label=$title; jql=$jql }
        }
        $hidden = @($fixedKeys | Where-Object { -not $presentFixed.ContainsKey($_) })
        Write-JsonConfig 'jql_queries.json' @{ queries=$queries; hidden_builtin=$hidden }
    }

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
    $tUrl  = Add-Row $tabConn 'Jira URL:'    16
    $tTok  = Add-Row $tabConn 'API Token:'   54 $true

    $chkSh = New-Object System.Windows.Forms.CheckBox
    $chkSh.Text = 'Show token'; $chkSh.Location = New-Object System.Drawing.Point(114,80)
    $chkSh.Size = New-Object System.Drawing.Size(100,22)
    $chkSh.Add_CheckedChanged({ $tTok.UseSystemPasswordChar = -not $chkSh.Checked })
    [void]$tabConn.Controls.Add($chkSh)

    $tProj = Add-Row $tabConn 'Project key:' 110
    $tMail = Add-Row $tabConn 'Email:'       148
    $tAcct = Add-Row $tabConn 'Account ID:'  186

    $tUrl.Text  = if ($d['JIRA_URL'])        { $d['JIRA_URL'] }        else { 'https://amcbridge.atlassian.net' }
    $tTok.Text  = if ($d['JIRA_API_TOKEN'])   { $d['JIRA_API_TOKEN'] }   else { '' }
    $tProj.Text = if ($d['JIRA_PROJECT'])     { $d['JIRA_PROJECT'] }     else { 'GT2' }
    $tMail.Text = if ($d['JIRA_EMAIL'])       { $d['JIRA_EMAIL'] }       else { '' }
    $tAcct.Text = if ($d['JIRA_ACCOUNT_ID'])  { $d['JIRA_ACCOUNT_ID'] }  else { '' }

    # Hint: which Jira user the activity queries (2,3,4,5,6) filter on. Empty = token owner.
    $lblAcctHint = New-Object System.Windows.Forms.Label
    $lblAcctHint.Text = 'Leave empty to use the API-token owner (/myself)'
    $lblAcctHint.Location = New-Object System.Drawing.Point(114,211)
    $lblAcctHint.Size = New-Object System.Drawing.Size(390,16)
    $lblAcctHint.ForeColor = [System.Drawing.Color]::Gray
    $lblAcctHint.Font = New-Object System.Drawing.Font('Segoe UI',8)
    [void]$tabConn.Controls.Add($lblAcctHint)

    # ---- Excel file row ----
    $lblEx = New-Object System.Windows.Forms.Label
    $lblEx.Text = 'Excel file:'; $lblEx.Location = New-Object System.Drawing.Point(16,237)
    $lblEx.Size = New-Object System.Drawing.Size(90,20); $lblEx.TextAlign = 'MiddleRight'
    [void]$tabConn.Controls.Add($lblEx)
    $tExcel = New-Object System.Windows.Forms.TextBox
    $tExcel.Location = New-Object System.Drawing.Point(114,234)
    $tExcel.Size = New-Object System.Drawing.Size(354,24)
    $tExcel.Text = $txtFile.Text
    [void]$tabConn.Controls.Add($tExcel)
    $btnExBrowse = New-Object System.Windows.Forms.Button
    $btnExBrowse.Text = '...'; $btnExBrowse.Location = New-Object System.Drawing.Point(474,234)
    $btnExBrowse.Size = New-Object System.Drawing.Size(30,24); $btnExBrowse.FlatStyle = 'Flat'
    $btnExBrowse.Add_Click({
        $fd = New-Object System.Windows.Forms.OpenFileDialog
        $fd.Filter = 'Excel files (*.xlsx)|*.xlsx'
        $fd.InitialDirectory = Split-Path $tExcel.Text -Parent
        if ($fd.ShowDialog() -eq 'OK') { $tExcel.Text = $fd.FileName }
    })
    [void]$tabConn.Controls.Add($btnExBrowse)

    $lblTest = New-Object System.Windows.Forms.Label
    $lblTest.Location = New-Object System.Drawing.Point(16,272)
    $lblTest.Size = New-Object System.Drawing.Size(460,20)
    $lblTest.ForeColor = [System.Drawing.Color]::Gray
    [void]$tabConn.Controls.Add($lblTest)

    $btnT = New-Object System.Windows.Forms.Button
    $btnT.Text = 'Test Connection'
    $btnT.Location = New-Object System.Drawing.Point(16,296)
    $btnT.Size = New-Object System.Drawing.Size(140,32)
    $btnT.Add_Click({
        $lblTest.Text = 'Testing...'; $lblTest.ForeColor = [System.Drawing.Color]::DodgerBlue
        $dlg.Refresh()
        Save-JiraEnv $tUrl.Text $tMail.Text $tTok.Text $tProj.Text $tAcct.Text $tExcel.Text $tAnchor.Text $tSprintLen.Text
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
    [void]$tabConn.Controls.Add($btnT)

    $btnSv = New-Object System.Windows.Forms.Button
    $btnSv.Text = 'Save & Close'; $btnSv.DialogResult = 'OK'
    $btnSv.Location = New-Object System.Drawing.Point(470,470)
    $btnSv.Size = New-Object System.Drawing.Size(120,32)
    $btnSv.BackColor = [System.Drawing.Color]::FromArgb(255,0,122,200)
    $btnSv.ForeColor = [System.Drawing.Color]::White; $btnSv.FlatStyle = 'Flat'
    $btnSv.Add_Click({
        Save-JiraEnv $tUrl.Text $tMail.Text $tTok.Text $tProj.Text $tAcct.Text $tExcel.Text $tAnchor.Text $tSprintLen.Text
        & $script:SaveStandardFromGrid
        & $script:SaveJqlConfig
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
            $lblWeeksSummary.Text = "File loaded: " + $weeks.Count + " week(s) found"
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