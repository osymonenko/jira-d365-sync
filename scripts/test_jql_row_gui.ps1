Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:weekCheckboxes = @()
$tProj = New-Object System.Windows.Forms.TextBox; $tProj.Text = 'GT2'
$tAcct = New-Object System.Windows.Forms.TextBox; $tAcct.Text = 'ACC123'
$panelJqlFlow = New-Object System.Windows.Forms.FlowLayoutPanel
$jqlTips = New-Object System.Windows.Forms.ToolTip

$modeDisplay = @{ count='Total count'; priority='Priority breakdown (P1-P3)'; per_issue='One row per issue' }
$modeValue   = @{ 'Total count'='count'; 'Priority breakdown (P1-P3)'='priority'; 'One row per issue'='per_issue' }
$modeHint = @{
    count     = 'Placeholders: {count}'
    priority  = 'Placeholders: {P1} {P2} {P3} {count} {buckets}'
    per_issue = 'Placeholders: {key} {summary} {number}'
}

function New-JqlRow($key, $titleText, $jqlText, $titleEditable, $modeKey = $null) {
    $row = New-Object System.Windows.Forms.Panel
    $row.Size = New-Object System.Drawing.Size(566,74)
    $row.Margin = New-Object System.Windows.Forms.Padding(4,4,4,0)
    $row.Tag = $key

    if ($titleEditable) {
        $titleCtl = New-Object System.Windows.Forms.TextBox
    } else {
        $titleCtl = New-Object System.Windows.Forms.Label
    }
    $titleCtl.Name = 'titleBox'
    $titleCtl.Text = $titleText; $titleCtl.Location = New-Object System.Drawing.Point(0,0)
    $titleWidth = if ($modeKey) { 260 } else { 420 }
    $titleCtl.Size = New-Object System.Drawing.Size($titleWidth,18)
    $titleCtl.Font = New-Object System.Drawing.Font('Segoe UI',9,[System.Drawing.FontStyle]::Bold)
    [void]$row.Controls.Add($titleCtl)

    if ($modeKey) {
        $combo = New-Object System.Windows.Forms.ComboBox
        $combo.Name = 'modeCombo'
        $combo.DropDownStyle = 'DropDownList'
        $combo.Location = New-Object System.Drawing.Point(264,0)
        $combo.Size = New-Object System.Drawing.Size(156,20)
        $combo.Font = New-Object System.Drawing.Font('Segoe UI',8)
        [void]$combo.Items.AddRange(@('Total count','Priority breakdown (P1-P3)','One row per issue'))
        $combo.SelectedItem = $modeDisplay[$modeKey]
        $jqlTips.SetToolTip($titleCtl, $modeHint[$modeKey])
        $jqlTips.SetToolTip($combo, [string]$combo.SelectedItem)
        $combo.Add_SelectedIndexChanged({
            $newMode = $modeValue[[string]$this.SelectedItem]
            $jqlTips.SetToolTip($this.Parent.Controls.Find('titleBox',$false)[0], $modeHint[$newMode])
            $jqlTips.SetToolTip($this, [string]$this.SelectedItem)
        })
        [void]$row.Controls.Add($combo)
    }

    $box = New-Object System.Windows.Forms.TextBox
    $box.Name = 'jqlBox'
    $box.Multiline = $true; $box.ScrollBars = 'Vertical'; $box.WordWrap = $true
    $box.Location = New-Object System.Drawing.Point(0,20); $box.Size = New-Object System.Drawing.Size(420,44)
    $box.Font = New-Object System.Drawing.Font('Consolas',8)
    $box.Text = $jqlText
    [void]$row.Controls.Add($box)

    $btnCopy = New-Object System.Windows.Forms.Button
    $btnCopy.Text = 'Copy'; $btnCopy.Location = New-Object System.Drawing.Point(424,20); $btnCopy.Size = New-Object System.Drawing.Size(60,21)
    $btnCopy.FlatStyle = 'Flat'
    $btnCopy.Add_Click({
        $tpl = [string]$this.Parent.Controls.Find('jqlBox',$false)[0].Text
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

$script:failures = @()
function Check($name, $cond) {
    if ($cond) { Write-Host "  OK $name" } else { $script:failures += $name; Write-Host "  FAIL $name" }
}

function Get-Copy($row) { ($row.Controls | Where-Object { $_ -is [System.Windows.Forms.Button] -and $_.Text -eq 'Copy' })[0] }
function Get-Delete($row) { ($row.Controls | Where-Object { $_ -is [System.Windows.Forms.Button] -and $_.Text -eq 'Delete' })[0] }

# --- fixed-shaped row (titleEditable=$false, no mode), e.g. other_qa ---
$fixedRow = New-JqlRow 'other_qa' 'Other QA activities (auto-named by keyword rules)' 'project = {project} AND creator = {account_id}' $false
Check "fixed row has no mode combo" ($fixedRow.Controls.Find('modeCombo',$false).Count -eq 0)
[System.Windows.Forms.Clipboard]::SetText('')
(Get-Copy $fixedRow).PerformClick()
Check "fixed row Copy substitutes project/account" ([System.Windows.Forms.Clipboard]::GetText() -eq 'project = GT2 AND creator = ACC123')

# --- custom-shaped row (titleEditable=$true, no mode) ---
$customRow = New-JqlRow 'extra_ab12cd34' 'My query' 'project = {project} custom' $true
Check "custom row has no mode combo" ($customRow.Controls.Find('modeCombo',$false).Count -eq 0)
[System.Windows.Forms.Clipboard]::SetText('')
(Get-Copy $customRow).PerformClick()
Check "custom row Copy substitutes project" ([System.Windows.Forms.Clipboard]::GetText() -eq 'project = GT2 custom')

# --- templatable row (titleEditable=$true, modeKey set), e.g. bug_verification ---
$tplRow = New-JqlRow 'bug_verification' 'Bug verification {buckets}' 'project = {project} custom' $true 'priority'
$combo = $tplRow.Controls.Find('modeCombo',$false)[0]
Check "templatable row has mode combo pre-selected" ([string]$combo.SelectedItem -eq 'Priority breakdown (P1-P3)')
Check "templatable row title tooltip matches priority hint" ($jqlTips.GetToolTip($tplRow.Controls.Find('titleBox',$false)[0]) -eq 'Placeholders: {P1} {P2} {P3} {count} {buckets}')
$combo.SelectedItem = 'Total count'
Check "changing mode updates title tooltip" ($jqlTips.GetToolTip($tplRow.Controls.Find('titleBox',$false)[0]) -eq 'Placeholders: {count}')
[System.Windows.Forms.Clipboard]::SetText('')
(Get-Copy $tplRow).PerformClick()
Check "templatable row Copy still substitutes JQL box, not the template text" ([System.Windows.Forms.Clipboard]::GetText() -eq 'project = GT2 custom')

# --- Delete removes the row from the panel, for all three shapes ---
foreach ($r in @($fixedRow, $customRow, $tplRow)) {
    $countBefore = $panelJqlFlow.Controls.Count
    (Get-Delete $r).PerformClick()
    Check "Delete removes row '$($r.Tag)' from the panel" (
        $panelJqlFlow.Controls.Count -eq ($countBefore - 1) -and -not $panelJqlFlow.Controls.Contains($r)
    )
}

if ($script:failures.Count -gt 0) {
    Write-Host "FAILURES: $($script:failures -join ', ')"
    exit 1
} else {
    Write-Host "ALL PASS"
    exit 0
}
