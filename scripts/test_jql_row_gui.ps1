Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:weekCheckboxes = @()
$tProj = New-Object System.Windows.Forms.TextBox; $tProj.Text = 'GT2'
$tAcct = New-Object System.Windows.Forms.TextBox; $tAcct.Text = 'ACC123'
$panelJqlFlow = New-Object System.Windows.Forms.FlowLayoutPanel

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

$script:failures = @()
function Check($name, $cond) {
    if ($cond) { Write-Host "  OK $name" } else { $script:failures += $name; Write-Host "  FAIL $name" }
}

# --- fixed-shaped row (titleEditable=$false), no {ws}/{we}/{account_id} needed for a stable assert ---
$fixedRow = New-JqlRow 'investigation' '1. Investigation issues' 'project = {project} AND creator = {account_id}' $false
[System.Windows.Forms.Clipboard]::SetText('')
$fixedRow.Controls[2].PerformClick()   # Copy
$clipFixed = [System.Windows.Forms.Clipboard]::GetText()
Check "fixed row Copy substitutes project/account" ($clipFixed -eq 'project = GT2 AND creator = ACC123')

# --- custom-shaped row (titleEditable=$true) ---
$customRow = New-JqlRow 'extra_ab12cd34' 'My query' 'project = {project} custom' $true
[System.Windows.Forms.Clipboard]::SetText('')
$customRow.Controls[2].PerformClick()  # Copy
$clipCustom = [System.Windows.Forms.Clipboard]::GetText()
Check "custom row Copy substitutes project" ($clipCustom -eq 'project = GT2 custom')

# --- Delete removes the row from the panel, for both shapes ---
$countBeforeFixed = $panelJqlFlow.Controls.Count
$fixedRow.Controls[3].PerformClick()   # Delete
Check "fixed row Delete removes it from the panel" (
    $panelJqlFlow.Controls.Count -eq ($countBeforeFixed - 1) -and -not $panelJqlFlow.Controls.Contains($fixedRow)
)

$countBeforeCustom = $panelJqlFlow.Controls.Count
$customRow.Controls[3].PerformClick()  # Delete
Check "custom row Delete removes it from the panel" (
    $panelJqlFlow.Controls.Count -eq ($countBeforeCustom - 1) -and -not $panelJqlFlow.Controls.Contains($customRow)
)

if ($script:failures.Count -gt 0) {
    Write-Host "FAILURES: $($script:failures -join ', ')"
    exit 1
} else {
    Write-Host "ALL PASS"
    exit 0
}
