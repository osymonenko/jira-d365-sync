Add-Type -AssemblyName System.Windows.Forms

$script:weekCheckboxes = @()

$chkWeeksAll = New-Object System.Windows.Forms.CheckBox
$chkWeeksAll.ThreeState = $true

function Update-MasterCheckboxState {
    if ($script:weekCheckboxes.Count -eq 0) { $chkWeeksAll.CheckState = 'Unchecked'; return }
    $checkedCount = @($script:weekCheckboxes | Where-Object { $_.Checked }).Count
    if ($checkedCount -eq 0) { $chkWeeksAll.CheckState = 'Unchecked' }
    elseif ($checkedCount -eq $script:weekCheckboxes.Count) { $chkWeeksAll.CheckState = 'Checked' }
    else { $chkWeeksAll.CheckState = 'Indeterminate' }
}

$chkWeeksAll.Add_Click({
    if ($chkWeeksAll.CheckState -eq 'Indeterminate') {
        foreach ($c in $script:weekCheckboxes) { $c.Checked = $false }
        $chkWeeksAll.CheckState = 'Unchecked'
    } else {
        foreach ($c in $script:weekCheckboxes) { $c.Checked = $true }
        $chkWeeksAll.CheckState = 'Checked'
    }
})

$week1 = New-Object System.Windows.Forms.CheckBox; $week1.Checked = $true
$week2 = New-Object System.Windows.Forms.CheckBox; $week2.Checked = $true
$week3 = New-Object System.Windows.Forms.CheckBox; $week3.Checked = $true
foreach ($w in @($week1,$week2,$week3)) { $w.Add_CheckedChanged({ Update-MasterCheckboxState }) }
$script:weekCheckboxes = @($week1,$week2,$week3)
Update-MasterCheckboxState

$script:failures = @()
function Check($name, $cond) {
    if ($cond) { Write-Host "  OK $name" } else { $script:failures += $name; Write-Host "  FAIL $name" }
}

# System.Windows.Forms.CheckBox does not expose a public PerformClick() (unlike Button,
# where it's public) -- on CheckBox, ButtonBase.OnClick is only protected ("Family").
# Simulate a real user click via reflection so the CheckBox's own built-in CheckState
# auto-cycle (Unchecked -> Checked -> Indeterminate -> Unchecked) runs exactly as it
# would from a mouse click, followed by our Click handler -- same net effect as
# PerformClick() on a Button.
$script:onClickMethod = [System.Windows.Forms.CheckBox].GetMethod('OnClick', [System.Reflection.BindingFlags]'NonPublic,Instance')
function Invoke-Click($checkbox) {
    $script:onClickMethod.Invoke($checkbox, @([EventArgs]::Empty))
}

Check "all three checked -> master Checked" ($chkWeeksAll.CheckState -eq 'Checked')

$week2.Checked = $false
Check "one unchecked -> master Indeterminate" ($chkWeeksAll.CheckState -eq 'Indeterminate')

$week1.Checked = $false
$week3.Checked = $false
Check "all unchecked -> master Unchecked" ($chkWeeksAll.CheckState -eq 'Unchecked')

Invoke-Click $chkWeeksAll   # was Unchecked -> auto-cycles to Checked -> handler keeps it, selects all
Check "click while Unchecked selects all" (
    $week1.Checked -and $week2.Checked -and $week3.Checked -and $chkWeeksAll.CheckState -eq 'Checked'
)

Invoke-Click $chkWeeksAll   # was Checked -> auto-cycles to Indeterminate -> handler deselects all
Check "click while Checked deselects all" (
    (-not $week1.Checked) -and (-not $week2.Checked) -and (-not $week3.Checked) -and $chkWeeksAll.CheckState -eq 'Unchecked'
)

$week1.Checked = $true
$week2.Checked = $false
$week3.Checked = $true
Check "mixed selection -> master Indeterminate (again)" ($chkWeeksAll.CheckState -eq 'Indeterminate')

Invoke-Click $chkWeeksAll   # was Indeterminate -> auto-cycles to Unchecked -> handler selects all
Check "click while Indeterminate selects all (never gets stuck)" (
    $week1.Checked -and $week2.Checked -and $week3.Checked -and $chkWeeksAll.CheckState -eq 'Checked'
)

$week1.Checked = $false
$beforeWeek2 = $week2.Checked
$beforeWeek3 = $week3.Checked
$chkWeeksAll.CheckState = 'Indeterminate'   # programmatic set, NOT a click
Check "programmatic CheckState set does not trigger toggle-all (no Click fired)" (
    $week1.Checked -eq $false -and $week2.Checked -eq $beforeWeek2 -and $week3.Checked -eq $beforeWeek3
)

if ($script:failures.Count -gt 0) {
    Write-Host "FAILURES: $($script:failures -join ', ')"
    exit 1
} else {
    Write-Host "ALL PASS"
    exit 0
}
