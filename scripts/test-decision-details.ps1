param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/decision-details-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'DecisionDetailsHarness',(Join-Path $root 'tests/DecisionDetailsHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestDecisionDetails()')
    if(-not $result.StartsWith('PASS:')){throw $result}
    [void]$app.Eval('OpenScreen("frmEditor")')
    $editor=$app.Forms.Item('frmEditor')
    $button=$editor.Controls.Item('cmdDecisionDetails')
    foreach($control in $editor.Controls){
        if($button.Visible -and $control.Visible -and $control.ControlType -ne 101 -and $control.Name -ne $button.Name){
            if($button.Left -lt ($control.Left+$control.Width) -and ($button.Left+$button.Width) -gt $control.Left -and $button.Top -lt ($control.Top+$control.Height) -and ($button.Top+$button.Height) -gt $control.Top){throw ('New button overlaps '+$control.Name)}
        }
    }
    [void]$app.Eval('OpenDecisionDetails()')
    foreach($formName in @('frmEditor','frmDecisionDetails')){
        $form=$app.Forms.Item($formName)
        foreach($control in $form.Controls){
            if($control.Visible -and ($control.Left -lt 0 -or ($control.Left+$control.Width) -gt $form.Width -or $control.Top -lt 0 -or ($control.Top+$control.Height) -gt $form.Section(0).Height)){throw ('Control outside form: '+$control.Name)}
        }
    }
    [void]$app.Eval('CancelDecisionDetails()')
    $app.DoCmd.Close(2,'frmEditor',2)
    $app.CloseCurrentDatabase(); $app.Quit(2); $app=$null
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    if(-not $app.Eval('VerifyDecisionDetailsRecovery()')){throw 'New-session decision details recovery failed'}
    $app.CloseCurrentDatabase()
    @{Status=$result;LayoutChecks=3;RecoveryInNewSession=$true;BuildSHA256=(Get-FileHash -LiteralPath $BuildPath).Hash;TestDatabase=$scratch} | ConvertTo-Json | Set-Content (Join-Path $root 'build/tests/decision-details-results.json')
    Write-Output $result
} finally {if($null -ne $app){$app.Quit(2)}}
