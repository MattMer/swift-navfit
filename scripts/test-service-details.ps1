param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/service-details-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'ServiceDetailsHarness',(Join-Path $root 'tests/ServiceDetailsHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestServiceDetails()')
    if(-not $result.StartsWith('PASS:')){throw $result}
    $reference=Get-Content (Join-Path $root 'src/service-reference.json') -Raw | ConvertFrom-Json
    [void]$app.Eval('OpenScreen("frmEditor")')
    [void]$app.Eval('OpenServiceDetails()')
    $referenceChecks=0
    foreach($spec in @(@{Control='txtReadinessDetail';Values=$reference.Readiness},@{Control='txtBilletSubcategory';Values=$reference.Billet})){
        $actual=[string]$app.Forms.Item('frmServiceDetails').Controls.Item($spec.Control).RowSource
        $entries=$actual.Split(';')
        if($entries.Count -ne $spec.Values.Count){throw 'Reference suggestion count changed'}
        $referenceChecks++
        for($i=0;$i -lt $entries.Count;$i++){
            if($entries[$i] -cne $spec.Values[$i]){throw 'Reference suggestion text changed'}
            $referenceChecks++
        }
    }
    [void]$app.Eval('CancelServiceDetails()')
    $app.DoCmd.Close(2,'frmEditor',2)
    $app.CloseCurrentDatabase(); $app.Quit(2); $app=$null
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    if(-not $app.Eval('VerifyServiceDetailsRecovery()')){throw 'New-session service details recovery failed'}
    $app.CloseCurrentDatabase()
    @{Status=$result;ReferenceChecks=$referenceChecks;RecoveryInNewSession=$true;BuildSHA256=(Get-FileHash -LiteralPath $BuildPath).Hash;TestDatabase=$scratch} | ConvertTo-Json | Set-Content (Join-Path $root 'build/tests/service-details-results.json')
    Write-Output $result
    Write-Output ('PASS: '+$referenceChecks+' readiness/billet reference checks')
} finally {if($null -ne $app){$app.Quit(2)}}
