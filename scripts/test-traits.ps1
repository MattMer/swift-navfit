param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/traits-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$catalogue=Get-Content (Join-Path $root 'src/trait-catalogue.json') -Raw | ConvertFrom-Json
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'TraitsHarness',(Join-Path $root 'tests/TraitsHarness.bas'))
    $app.DoCmd.RunCommand(126)
    if($app.Eval('TraitCatalogueHash()') -cne $catalogue.SourceSHA256){throw 'Embedded catalogue identity changed'}
    $catalogueChecks=1
    foreach($trait in $catalogue.Traits){
        $prefix='("'+$trait.FormKind+'",'+$trait.Block
        if($app.Eval('CatalogueTraitTitle'+$prefix+')') -cne $trait.Title){throw 'Trait title changed'}
        $catalogueChecks++
        if($app.Eval('CatalogueTraitDescription'+$prefix+',"")') -cne [string]$trait.GeneralDescription){throw 'Trait general description changed'}
        $catalogueChecks++
        foreach($mark in @('NOB','1.0','2.0','3.0','4.0','5.0')){
            if($app.Eval('CatalogueTraitDescription'+$prefix+',"'+$mark+'")') -cne [string]$trait.Marks.$mark){throw 'Trait mark description changed'}
            $catalogueChecks++
        }
    }
    $result=$app.Eval('TestTraits()')
    if(-not $result.StartsWith('PASS:')) {throw $result}
    $app.CloseCurrentDatabase(); $app.Quit(2); $app=$null
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    if(-not $app.Eval('VerifyTraitRecovery()')) {throw 'Trait recovery in new session failed'}
    $app.CloseCurrentDatabase()
    @{Status=$result;CatalogueChecks=$catalogueChecks;RecoveryInNewSession=$true;BuildSHA256=(Get-FileHash -LiteralPath $BuildPath).Hash;TestDatabase=$scratch} | ConvertTo-Json | Set-Content (Join-Path $root 'build/tests/traits-results.json')
    Write-Output $result
    Write-Output ('PASS: '+$catalogueChecks+' catalogue fidelity checks')
} finally {if($null -ne $app){$app.Quit(2)}}
