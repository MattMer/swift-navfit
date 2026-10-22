param([string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$catalogue=Get-Content (Join-Path $root 'src/field-definitions.json') -Raw|ConvertFrom-Json
$fields=@($catalogue.fields|Sort-Object ordinal)
$checks=0
function Check([bool]$ok,[string]$message){if(-not $ok){throw $message};$script:checks++}
Check ($catalogue.catalogueVersion -eq 1) 'Unsupported catalogue version'
Check (@($fields.key|Sort-Object -Unique).Count -eq $fields.Count) 'Duplicate field keys'
Check (@($fields.control|Sort-Object -Unique).Count -eq $fields.Count) 'Duplicate editor controls'
for($i=0;$i -lt $fields.Count;$i++){
    $field=$fields[$i]
    Check ($field.ordinal -eq $i) ('Noncontiguous field ordinal: '+$field.key)
    Check ($field.storageType -match '^(TEXT\([1-9][0-9]*\)|LONGTEXT|DATETIME)$') ('Invalid field type: '+$field.key)
    Check ($field.definitionVersion -ge 1 -and $field.definitionVersion -le $catalogue.catalogueVersion) ('Invalid version: '+$field.key)
    Check ($field.label.Length -gt 0 -and $field.control.Length -gt 0) ('Missing editor metadata: '+$field.key)
}
# Retained interchange/geometry catalogues are compatibility evidence, not build
# authorities. Compare exact order and coordinates before accepting this cutover.
foreach($adapter in @('navfit','pdf')){
    $old=Get-Content (Join-Path $root ('src/'+$adapter+'-field-map.json')) -Raw|ConvertFrom-Json
    foreach($kind in @('EVAL','CHIEFEVAL','FITREP')){
        $rows=@(foreach($field in $fields){foreach($binding in $field.bindings.$kind.$adapter){[pscustomobject]@{ordinal=$binding.ordinal;row=($field.key+'|'+$binding.source+'|'+$binding.conversion)}}})
        $actual=@($rows|Sort-Object ordinal|ForEach-Object {$_.row})
        $expected=@($old.$kind|ForEach-Object {$_.target+'|'+$_.source+'|'+$_.conversion})
        Check (($actual -join "`n") -ceq ($expected -join "`n")) ('Interchange bindings changed: '+$adapter+' '+$kind)
    }
}
$oldPages=Get-Content (Join-Path $root 'src/report-page-map.json') -Raw|ConvertFrom-Json
foreach($kind in @('EVAL','CHIEFEVAL','FITREP')){
    foreach($field in $fields){
        $oldPage=$oldPages.$kind.($field.key)
        $page=$field.bindings.$kind.page
        Check (($null -eq $oldPage -and $null -eq $page) -or ($null -ne $oldPage -and $null -ne $page -and $oldPage.page -eq $page.page -and $oldPage.block -eq $page.block)) ('Page/block mapping changed: '+$kind+' '+$field.key)
    }
}
$oldLayouts=Get-Content (Join-Path $root 'src/report-page-layout.json') -Raw|ConvertFrom-Json
foreach($layout in $oldLayouts.PSObject.Properties){
    $actual=$catalogue.fixedLayout.($layout.Name)
    Check (($actual|ConvertTo-Json -Depth 10 -Compress) -ceq ($layout.Value|ConvertTo-Json -Depth 10 -Compress)) ('Fixed geometry changed: '+$layout.Name)
}
$manifest=Get-Content (Join-Path $root 'scripts/build-manifest.json') -Raw|ConvertFrom-Json
Check ($manifest.inputs -contains 'src/field-definitions.json') 'Canonical catalogue is not fingerprinted'
Write-Output ('PASS: '+$checks+' canonical definitions, exact adapter bindings and fixed-geometry checks')
if($BuildPath){
    $before=(Get-FileHash -LiteralPath $BuildPath).Hash
    $fixture=Join-Path $root ('build/tests/catalogue-extension-'+[guid]::NewGuid())
    [void](New-Item -ItemType Directory -Path (Join-Path $fixture 'src'))
    [void](New-Item -ItemType Directory -Path (Join-Path $fixture 'build'))
    $catalogue.catalogueVersion=2
    $extra=[pscustomobject]@{key='CatalogueTestValue';ordinal=$fields.Count;definitionVersion=2;storageType='TEXT(40)';label='Synthetic catalogue field';control='txtCatalogueTestValue';recoveryKey='CatalogueTestValue';section=7;block='';limit=40;date=$false;digits=$null;options=$null;multiline=$false;bindings=[pscustomobject]@{EVAL=$null;CHIEFEVAL=$null;FITREP=$null}}
    $extraDate=[pscustomobject]@{key='CatalogueTestDate';ordinal=($fields.Count+1);definitionVersion=2;storageType='DATETIME';label='Synthetic catalogue date';control='txtCatalogueTestDate';recoveryKey='CatalogueTestDate';section=7;block='';limit=10;date=$true;digits=$null;options=$null;multiline=$false;bindings=[pscustomobject]@{EVAL=$null;CHIEFEVAL=$null;FITREP=$null}}
    $catalogue.fields+=@($extra,$extraDate)
    $catalogue|ConvertTo-Json -Depth 40|Set-Content (Join-Path $fixture 'src/field-definitions.json') -Encoding UTF8
    foreach($file in @('trait-catalogue.json','modUX.bas')){Copy-Item -LiteralPath (Join-Path $root ('src/'+$file)) -Destination (Join-Path $fixture ('src/'+$file))}
    $scratch=Join-Path $fixture 'synthetic.accdb'
    Copy-Item -LiteralPath $BuildPath -Destination $scratch
    $app=New-Object -ComObject Access.Application
    $projectRoot=$root
    try {
        $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
        $db=$app.CurrentDb()
        foreach($field in @($extra,$extraDate)){
            $db.Execute(('ALTER TABLE Report ADD COLUMN ['+$field.key+'] '+$field.storageType),128)
            $db.Execute(('ALTER TABLE RecoveryDraft ADD COLUMN ['+$field.recoveryKey+'] LONGTEXT'),128)
        }
        $db.Close()
        [void]$app.Eval('ConfigureWorkspace("SYNTHETIC CATALOGUE EXTENSION","Tester")')
        [void]$app.Eval('CommitDraft(0,"EVAL","Catalogue extension","Synthetic",Null,Null,"")')
        # Exercise the production metadata generators against a disposable catalogue.
        $root=$fixture;$script:swiftManagedBuild=$true
        . (Join-Path $PSScriptRoot 'apply-ux-field-catalogue.ps1')
        . (Join-Path $PSScriptRoot 'apply-ux-surface-catalogue.ps1')
        $root=$projectRoot
        $app.DoCmd.OpenForm('frmEditor',1)
        foreach($field in @($extra,$extraDate)){
            $control=$app.CreateControl('frmEditor',109,0,'','',0,0,100,100)
            $control.Name=$field.control;$control.Visible=$false
        }
        $app.DoCmd.Close(2,'frmEditor',1)
        $app.LoadFromText(5,'PersistenceHarness',(Join-Path $root 'tests/PersistenceHarness.bas'))
        $app.DoCmd.RunCommand(126)
        $result=$app.Eval('TestAdditionalField()');Write-Output $result
        if($result -notlike 'PASS:*'){throw $result}
        $app.CloseCurrentDatabase();$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
        $app=New-Object -ComObject Access.Application
        $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
        if(-not $app.Eval('VerifyAdditionalRecovery()')){throw 'New-session catalogue field recovery failed'}
        Write-Output 'PASS: added catalogue field recovered in a new Access session'
    } finally {
        $root=$projectRoot;$script:swiftManagedBuild=$false
        if($app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
        if((Get-FileHash -LiteralPath $BuildPath).Hash -ne $before){throw 'Catalogue test changed the supplied build'}
    }
}
