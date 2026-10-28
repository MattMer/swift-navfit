param([string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$catalogue=Get-Content (Join-Path $root 'src/field-definitions.json') -Raw|ConvertFrom-Json
$fields=@($catalogue.fields|Sort-Object ordinal)
$checks=0
function Check([bool]$ok,[string]$message){if(-not $ok){throw $message};$script:checks++}
Check ($catalogue.catalogueVersion -eq 2) 'Unsupported catalogue version'
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
    if($layout.Name -eq 'CHIEFEVAL:0'){
        foreach($row in $layout.Value.fields|Where-Object {$_.key -notmatch '^Trait3[3-9](Comment[123])?$' -and $_.key -ne '@guidance'}){
            $current=@($actual.fields|Where-Object key -eq $row.key)
            Check ($current.Count -eq 1 -and ($current[0]|ConvertTo-Json -Compress) -ceq ($row|ConvertTo-Json -Compress)) ('Unrelated Chief field moved: '+$row.key)
        }
        Check ($actual.height -eq 30500 -and @($actual.fields|Where-Object key -eq '@guidance').Count -eq 0) 'Chief trait groups require the fixed tall canvas and per-trait guidance'
    }else{Check (($actual|ConvertTo-Json -Depth 10 -Compress) -ceq ($layout.Value|ConvertTo-Json -Depth 10 -Compress)) ('Fixed geometry changed: '+$layout.Name)}
}
$manifest=Get-Content (Join-Path $root 'scripts/build-manifest.json') -Raw|ConvertFrom-Json
Check ($manifest.inputs -contains 'src/field-definitions.json') 'Canonical catalogue is not fingerprinted'
$chief=@($fields|Where-Object {$_.key -match '^Trait(3[3-9])Comment([123])$'})
Check ($chief.Count -eq 21) 'Chief comments must contain three independent fields per trait'
$prefixes=@('LEAD','TAC','PROF','MIS','EO','TEAM','MIL')
foreach($field in $chief){
    [void]($field.key -match '^Trait(3[3-9])Comment([123])$');$block=[int]$Matches[1];$entry=[int]$Matches[2]
    Check ($field.storageType -eq 'LONGTEXT' -and $field.definitionVersion -eq 2 -and $field.multiline -and ($field.forms -join '|') -eq 'CHIEFEVAL') ('Chief comment storage/applicability differs: '+$field.key)
    Check ($field.bindings.CHIEFEVAL.navfit[0].source -ceq ($prefixes[$block-33]+'DN'+$entry)) ('Chief native comment order differs: '+$field.key)
    Check ($field.bindings.EVAL.navfit.Count -eq 0 -and $field.bindings.FITREP.navfit.Count -eq 0) ('Chief comments leaked into another native kind: '+$field.key)
    $row=$catalogue.fixedLayout.'CHIEFEVAL:0'.fields|Where-Object key -eq $field.key
    Check ($row.y -eq 10380+($block-33)*2800+250+($entry-1)*720 -and $row.x -eq 4932 -and $row.w -eq 6876 -and $row.h -eq 580) ('Chief comment must match the printed width/height in its trait group: '+$field.key)
}
# Forms are source: a catalogue change must include the form edit. Every field's
# control exists in the editor and its surface; runtime layout comes from modUXFields.
foreach($formName in @('frmEditor','frmUXEditorFields')){
    $text=[IO.File]::ReadAllText((Join-Path $root ('src/forms/'+$formName+'.form.txt')))
    $controls=@([regex]::Matches($text,'(?m)^\s+Name ="([^"]+)"')|ForEach-Object {$_.Groups[1].Value})
    # The report type is chosen when a report is created; the surface never edits it.
    foreach($field in $fields|Where-Object {-not ($formName -eq 'frmUXEditorFields' -and $_.key -eq 'FormKind')}){Check ($controls -contains $field.control) ('Catalogue field has no control in src/forms/'+$formName+'.form.txt: '+$field.key+' ('+$field.control+')')}
}
Write-Output ('PASS: '+$checks+' canonical definitions, exact adapter bindings, fixed-geometry and form-control checks')
if($BuildPath){
    $before=(Get-FileHash -LiteralPath $BuildPath).Hash
    $fixture=Join-Path $root ('build/tests/catalogue-extension-'+[guid]::NewGuid())
    [void](New-Item -ItemType Directory -Path (Join-Path $fixture 'src'))
    [void](New-Item -ItemType Directory -Path (Join-Path $fixture 'build'))
    $catalogue.catalogueVersion++
    $extra=[pscustomobject]@{key='CatalogueTestValue';ordinal=$fields.Count;definitionVersion=$catalogue.catalogueVersion;storageType='TEXT(40)';label='Synthetic catalogue field';control='txtCatalogueTestValue';recoveryKey='CatalogueTestValue';section=7;block='';limit=40;date=$false;digits=$null;options=$null;multiline=$false;bindings=[pscustomobject]@{EVAL=$null;CHIEFEVAL=$null;FITREP=$null}}
    $extraDate=[pscustomobject]@{key='CatalogueTestDate';ordinal=($fields.Count+1);definitionVersion=$catalogue.catalogueVersion;storageType='DATETIME';label='Synthetic catalogue date';control='txtCatalogueTestDate';recoveryKey='CatalogueTestDate';section=7;block='';limit=10;date=$true;digits=$null;options=$null;multiline=$false;bindings=[pscustomobject]@{EVAL=$null;CHIEFEVAL=$null;FITREP=$null}}
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
        . (Join-Path (Split-Path $PSScriptRoot -Parent) 'apply-ux-field-catalogue.ps1')
        . (Join-Path (Split-Path $PSScriptRoot -Parent) 'apply-ux-surface-catalogue.ps1')
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
