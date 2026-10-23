param([Parameter(Mandatory=$true)][string]$BuildPath,[Parameter(Mandatory=$true)][string]$CompiledPath,[switch]$CheckNormalSecurity)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$BuildPath=[IO.Path]::GetFullPath($BuildPath);$CompiledPath=[IO.Path]::GetFullPath($CompiledPath)
if([IO.Path]::GetExtension($BuildPath) -ne '.accdb' -or [IO.Path]::GetExtension($CompiledPath) -ne '.accde'){throw 'Use a blank ACCDB source and its compiled ACCDE candidate'}
$originals=@{}
foreach($path in @($BuildPath,$CompiledPath)){$originals[$path]=(Get-FileHash -LiteralPath $path).Hash}
$folder=Join-Path $root ('build/tests/compiled-'+[guid]::NewGuid())
[void](New-Item -ItemType Directory -Path $folder)
$script:checks=0;$script:app=$null;$script:cleanupFailure=''
function AssertCompiled([bool]$Condition,[string]$Message){
    if(-not $Condition){throw $Message}
    $script:checks++
}
function CloseOwnedCompiledSession {
    if($script:app){
        try{$script:app.Quit(2)}catch{$script:cleanupFailure=$_.Exception.Message;Write-Warning ('Owned test cleanup: '+$script:cleanupFailure)}
        try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:app)}catch{Write-Warning ('Owned test release: '+$_.Exception.Message)}
        $script:app=$null
    }
}
function OpenCompiledSession([string]$Path,[switch]$NormalSecurity){
    CloseOwnedCompiledSession
    if($script:cleanupFailure){throw ('Previous controller did not close normally: '+$script:cleanupFailure)}
    Write-Output ('Controller: compiled-distribution runner; database: '+$Path)
    $script:app=New-Object -ComObject Access.Application
    if(-not $NormalSecurity){$script:app.AutomationSecurity=1}
    $script:app.OpenCurrentDatabase($Path,$true)
}
function CheckCompiled {
    $db=$app.CurrentDb()
    try{AssertCompiled ($db.Properties.Item('MDE').Value -eq 'T') 'Candidate is not genuinely compiled'}finally{$db.Close();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($db)}
    AssertCompiled (@($app.References|Where-Object {$_.IsBroken}).Count -eq 0) 'Compiled candidate has a broken reference'
}
try {
    # Seed reports through production services and add explicitly synthetic retained
    # import evidence. No harness is injected into the compiled binary, and the
    # supplied ACCDB/candidate remain untouched.
    $source=Join-Path $folder 'synthetic-source.accdb'
    Copy-Item -LiteralPath $BuildPath -Destination $source
    OpenCompiledSession $source
    AssertCompiled ([bool]$app.Eval('ConfigureWorkspace("COMPILED TEST - SYNTHETIC","Compiled Tester")')) 'Synthetic setup failed'
    $ids=@()
    foreach($kind in @('EVAL','CHIEFEVAL','FITREP')){
        $id=$app.Eval('CommitDraft(0,"'+$kind+'","SYNTHETIC '+$kind+'","TEST '+$kind+'",#2026-01-01#,#2026-09-30#,"Private synthetic notes","Compiled test seed","SYNTHETIC PRINT TEXT")')
        AssertCompiled ([int]$id -gt 0) ('Cannot seed '+$kind)
        $ids+= [int]$id
    }
    $receiptGuid=[guid]::NewGuid().ToString('B')
    $db=$app.CurrentDb()
    try{
        $author=[int]$app.Eval('DMin("UserProfileID","UserProfile")')
        $db.Execute("INSERT INTO ImportReceipt (ReceiptGuid,SourceHash,SourcePath,ExternalReportID,ReportID,AuthorProfileID) VALUES ('$receiptGuid','$('A'*64)','SYNTHETIC provenance',1,$($ids[0]),$author)",128)
        $receiptID=[int]$app.Eval('DMax("ImportReceiptID","ImportReceipt")')
        $db.Execute("INSERT INTO ImportRawValue (ImportReceiptID,SourceTable,SourceRow,FieldOrdinal,FieldKey,DAOType,RawValue,IsNullValue) VALUES ($receiptID,'SYNTHETIC',1,1,'OfficialComments',12,'SYNTHETIC ORIGINAL',False)",128)
    }finally{$db.Close();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($db)}
    CloseOwnedCompiledSession
    $sourceHash=(Get-FileHash -LiteralPath $source).Hash
    # Explicit attended acceptance only: a signing-only Office policy can display
    # blocking security dialogs. Never bypass/change that policy automatically.
    if($CheckNormalSecurity){
        $launch=Join-Path $folder 'normal-security.accde'
        Copy-Item -LiteralPath $CompiledPath -Destination $launch
        OpenCompiledSession $launch -NormalSecurity
        if(-not $app.CurrentProject.FullName){throw 'Normal Office security refused the compiled file; signing/trust acceptance is required'}
        $expected=if($app.CurrentProject.IsTrusted){'frmSetup'}else{'frmEnableContent'}
        AssertCompiled ($app.CurrentProject.AllForms.Item($expected).IsLoaded) ('Normal-security compiled startup did not show '+$expected)
        CloseOwnedCompiledSession
    }
    $target=Join-Path $folder 'working.accde'
    Copy-Item -LiteralPath $CompiledPath -Destination $target
    OpenCompiledSession $target
    CheckCompiled
    AssertCompiled ($app.CurrentProject.AllForms.Item('frmSetup').IsLoaded) 'Blank compiled startup did not show setup'
    AssertCompiled ($app.Eval('DCount("*","Report")') -eq 0) 'Compiled blank contains reports'
    $blocked=$false
    try{$app.SaveAsText(5,'modSwift',(Join-Path $folder 'unexpected-source.txt'))}catch{$blocked=$true}
    AssertCompiled $blocked 'Editable VBA source could be exported from the candidate'
    AssertCompiled ([bool]$app.Eval('UXPreviewSwiftWorkspace("'+$source.Replace('"','""')+'")')) 'Settings import preview failed in ACCDE'
    AssertCompiled ([bool]$app.Eval('UXApplySwiftWorkspace()')) 'Settings import failed in ACCDE'
    AssertCompiled ($app.Eval('DCount("*","Report")') -eq 3) 'Compiled import lost reports'
    AssertCompiled ($app.Eval('DCount("*","ImportReceipt")') -gt 0) 'Compiled import lost provenance'
    AssertCompiled ($app.Eval('DLookup("ReceiptGuid","ImportReceipt")') -eq $receiptGuid) 'Compiled import changed provenance identity'
    AssertCompiled ($app.Eval('DLookup("RawValue","ImportRawValue")') -eq 'SYNTHETIC ORIGINAL') 'Compiled import changed original raw evidence'
    AssertCompiled ((Get-FileHash -LiteralPath $source).Hash -eq $sourceHash) 'Import modified the source workspace'
    CloseOwnedCompiledSession
    OpenCompiledSession $target
    AssertCompiled ($app.CurrentProject.AllForms.Item('frmContributorStart').IsLoaded) 'Imported compiled reopen did not require contributor selection'
    $contributor=[int]$app.Eval('DMin("UserProfileID","UserProfile")')
    [void]$app.Eval('SwitchContributor('+$contributor+')')
    foreach($id in $ids){
        $app.DoCmd.OpenForm('frmEditor',0,[Type]::Missing,[Type]::Missing,[Type]::Missing,[Type]::Missing,[string]$id)
        AssertCompiled ([bool]$app.Eval('UXSelectSection(5)')) 'Compiled report page failed'
        AssertCompiled ([bool]$app.Eval('UXLeaveEditor("Discard")')) 'Compiled editor close failed'
    }
    $revision=[int]$app.Eval('CurrentRevision('+$ids[0]+')')
    $new=$app.Eval('CommitDraft('+$ids[0]+',"EVAL","SYNTHETIC UPDATED","TEST EVAL",#2026-01-01#,#2026-09-30#,"Private synthetic notes","Compiled save","SYNTHETIC UPDATED TEXT")')
    AssertCompiled ([int]$new -gt 0 -and [int]$app.Eval('CurrentRevision('+$ids[0]+')') -ne $revision) 'Compiled save did not create a revision'
    CloseOwnedCompiledSession
    OpenCompiledSession $target
    AssertCompiled ($app.Eval('DCount("*","ReportRevision")') -eq 4) 'Compiled revisions did not persist'
    AssertCompiled ($app.Eval('DLookup("OfficialComments","Report","ReportID='+$ids[0]+'")') -eq 'SYNTHETIC UPDATED TEXT') 'Compiled saved text did not persist'
    CloseOwnedCompiledSession
    # Upgrade from an actual populated ACCDE, through the same production UI
    # handlers. DAO reads data without importing or executing source code.
    $compiledSourceHash=(Get-FileHash -LiteralPath $target).Hash
    $upgraded=Join-Path $folder 'upgrade.accde'
    Copy-Item -LiteralPath $CompiledPath -Destination $upgraded
    OpenCompiledSession $upgraded
    AssertCompiled ([bool]$app.Eval('UXPreviewSwiftWorkspace("'+$target.Replace('"','""')+'")')) 'Compiled-source preview failed'
    AssertCompiled ([bool]$app.Eval('UXApplySwiftWorkspace()')) 'Compiled-source import failed'
    AssertCompiled ($app.Eval('DCount("*","ReportRevision")') -eq 4) 'Compiled-source upgrade lost history'
    AssertCompiled ($app.Eval('DLookup("ReceiptGuid","ImportReceipt")') -eq $receiptGuid) 'Compiled-source upgrade lost provenance'
    AssertCompiled ($app.Eval('DLookup("RawValue","ImportRawValue")') -eq 'SYNTHETIC ORIGINAL') 'Compiled-source upgrade lost raw evidence'
    CloseOwnedCompiledSession
    AssertCompiled ((Get-FileHash -LiteralPath $target).Hash -eq $compiledSourceHash) 'Upgrade modified its populated ACCDE source'
    foreach($path in $originals.Keys){AssertCompiled ((Get-FileHash -LiteralPath $path).Hash -eq $originals[$path]) 'A supplied artifact changed'}
} catch {$primary=$_;throw $primary}
finally {CloseOwnedCompiledSession}
if($script:cleanupFailure){throw ('Owned test session cleanup failed: '+$script:cleanupFailure)}
[pscustomobject]@{Status='PASS';Assertions=$checks;Source=$BuildPath;SourceSHA256=$originals[$BuildPath];Compiled=$CompiledPath;CompiledSHA256=$originals[$CompiledPath];NormalSecurityTested=[bool]$CheckNormalSecurity;PDFOutputTested=$false;Scope='Trusted development automation: native compiled startup, protected source, three report types, save/reopen, ACCDB and ACCDE Settings import, history/provenance and source hashes; current schema and installed desktop Access only. Normal Office security/signing acceptance is separate unless explicitly tested; compiled PDF output requires further verification.'}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $folder 'verification.json')
Write-Output ('PASS: '+$checks+' native compiled-distribution assertions; evidence: '+$folder)
