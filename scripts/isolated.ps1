# Run one script on the hidden test desktop (see workflow-common.ps1), so its Access
# windows stay off the screen and cannot take keystrokes or focus:
#   powershell.exe -NoProfile -File scripts/isolated.ps1 test-pdf-output -BuildPath <build>
# The first argument is a script in scripts/ (name or path); the rest are passed on.
# develop.ps1 and release-check.ps1 already isolate themselves.
$ErrorActionPreference='Stop'
if($args.Count -lt 1){throw 'Usage: isolated.ps1 <script> [arguments]'}
. (Join-Path $PSScriptRoot 'workflow-common.ps1')
. (Join-Path $PSScriptRoot 'build-retention.ps1')
$target=[string]$args[0]
if(-not $target.EndsWith('.ps1')){$target+='.ps1'}
if(-not [IO.Path]::IsPathRooted($target)){
    $inScripts=Join-Path $PSScriptRoot $target
    $target=if(Test-Path -LiteralPath $inScripts){$inScripts}else{[IO.Path]::GetFullPath($target)}
}
if(-not (Test-Path -LiteralPath $target)){throw ('Script not found: '+$target)}
$rest=@();if($args.Count -gt 1){$rest=@($args[1..($args.Count-1)]|ForEach-Object {[string]$_})}
# Clear expired test copies first, then record this run's copies with its result,
# so those from a passing run expire after a day like a full build's.
$root=Split-Path $PSScriptRoot -Parent
$inputs=@($rest|Where-Object {$_ -and -not $_.StartsWith('-') -and (Test-Path -LiteralPath $_)})
[void](InvokeSwiftScratchCleanup -Root $root -Apply -PreservePath $inputs)
$scratchRun=StartSwiftScratchRun $root
$exitCode=1
try{
    if($env:SWIFT_TEST_DESKTOP){
        & powershell.exe -NoProfile -File $target @rest
        $exitCode=$LASTEXITCODE
    }else{
        $exitCode=InvokeSwiftIsolated $target $rest
    }
}finally{
    try{CompleteSwiftScratchRun $scratchRun $(if($exitCode -eq 0){'PASS'}else{'FAIL'})}catch{Write-Warning ('Scratch retention record failed: '+$_.Exception.Message)}
}
exit $exitCode
