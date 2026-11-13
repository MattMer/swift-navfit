# Run one script on the hidden test desktop (see workflow-common.ps1), so its Access
# windows stay off the screen and cannot take keystrokes or focus:
#   powershell.exe -NoProfile -File scripts/isolated.ps1 test-pdf-output -BuildPath <build>
# The first argument is a script in scripts/ (name or path); the rest are passed on.
# develop.ps1 and release-check.ps1 already isolate themselves.
$ErrorActionPreference='Stop'
if($args.Count -lt 1){throw 'Usage: isolated.ps1 <script> [arguments]'}
. (Join-Path $PSScriptRoot 'workflow-common.ps1')
$target=[string]$args[0]
if(-not $target.EndsWith('.ps1')){$target+='.ps1'}
if(-not [IO.Path]::IsPathRooted($target)){
    $inScripts=Join-Path $PSScriptRoot $target
    $target=if(Test-Path -LiteralPath $inScripts){$inScripts}else{[IO.Path]::GetFullPath($target)}
}
if(-not (Test-Path -LiteralPath $target)){throw ('Script not found: '+$target)}
$rest=@();if($args.Count -gt 1){$rest=@($args[1..($args.Count-1)]|ForEach-Object {[string]$_})}
if($env:SWIFT_TEST_DESKTOP){
    & powershell.exe -NoProfile -File $target @rest
    exit $LASTEXITCODE
}
exit (InvokeSwiftIsolated $target $rest)
