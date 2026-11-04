param([switch]$Apply,[string[]]$PreservePath=@())
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'build-retention.ps1')
InvokeSwiftScratchCleanup -Root (Split-Path $PSScriptRoot -Parent) -Apply:$Apply -PreservePath $PreservePath | ConvertTo-Json -Depth 5
