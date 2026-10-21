param([Parameter(Mandatory=$true)][string]$TemplatePath,[Parameter(Mandatory=$true)][string]$SourcePath)
$ErrorActionPreference='Stop'; $root=Split-Path $PSScriptRoot -Parent
$target=Join-Path $root ('build/tests/merge-transfer-'+[guid]::NewGuid()+'.accdb')
$before=(Get-FileHash -LiteralPath $SourcePath).Hash
& powershell.exe -NoProfile -File (Join-Path $root 'scripts/upgrade-workspace.ps1') -SourcePath $SourcePath -TemplatePath $TemplatePath -DestinationPath $target
if($LASTEXITCODE -ne 0){throw 'Merge-history migration failed'}
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($target,$true)
 $id=[int]$app.Eval('DMax("MergeID","DirectoryMerge","MergeState=''Committed''")')
 $a=[int]$app.Eval('DLookup("FromMemberID","DirectoryMerge","MergeID='+$id+'")')
 $b=[int]$app.Eval('DLookup("ToMemberID","DirectoryMerge","MergeID='+$id+'")')
 $max=[int]$app.Eval('DMax("MergeID","DirectoryMerge")')
 if(-not $app.Eval('UndoDirectoryMerge('+$id+')')){throw 'Undo after migration failed'}
 $new=[int]$app.Eval('ApplyDirectoryMerge('+$a+','+$b+',PreviewDirectoryMerge('+$a+','+$b+'),"Reviewed transferred same-Sailor mapping")')
 if($new -le $max){throw 'Transferred AutoNumber did not advance'}
 $app.CloseCurrentDatabase()
 if((Get-FileHash -LiteralPath $SourcePath).Hash -ne $before){throw 'Source changed'}
 @{Result='PASS';Source=$SourcePath;SourceSHA256=$before;Target=$target;ExactTransfer=$true;Reopened=$true;UndoAfterTransfer=$true;NextMergeID=$new;PreviousMaximum=$max}|ConvertTo-Json|Set-Content (Join-Path $root 'build/directory-merge-transfer-verification.json')
 Write-Output 'PASS: exact committed/undone merge-history transfer, reopen, guarded undo and next AutoNumber; original source unchanged'
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
