param([string]$OutputPath='D:\Development\Swift NAVFIT\build\ux-r27-navfit-candidate.accdb',[string]$TemplatePath='D:\Development\Swift NAVFIT\build\ux-r26-final.accdb')
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$OutputPath=[IO.Path]::GetFullPath($OutputPath);$TemplatePath=[IO.Path]::GetFullPath($TemplatePath)
if(Test-Path -LiteralPath $OutputPath){throw 'Refusing overwrite'}
Copy-Item -LiteralPath $TemplatePath -Destination $OutputPath
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($OutputPath,$true)
 foreach($m in @('modNAVFITMapping','modNAVFITAdapter','modNAVFITExchange','modImportReview','modUXNAVFIT','modSummaryGroups','modSwift','modUXHelp')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}
 $source=Get-Content (Join-Path $root 'scripts/build.ps1') -Raw
 Invoke-Expression ([regex]::Match($source,'(?s)    function Control\(.*?(?=    NewScreen ''frmSetup'')').Value)
 & (Join-Path $PSScriptRoot 'apply-ux-navfit.ps1')
 $app.DoCmd.RunCommand(126)
 $app.CloseCurrentDatabase()
 Write-Output ('Compiled NAVFIT candidate: '+$OutputPath)
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
