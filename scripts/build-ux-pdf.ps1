param([string]$OutputPath='D:\Development\Swift NAVFIT\build\ux-r26-blank.accdb',[string]$TemplatePath='D:\Development\Swift NAVFIT\build\ux-r25-candidate.accdb')
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$OutputPath=[IO.Path]::GetFullPath($OutputPath)
if(Test-Path -LiteralPath $OutputPath){throw "Refusing to overwrite $OutputPath"}
Copy-Item -LiteralPath $TemplatePath -Destination $OutputPath
$app=New-Object -ComObject Access.Application
try{
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($OutputPath,$true)
    foreach($module in @('modPDFImport','modPDFMapping','modUXHelp')){$app.LoadFromText(5,$module,(Join-Path $root ('src/'+$module+'.bas')))}
    $source=Get-Content (Join-Path $root 'scripts/build.ps1') -Raw
    Invoke-Expression ([regex]::Match($source,'(?s)    function Control\(.*?(?=    NewScreen ''frmSetup'')').Value)
    foreach($form in @('frmPDFImport','frmUXImportChoose')){$app.DoCmd.DeleteObject(2,$form)}
    $importSource=Get-Content (Join-Path $PSScriptRoot 'apply-ux-import-flow.ps1') -Raw
    Invoke-Expression ($importSource.Substring(0,$importSource.IndexOf('$app.DoCmd.OpenForm(''frmImportReview''')))
    $app.DoCmd.RunCommand(126)
    if($app.Eval('PDFMappingCount("EVAL")') -ne 40){throw 'PDF mapping did not execute'}
    $app.CloseCurrentDatabase()
    Write-Output "Compiled: $OutputPath"
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
