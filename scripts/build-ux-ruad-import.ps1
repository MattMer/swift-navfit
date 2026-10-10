param([Parameter(Mandatory=$true)][string]$TemplatePath,[Parameter(Mandatory=$true)][string]$OutputPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$OutputPath=[IO.Path]::GetFullPath($OutputPath)
if(Test-Path -LiteralPath $OutputPath){throw 'Refusing overwrite'}
Copy-Item -LiteralPath $TemplatePath -Destination $OutputPath
$app=New-Object -ComObject Access.Application
try {
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($OutputPath,$true)
 foreach($opened in @($app.Forms|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$opened,2)}
 $source=Get-Content (Join-Path $root 'scripts/build.ps1') -Raw
 Invoke-Expression ([regex]::Match($source,'(?s)    function Control\(.*?(?=    NewScreen ''frmSetup'')').Value)
 . (Join-Path $PSScriptRoot 'apply-ux-ruad-import.ps1')
 $app.DoCmd.RunCommand(126)
 'Compiled RUAD build: '+$OutputPath
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
