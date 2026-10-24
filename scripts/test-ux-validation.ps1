param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$InstallSource)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/ux-validation-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
Write-Output ('Controller: live validation runner; disposable database: '+$scratch)
$app=New-Object -ComObject Access.Application
try{
$app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
if($InstallSource){
 $app.LoadFromText(5,'modUXValidation',(Join-Path $root 'src/modUXValidation.bas'))
 $app.DoCmd.OpenForm('frmUXValidation',1);$f=$app.Forms.Item('frmUXValidation')
 $f.Controls('lblCount').Height=300;$f.Controls('txtEvidence').Top=5450
 foreach($entry in @(@('lblIssues','lstIssues','&Findings',2200,180),@('lblEvidence','txtEvidence','&Evidence and source',5200,200))){
  $c=$app.CreateControl($f.Name,100,0,$entry[1],'',400,$entry[3],12600,$entry[4]);$c.Name=$entry[0];$c.Caption=$entry[2];$c.FontName='Segoe UI';$c.FontSize=9
 }
 $i=0;foreach($name in @('lstIssues','txtEvidence','cmdOpen','cmdPrevious','cmdNext','cmdBack')){$f.Controls($name).TabIndex=$i;$i++}
 $app.DoCmd.Close(2,$f.Name,1)
}
$app.LoadFromText(5,'UXValidationHarness',(Join-Path $root 'tests/UXValidationHarness.bas'))
$app.DoCmd.RunCommand(126)
$result=$app.Eval('TestUXValidation()')
if(-not $result.StartsWith('PASS:')){throw $result}
Write-Output $result
}finally{try{[void]$app.Eval('UXLeaveEditor("Discard")')}catch{};$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
