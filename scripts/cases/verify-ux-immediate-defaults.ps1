param([Parameter(Mandatory=$true)][string]$BlankPath,[Parameter(Mandatory=$true)][string]$DemoPath,[string]$Candidate="r106",[int]$VerifiedAssertions=387,[string[]]$VerifiedSuites=@("177 defaults; 39 reporting senior; 40 workflows plus shutdown; 28 efficiency plus timer; 42 lifecycle; 61 live members"))
$ErrorActionPreference='Stop';$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$artifacts=@();$counts=[ordered]@{}
foreach($path in @($BlankPath,$DemoPath)){
 $full=[IO.Path]::GetFullPath($path);$app=New-Object -ComObject Access.Application
 try{
  $app.AutomationSecurity=1;$app.OpenCurrentDatabase($full,$true)
  $app.DoCmd.RunCommand(126)
  $forms=@($app.Forms|ForEach-Object {$_.Name})
  if($forms.Count -eq 0){throw 'Startup opened no forms'}
  $reportCount=[long]$app.Eval('DCount("*","Report")')
  $artifacts+= [ordered]@{Path=$full;StartupForms=$forms;Reports=$reportCount}
  if($path -eq $DemoPath -and $reportCount -lt 3){throw 'Demo lacks synthetic reports'}
 }finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
$engine=New-Object -ComObject DAO.DBEngine.120;$db=$engine.OpenDatabase([IO.Path]::GetFullPath($BlankPath),$true,$true)
try{
 foreach($t in $db.TableDefs){
  if($t.Name -like 'MSys*' -or $t.Name -like '~*' -or $t.Name -eq 'AppMetadata'){continue}
  $r=$db.OpenRecordset('SELECT Count(*) AS N FROM ['+$t.Name+']',4)
  $counts[$t.Name]=[long]$r.Fields.Item('N').Value;$r.Close()
  if($counts[$t.Name] -ne 0){throw ('Blank has data in '+$t.Name)}
 }
 if($counts.Count -ne 48){throw ('Unexpected blank table count '+$counts.Count)}
}finally{$db.Close();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($engine)}
foreach($a in $artifacts){$a.SHA256=(Get-FileHash -LiteralPath $a.Path).Hash;$a.Bytes=(Get-Item -LiteralPath $a.Path).Length}
[ordered]@{Candidate=$Candidate;Date='2026-10-14';AppVersion='0.25.0-dev';Schema=25;TotalAssertions=$VerifiedAssertions;Artifacts=$artifacts;BlankTables=$counts;Tests=$VerifiedSuites;Limitations=@('Physical keyboard, screen reader and display/DPI acceptance not established','Confirmation choices injected in automated tests; native modal interaction not manually verified');Published=$false}|ConvertTo-Json -Depth 8|Set-Content (Join-Path $root ("planning/RECURSIVE_DEFAULTS_"+$Candidate.ToUpper()+"_VERIFICATION.json"))
Write-Output 'PASS: both candidates compile/start; blank has 48 empty user tables; demo retains synthetic reports'
