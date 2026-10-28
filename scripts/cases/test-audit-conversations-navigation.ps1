param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$scratch=Join-Path $root ('build/tests/audit-messages-navigation-'+[guid]::NewGuid()+'.accdb')
$before=(Get-FileHash -LiteralPath $BuildPath).Hash
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try {
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 $app.LoadFromText(5,'AuditConversationHarness',(Join-Path $root 'tests/AuditConversationHarness.bas'))
 $app.DoCmd.RunCommand(126)
 foreach($expression in @('TestAuditConversation()','TestNavigation()')){
  $result=$app.Eval($expression);Write-Output $result
  if($result -notlike 'PASS:*'){throw $result}
 }
 $inventory=Join-Path $root ('build/tests/navigation-inventory-'+[guid]::NewGuid())
 [void](New-Item -ItemType Directory $inventory)
 $forms=@($app.CurrentProject.AllForms|ForEach-Object {$_.Name})
 $macros=@($app.CurrentProject.AllMacros|ForEach-Object {$_.Name})
 foreach($name in $forms){$app.SaveAsText(2,$name,(Join-Path $inventory ($name+'.txt')))}
 foreach($name in $macros){$app.SaveAsText(4,$name,(Join-Path $inventory ($name+'.txt')))}
 foreach($name in @('frmHelp','frmChecklist','frmReviewComments','frmUXReview','frmUXNoteComposer')){
  if($forms -contains $name){throw ('Retired duplicate remains: '+$name)}
 }
 @{Forms=$forms;Macros=$macros;Routes=$app.Eval('UXWorkspaceRouteList()');BuildPath=$BuildPath;BuildSHA256=$before;Bindings='Full native form and macro text exported, including events, control expressions and source bindings'}|ConvertTo-Json -Depth 5|Set-Content (Join-Path $inventory 'inventory.json') -Encoding UTF8
 Write-Output ('PASS: live form/macro/event inventory: '+$inventory)
} finally {
 try{[void]$app.Eval('UXLeaveEditor("Discard")')}catch{}
 $app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
 if((Get-FileHash -LiteralPath $BuildPath).Hash -ne $before){throw 'Checks changed the supplied build'}
}
Write-Output ('Fixture: '+$scratch)
