param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/sort-bindings-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$coverage=Get-Content (Join-Path $root 'build/r115-sort-coverage.json') -Raw|ConvertFrom-Json
$app=New-Object -ComObject Access.Application;$count=0
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 for($wait=0;$wait -lt 30 -and $app.Forms.Count -eq 0;$wait++){Start-Sleep -Milliseconds 100}
 foreach($name in @($app.CurrentProject.AllForms|Where-Object {$_.IsLoaded}|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$name,2)}
 foreach($group in $coverage|Group-Object Form){
  $app.DoCmd.OpenForm($group.Name,1);$f=$app.Forms.Item($group.Name)
  foreach($entry in $group.Group){
   $c=$f.Controls($entry.Control)
   switch($entry.Method){
    'Native header click' {
     if($c.OnMouseDown -ne '[Event Procedure]' -or -not $f.HasModule){throw ('Missing class binding: '+$f.Name)}
     $text=$f.Module.Lines(1,$f.Module.CountOfLines)
     if($text -notmatch ($c.Name+'_MouseDown') -or $text -notmatch ('UXSortNativeHeader\(Me.Name, "'+$c.Name+'"')){throw ('Missing persisted event code: '+$f.Name+'.'+$c.Name)}
    }
    'Custom clickable headings' {if($f.Controls('sort_'+$c.Name+'_1').OnClick -notlike '=UXSortList*'){throw 'Missing custom heading'}}
    'Editable grid header' {if($c.OnClick -ne ('=UXSortGrid("'+$entry.Field+'")')){throw 'Missing grid heading'}}
    'RSCA table header' {if($c.OnClick -notlike '=RSTableSort*'){throw 'Missing RSCA heading'}}
   }
   $count++
  }
  $app.DoCmd.Close(2,$f.Name,2)
 }
 $app.DoCmd.RunCommand(126)
 Write-Output ('PASS: '+$count+' persisted table/header bindings reopened and verified')
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
