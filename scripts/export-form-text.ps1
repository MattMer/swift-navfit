param([string]$BuildPath,[string]$OutDir)
# Trial (forms-as-source): export every form and report of a build as
# Access text, from a disposable copy so the build itself is never opened.
$ErrorActionPreference='Stop'
[void](New-Item -ItemType Directory -Force -Path $OutDir)
$copy=Join-Path $OutDir 'export-source.accdb';Copy-Item -LiteralPath $BuildPath -Destination $copy -Force
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($copy,$true)
 foreach($form in @($app.Forms|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$form,2)}
 $forms=@($app.CurrentProject.AllForms|ForEach-Object {$_.Name})
 $reports=@($app.CurrentProject.AllReports|ForEach-Object {$_.Name})
 $sw=[Diagnostics.Stopwatch]::StartNew()
 foreach($name in $forms){$app.SaveAsText(2,$name,(Join-Path $OutDir ($name+'.form.txt')))}
 foreach($name in $reports){$app.SaveAsText(3,$name,(Join-Path $OutDir ($name+'.report.txt')))}
 Write-Output ('Exported '+$forms.Count+' forms and '+$reports.Count+' reports in '+[int]$sw.Elapsed.TotalSeconds+'s')
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);Remove-Item -LiteralPath $copy -ErrorAction SilentlyContinue}
