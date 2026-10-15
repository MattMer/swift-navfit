param([string]$Executable='D:\Development\Swift NAVFIT\NAVFIT98A\NAVFIT98A.exe', [Parameter(Mandatory=$true)][string]$OutputDirectory)
# Bounded runtime proof: invoke only drawing-only methods identified by REA.
# No PrintDocument.Print, dialog, report-data routine or application entry point.
$ErrorActionPreference='Stop'
$Executable=[IO.Path]::GetFullPath($Executable);$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
$expected='5366F1C1BB147F1706094C5E3E1D62981F34FC6CB301E4006A892ACFBFAB58CF'
if((Get-FileHash -LiteralPath $Executable).Hash -ne $expected){throw 'Binary differs from inspected NAVFIT98A 2.2.0.332; inspect before execution.'}
if(Test-Path -LiteralPath $OutputDirectory){throw 'Refusing existing output directory'}
[void](New-Item -ItemType Directory -Path $OutputDirectory)
Add-Type -AssemblyName System.Drawing
$assembly=[Reflection.Assembly]::LoadFrom($Executable)
$type=$assembly.GetType('NAVFIT98A.PrintModule',$true)
$flags=[Reflection.BindingFlags]'Public,Static'
$results=@()
foreach($proof in @(@('EVAL-page2','NF_EvalPage2Frame','NF_EvalPage2Text'),@('FITREP-page2','NF_FitRepPage2Frame','NF_FitRepPage2Text'),@('CHIEFEVAL-page2-frame','NF_ChiefPage2Frame'))){
 $bitmap=$null;$graphics=$null
 try{
  $bitmap=New-Object Drawing.Bitmap 850,1100;$bitmap.SetResolution(100,100)
  $graphics=[Drawing.Graphics]::FromImage($bitmap);$graphics.Clear([Drawing.Color]::White)
  $graphics.PageUnit=[Drawing.GraphicsUnit]::Pixel
  $bounds=New-Object Drawing.Rectangle 0,0,850,1100
  $settings=New-Object Drawing.Printing.PageSettings
  $pageEvent=New-Object Drawing.Printing.PrintPageEventArgs $graphics,$bounds,$bounds,$settings
  for($i=1;$i -lt $proof.Count;$i++){
   $method=$type.GetMethod($proof[$i],$flags)
   if(-not $method){throw ('Missing drawing method '+$proof[$i])}
   [void]$method.Invoke($null,[object[]]@($null,$pageEvent.PSObject.BaseObject))
  }
  $path=Join-Path $OutputDirectory ($proof[0]+'.png');$bitmap.Save($path,[Drawing.Imaging.ImageFormat]::Png)
  $ink=0
  for($y=0;$y -lt 1100;$y+=4){for($x=0;$x -lt 850;$x+=4){if($bitmap.GetPixel($x,$y).R -lt 240){$ink++}}}
  if($ink -lt 100){throw ('Rendering unexpectedly empty: '+$proof[0])}
  $results+=@{Image=$path;Methods=@($proof|Select-Object -Skip 1);SampledInkPixels=$ink;SHA256=(Get-FileHash -LiteralPath $path).Hash}
 }finally{if($graphics){$graphics.Dispose()};if($bitmap){$bitmap.Dispose()}}
}
$fonts=@($type.GetFields($flags)|Where-Object {$_.FieldType.FullName -eq 'System.Drawing.Font'}|ForEach-Object {$font=$_.GetValue($null);@{Field=$_.Name;Family=$font.Name;Points=$font.SizeInPoints;Style=$font.Style.ToString()}})
$com=@($assembly.GetCustomAttributesData()|Where-Object {$_.AttributeType.FullName -eq 'System.Runtime.InteropServices.ComVisibleAttribute'}|ForEach-Object {$_.ConstructorArguments[0].Value})
if((Get-FileHash -LiteralPath $Executable).Hash -ne $expected){throw 'Input binary changed during proof'}
@{Fonts=$fonts;AssemblyComVisible=$com;ExecutableSHA256=$expected;AssemblyVersion=$assembly.GetName().Version.ToString();Runtime=[Environment]::Version.ToString();Kind='Bounded non-REA runtime proof';PrintJobsSubmitted=0;DatabaseMethodsCalled=$false;Images=$results;Limits='Static frame/text only; CHIEF text and all report data excluded; bitmap is not a printer fidelity proof.'}|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $OutputDirectory 'proof.json') -Encoding utf8
Write-Output ('PASS: '+$results.Count+' offscreen drawing proofs; no print jobs; '+$OutputDirectory)
