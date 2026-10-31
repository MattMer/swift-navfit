param([Parameter(Mandatory=$true)][string]$BuildPath)
# Narratives suite: narrative persistence and the one-snapshot-per-revision check after
# reopening; narrative layout; the line-fit meter (Swift PDF and NAVFIT98A against a System.Drawing reference); NavyWriter snippets with pending and saved values.
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'suite-common.ps1')
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash

Invoke-SuiteCase 'narratives' {
    $scratch=Join-Path $root ('build\tests\narrative-'+[guid]::NewGuid()+'.accdb')
    Copy-Item -LiteralPath $BuildPath -Destination $scratch
    $app=$null
    try {
        $app=New-Object -ComObject Access.Application
        $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
        $app.LoadFromText(5,'NarrativeHarness',(Join-Path $root 'tests\NarrativeHarness.bas'))
        $app.DoCmd.RunCommand(126)
        $result=$app.Eval('TestNarratives()')
        if(-not $result.StartsWith('PASS:')){throw $result}
        $app.CloseCurrentDatabase();$app.OpenCurrentDatabase($scratch,$true)
        $db=$app.CurrentDb()
        $r=$db.OpenRecordset("SELECT RevisionID FROM ReportRevision AS R WHERE (SELECT Count(*) FROM RevisionValue AS V WHERE V.RevisionID=R.RevisionID AND V.FieldKey='OfficialComments')<>1",4)
        if(-not $r.EOF){throw 'A saved revision is missing its unique narrative snapshot.'}
        $r.Close();$db.Close();$app.CloseCurrentDatabase()
        Write-Output $result
    }finally{
        if($null -ne $app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
    }
}
Invoke-HarnessCase 'narrative-layout' $BuildPath 'NarrativeLayoutHarness' 'TestNarrativeLayout()'
Invoke-HarnessCase 'narrative-line-fit' $BuildPath 'NarrativeLayoutHarness' 'TestNarrativeLineFit()'
Invoke-SuiteCase 'navfit-line-fit' {
    # Reference measurer: NAVFIT98A 2.2.0.332 prints each narrative with
    # Graphics.DrawString(text, font, brush, RectangleF) and no StringFormat on the printer
    # Graphics (page unit Display). Expected line counts come from System.Drawing here;
    # the VBA profile (NAVFITFieldFit) must agree. Fonts, widths and limits are listed in
    # planning/NAVFIT_PRINT_ENGINE_FINDINGS.md.
    Add-Type -AssemblyName System.Drawing
    $folder=Join-Path $root ('build\tests\navfit-fit-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $folder)
    $graphics=(New-Object System.Drawing.Printing.PrinterSettings).CreateMeasurementGraphics()
    $graphics.PageUnit=[System.Drawing.GraphicsUnit]::Display
    function NavfitLines([string]$text,[single]$size,[single]$width){
        if($text -eq ''){return 0}
        $font=New-Object System.Drawing.Font('Courier New',$size)
        $fitted=0;$lines=0
        [void]$graphics.MeasureString($text,$font,(New-Object System.Drawing.SizeF($width,1000000)),[System.Drawing.StringFormat]$null,[ref]$fitted,[ref]$lines)
        $font.Dispose();return $lines
    }
    $names='alpha bravo charlie delta echo foxtrot golf hotel india juliet kilo lima'.Split(' ')
    function Words([int]$n){(@(for($i=0;$i -lt $n;$i++){$names[$i%12]}) -join ' ')}
    function Numbered([int]$n){(@(for($i=1;$i -le $n;$i++){'Line '+$i}) -join "`r`n")}
    # kind, key, Courier New size, width, NAVFIT98A line limit, printed prefix, edit MaxLength
    $profiles=@(
        @('EVAL','OfficialComments',9,715,18,'',2500),@('FITREP','OfficialComments',9,715,18,'',2000),@('CHIEFEVAL','OfficialComments',10,760,6,'',700),
        @('EVAL','CommandEmployment',9.65,755,3,'',600),@('FITREP','CommandEmployment',9.65,755,3,'',600),@('CHIEFEVAL','CommandEmployment',9.65,755,3,'',600),
        @('EVAL','DutyDescription',9.65,755,4,(' '*21),500),@('FITREP','DutyDescription',9.65,755,4,(' '*21),600),@('CHIEFEVAL','DutyDescription',9.65,755,4,(' '*21),600),
        @('CHIEFEVAL','Trait33Comment1',9.65,450,2,'',400),@('CHIEFEVAL','Trait39Comment3',9.65,450,2,'',400))
    $rows=New-Object System.Collections.Generic.List[string];$n=0
    foreach($p in $profiles){
        $kind,$key,$size,$width,$limit,$prefix,$max=$p
        $cases=New-Object System.Collections.Generic.List[object]
        $m=1;while((NavfitLines ($prefix+(Words ($m+1))) $size $width) -eq 1){$m++}
        $cases.Add(@('one-word last line',(Words ($m+1)),1))
        $m=1;while((NavfitLines ($prefix+(Words ($m+1))) $size $width) -le $limit){$m++}
        if((NavfitLines ($prefix+(Words $m)) $size $width) -ne $limit){throw ($kind+' '+$key+': no word count fills exactly '+$limit+' lines')}
        $cases.Add(@('exactly full',(Words $m),-1));$cases.Add(@('one line over',(Words ($m+1)),-1))
        $cases.Add(@('numbered full',(Numbered $limit),0));$cases.Add(@('numbered over',(Numbered ($limit+1)),-1))
        $cases.Add(@('blank lines',"First`r`n`r`nThird line",-1))
        $cases.Add(@('trailing blank lines (CRLF)',"First line`r`nEnd`r`n`r`n",0));$cases.Add(@('trailing blank line (LF)',"First line`nEnd`n",0))
        $cases.Add(@('maximum length',(Words 1000).Substring(0,$max),-1))
        foreach($c in $cases){
            $n++;$file=Join-Path $folder ('case'+$n+'.txt')
            [IO.File]::WriteAllText($file,$c[1],[Text.Encoding]::Unicode)
            $rows.Add((@($kind,$key,(NavfitLines ($prefix+$c[1]) $size $width),$limit,$c[2],$file,$c[0]) -join '|'))
        }
    }
    $graphics.Dispose()
    $manifest=Join-Path $folder 'cases.txt'
    [IO.File]::WriteAllLines($manifest,$rows,(New-Object Text.UTF8Encoding($false)))
    $scratch=Join-Path $folder 'fixture.accdb'
    Copy-Item -LiteralPath $BuildPath -Destination $scratch
    $app=$null
    try {
        $app=New-Object -ComObject Access.Application
        $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
        $app.LoadFromText(5,'NarrativeLayoutHarness',(Join-Path $root 'tests\NarrativeLayoutHarness.bas'))
        $app.DoCmd.RunCommand(126)
        $result=$app.Eval('TestNavfitReference("'+$manifest.Replace('"','""')+'")')
        if(-not $result.StartsWith('PASS:')){throw $result}
        Write-Output ($result+'; '+$rows.Count+' cases in '+$folder)
    }finally{
        if($null -ne $app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
    }
}
Invoke-HarnessCase 'navywriter' $BuildPath 'UXNavyWriterHarness' 'TestUXNavyWriter()' -After {param($app) [void]$app.Eval('UXLeaveEditor("Discard")')}

Complete-SuiteRun 'narratives' $BuildPath $buildHash
