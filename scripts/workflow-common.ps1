# Shared process boundary: native Access suites run serially in Windows PowerShell.
# A failed child cannot be mistaken for a successful task or receipt.
function InvokeSwiftWorkflowStep([string]$Name,[string[]]$Arguments,[string]$LogFolder){
    $log=Join-Path $LogFolder ('{0:D2}-{1}.log' -f $script:workflowResults.Count,$Name)
    Write-Output ('Running '+$Name+'; log: '+$log)
    $timer=[Diagnostics.Stopwatch]::StartNew()
    try {
        $ErrorActionPreference='Continue'
        & powershell.exe -NoProfile -File (Join-Path $PSScriptRoot ($Name+'.ps1')) @Arguments *> $log
        $stepExit=$LASTEXITCODE
    } finally { $ErrorActionPreference='Stop'; $timer.Stop() }
    if($stepExit -ne 0){
        Get-Content -LiteralPath $log -Tail 35
        throw ($Name+' failed (exit '+$stepExit+'). See '+$log)
    }
    Get-Content -LiteralPath $log | Where-Object {$_ -like 'PASS*'} | Select-Object -Last 3
    $script:workflowResults.Add([pscustomobject]@{Script=$Name;Status='PASS';Seconds=$timer.Elapsed.TotalSeconds;Log=$log})
}

function GetSwiftWorkflowArtifact([string]$Path){
    $item=Get-Item -LiteralPath $Path -ErrorAction Stop
    if($item.PSIsContainer -or $item.Extension -ne '.accdb'){throw 'Expected an existing .accdb file'}
    [pscustomobject]@{Path=$item.FullName;Bytes=$item.Length;SHA256=(Get-FileHash -LiteralPath $item.FullName -Algorithm SHA256).Hash}
}

# Build receipts cover compile/reopen, installed source and form bindings only.
# They never replace a current blank inventory or workflow/migration checks.
function TestSwiftBuildReceipt($Artifact,$Plan){
    try {
        AssertSwiftBuildInputs $Plan
        $receipt=Get-Content -LiteralPath ($Artifact.Path+'.build.json') -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if($receipt.Status -cne 'Built and structurally verified; workflow tests required'){return $false}
        if([IO.Path]::GetFullPath($receipt.Output) -ne $Artifact.Path -or $receipt.SHA256 -ne $Artifact.SHA256){return $false}
        if($receipt.AppVersion -cne $Plan.AppVersion -or $receipt.SchemaVersion -ne $Plan.SchemaVersion){return $false}
        $verification=$receipt.Verification
        foreach($flag in @('Identity','NativeUnloadGuards','RetiredFormsAbsent')){
            if($verification.$flag -isnot [bool] -or -not $verification.$flag){return $false}
        }
        if($verification.SourceModules -ne $Plan.Manifest.sourceModules.Count -or $verification.Forms -le 0){return $false}
        $inputs=@($receipt.Inputs.PSObject.Properties)
        if($inputs.Count -ne $Plan.InputHashes.Count){return $false}
        foreach($key in $Plan.InputHashes.Keys){
            $property=$receipt.Inputs.PSObject.Properties[$key]
            if($null -eq $property -or $property.Value -ne $Plan.InputHashes[$key]){return $false}
        }
        return $true
    } catch {
        # A receipt is an optimization, never permission to bypass verification.
        return $false
    }
}

# Extended-tier runner with evidence capture. While the runner works, VBA error
# dialogs in Access processes it started are recorded and closed (OK for compile
# errors, End for run-time errors) so a dialog fails the step instead of hanging
# the tier. After a failure the visible Access processes are recorded too.
function InvokeSwiftMonitoredStep([string]$Name,[string[]]$Arguments,[string]$LogFolder,[int]$TimeoutMinutes=45){
    if(-not ('SwiftDialogs' -as [type])){
        Add-Type @'
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;
public static class SwiftDialogs {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc p, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr h, EnumProc p, IntPtr l);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  public static string CloseVbaError(uint target){
    string done=null;
    EnumWindows((h,l)=>{uint p;GetWindowThreadProcessId(h,out p);
      if(p!=target || !IsWindowVisible(h)) return true;
      var t=new StringBuilder(512);var c=new StringBuilder(64);GetWindowText(h,t,512);GetClassName(h,c,64);
      if(c.ToString()!="#32770" || !t.ToString().StartsWith("Microsoft Visual Basic")) return true;
      var texts=new List<string>();var buttons=new Dictionary<string,IntPtr>();
      EnumChildWindows(h,(ch,l2)=>{var t2=new StringBuilder(2048);var c2=new StringBuilder(64);GetWindowText(ch,t2,2048);GetClassName(ch,c2,64);
        if(c2.ToString()=="Button") buttons[t2.ToString().Replace("&","")]=ch; else if(t2.Length>0) texts.Add(t2.ToString());return true;},IntPtr.Zero);
      var body=string.Join(" ",texts).Replace("\r"," ").Replace("\n"," ");
      string press=body.Contains("Compile error")?"OK":(body.Contains("Run-time error")?"End":null);
      if(press!=null && buttons.ContainsKey(press)){SendMessage(buttons[press],0x00F5,IntPtr.Zero,IntPtr.Zero);done="pressed "+press+": "+body;}
      return true;},IntPtr.Zero);
    return done;}
}
'@
    }
    $log=Join-Path $LogFolder ('{0:D2}-{1}.log' -f $script:workflowResults.Count,$Name)
    Write-Output ('Running '+$Name+'; log: '+$log)
    $started=Get-Date;$timer=[Diagnostics.Stopwatch]::StartNew();$dialogs=[Collections.Generic.List[string]]::new()
    $quoted=@($Arguments|ForEach-Object {if($_ -match '\s'){'"'+$_+'"'}else{$_}})
    $process=Start-Process powershell.exe -ArgumentList (@('-NoProfile','-File',('"'+(Join-Path $PSScriptRoot ($Name+'.ps1'))+'"'))+$quoted) -RedirectStandardOutput $log -RedirectStandardError ($log+'.err') -WindowStyle Hidden -PassThru
    $null=$process.Handle
    while(-not $process.WaitForExit(2000)){
        foreach($access in @(Get-Process MSACCESS -ErrorAction SilentlyContinue|Where-Object {$_.StartTime -ge $started})){
            $closed=[SwiftDialogs]::CloseVbaError([uint32]$access.Id);if($closed){$dialogs.Add('PID '+$access.Id+' '+$closed)}
        }
        if($timer.Elapsed.TotalMinutes -ge $TimeoutMinutes){$dialogs.Add('Timed out after '+$TimeoutMinutes+' minutes; runner left running (PID '+$process.Id+')');break}
    }
    $exit=if($process.HasExited){$process.ExitCode}else{-1}
    $processes=@()
    if($exit -ne 0){$processes=@(Get-Process MSACCESS -ErrorAction SilentlyContinue|ForEach-Object {'PID '+$_.Id+' started '+$_.StartTime.ToString('s')+' window "'+$_.MainWindowTitle+'"'})}
    $errorText=@(Get-Content -LiteralPath ($log+'.err') -ErrorAction SilentlyContinue|Where-Object {$_ -and $_ -notmatch '^\s*(\+|At |CategoryInfo|FullyQualified)'}|Select-Object -First 3)-join ' '
    [pscustomobject]@{Script=$Name;ExitCode=$exit;Seconds=$timer.Elapsed.TotalSeconds;Log=$log;Error=$errorText;Dialogs=@($dialogs.ToArray());AccessProcesses=$processes}
}
