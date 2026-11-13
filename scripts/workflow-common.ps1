# Shared process boundary: native Access suites run serially in Windows PowerShell.
# A failed child cannot be mistaken for a successful task or receipt.
# Progress lines are shown live with the time, so a long run shows where it is.
# Nested runs pass them up; a line that already has a time keeps it.
$script:swiftProgressPattern='^(\[\d\d:\d\d:\d\d\] )?(Running |Done |Still running |Retrying |PASS|FAIL|FLAKY|SKIP)'
function WriteSwiftProgress([string]$Line){
    if($Line -match '^\[\d\d:\d\d:\d\d\] '){Write-Host $Line}else{Write-Host ('['+(Get-Date -Format 'HH:mm:ss')+'] '+$Line)}
}

function FormatSwiftElapsed([TimeSpan]$Elapsed){'{0}:{1:D2}' -f [int][Math]::Floor($Elapsed.TotalMinutes),$Elapsed.Seconds}

function InvokeSwiftWorkflowStep([string]$Name,[string[]]$Arguments,[string]$LogFolder){
    $log=Join-Path $LogFolder ('{0:D2}-{1}.log' -f $script:workflowResults.Count,($Name -replace '[\\/]','-'))
    WriteSwiftProgress ('Running '+$Name+'; log: '+$log)
    $timer=[Diagnostics.Stopwatch]::StartNew()
    try {
        $ErrorActionPreference='Continue'
        & powershell.exe -NoProfile -File (Join-Path $PSScriptRoot ($Name+'.ps1')) @Arguments *>&1 | ForEach-Object {
            $_ | Out-File -LiteralPath $log -Append
            $line=if($_ -is [Management.Automation.ErrorRecord]){$_.Exception.Message}else{[string]$_}
            if($line -match $script:swiftProgressPattern){WriteSwiftProgress $line}
        }
        $stepExit=$LASTEXITCODE
    } finally { $ErrorActionPreference='Stop'; $timer.Stop() }
    if($stepExit -ne 0){
        Get-Content -LiteralPath $log -Tail 35
        throw ($Name+' failed (exit '+$stepExit+'). See '+$log)
    }
    WriteSwiftProgress ('Done '+$Name+' in '+(FormatSwiftElapsed $timer.Elapsed))
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
    #A retry keeps the first run's log as evidence and writes its own.
    if(Test-Path -LiteralPath $log){$log=Join-Path $LogFolder ('{0:D2}-{1}-retry.log' -f $script:workflowResults.Count,$Name)}
    WriteSwiftProgress ('Running '+$Name+'; log: '+$log)
    $started=Get-Date;$timer=[Diagnostics.Stopwatch]::StartNew();$dialogs=[Collections.Generic.List[string]]::new()
    $shown=0;$nextHeartbeat=5
    $quoted=@($Arguments|ForEach-Object {if($_ -match '\s'){'"'+$_+'"'}else{$_}})
    $process=Start-Process powershell.exe -ArgumentList (@('-NoProfile','-File',('"'+(Join-Path $PSScriptRoot ($Name+'.ps1'))+'"'))+$quoted) -RedirectStandardOutput $log -RedirectStandardError ($log+'.err') -WindowStyle Hidden -PassThru
    $null=$process.Handle
    #Echo the runner's new progress lines; a quiet runner gets a heartbeat every 5 minutes.
    $showProgress={
        $lines=@(Get-Content -LiteralPath $log -ErrorAction SilentlyContinue)
        for(;$shown -lt $lines.Count;$shown++){if($lines[$shown] -match $script:swiftProgressPattern){WriteSwiftProgress $lines[$shown]}}
    }
    while(-not $process.WaitForExit(2000)){
        . $showProgress
        if($timer.Elapsed.TotalMinutes -ge $nextHeartbeat){WriteSwiftProgress ('Still running '+$Name+' ('+(FormatSwiftElapsed $timer.Elapsed)+')');$nextHeartbeat+=5}
        foreach($access in @(Get-Process MSACCESS -ErrorAction SilentlyContinue|Where-Object {$_.StartTime -ge $started})){
            $closed=[SwiftDialogs]::CloseVbaError([uint32]$access.Id);if($closed){$dialogs.Add('PID '+$access.Id+' '+$closed)}
        }
        if($timer.Elapsed.TotalMinutes -ge $TimeoutMinutes){$dialogs.Add('Timed out after '+$TimeoutMinutes+' minutes; runner left running (PID '+$process.Id+')');break}
    }
    . $showProgress
    $exit=if($process.HasExited){$process.ExitCode}else{-1}
    $processes=@()
    if($exit -ne 0){$processes=@(Get-Process MSACCESS -ErrorAction SilentlyContinue|ForEach-Object {'PID '+$_.Id+' started '+$_.StartTime.ToString('s')+' window "'+$_.MainWindowTitle+'"'})}
    $errorText=@(Get-Content -LiteralPath ($log+'.err') -ErrorAction SilentlyContinue|Where-Object {$_ -and $_ -notmatch '^\s*(\+|At |CategoryInfo|FullyQualified)'}|Select-Object -First 3)-join ' '
    [pscustomobject]@{Script=$Name;ExitCode=$exit;Seconds=$timer.Elapsed.TotalSeconds;Log=$log;Error=$errorText;Dialogs=@($dialogs.ToArray());AccessProcesses=$processes}
}

# Desktop isolation. Native runs happen on a separate Windows desktop in the same
# sign-in session, so Access windows, pop-up forms and dialogs never appear on the
# user's screen and cannot take the user's keystrokes, clicks or focus. Access that
# a runner starts through COM lands on the runner's desktop, and child processes
# inherit it. SWIFT_TEST_DESKTOP names the desktop inside an isolated run; set it to
# 'visible' beforehand to run on the normal desktop and watch a failure.
# A hidden VBA error dialog would wait for ever, so the watcher presses OK (compile
# error) or End (run-time error) once it has been open six seconds, which leaves the
# extended tier's own monitor time to record it first. Any other dialog still open
# after a minute is reported once and left alone.
function AddSwiftDesktopType{
    if('SwiftDesktop' -as [type]){return}
    Add-Type @'
using System; using System.Text; using System.Threading; using System.Collections.Generic; using System.Runtime.InteropServices;
public sealed class SwiftDesktop : IDisposable {
  [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)] struct STARTUPINFO { public int cb; public string lpReserved; public string lpDesktop; public string lpTitle; public int dwX, dwY, dwXSize, dwYSize, dwXCountChars, dwYCountChars, dwFillAttribute, dwFlags; public short wShowWindow, cbReserved2; public IntPtr lpReserved2, hStdInput, hStdOutput, hStdError; }
  [StructLayout(LayoutKind.Sequential)] struct PROCESS_INFORMATION { public IntPtr hProcess, hThread; public int dwProcessId, dwThreadId; }
  [StructLayout(LayoutKind.Sequential)] struct SECURITY_ATTRIBUTES { public int nLength; public IntPtr lpSecurityDescriptor; public bool bInheritHandle; }
  delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)] static extern IntPtr CreateDesktop(string name, IntPtr device, IntPtr devmode, int flags, uint access, IntPtr sa);
  [DllImport("user32.dll")] static extern bool CloseDesktop(IntPtr h);
  [DllImport("user32.dll")] static extern bool SetThreadDesktop(IntPtr h);
  [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)] static extern bool CreateProcess(string app, StringBuilder cmd, IntPtr pa, IntPtr ta, bool inherit, uint flags, IntPtr env, string dir, ref STARTUPINFO si, out PROCESS_INFORMATION pi);
  [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)] static extern IntPtr CreateFile(string name, uint access, uint share, ref SECURITY_ATTRIBUTES sa, uint disposition, uint flags, IntPtr template);
  [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr h);
  [DllImport("kernel32.dll")] static extern uint WaitForSingleObject(IntPtr h, uint ms);
  [DllImport("kernel32.dll")] static extern bool GetExitCodeProcess(IntPtr h, out uint code);
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc p, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr h, EnumProc p, IntPtr l);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  IntPtr desk, process; volatile bool stop; Thread watcher;
  readonly Queue<string> notes=new Queue<string>();
  readonly Dictionary<IntPtr,DateTime> seen=new Dictionary<IntPtr,DateTime>();
  readonly HashSet<IntPtr> handled=new HashSet<IntPtr>();
  public int ProcessId;
  public SwiftDesktop(string name, string commandLine, string logPath){
    desk=CreateDesktop(name,IntPtr.Zero,IntPtr.Zero,0,0x10000000,IntPtr.Zero);
    if(desk==IntPtr.Zero) throw new System.ComponentModel.Win32Exception();
    var sa=new SECURITY_ATTRIBUTES();sa.nLength=Marshal.SizeOf(typeof(SECURITY_ATTRIBUTES));sa.bInheritHandle=true;
    IntPtr log=CreateFile(logPath,0x40000000,3,ref sa,2,0x80,IntPtr.Zero);
    if(log==new IntPtr(-1)){int e=Marshal.GetLastWin32Error();CloseDesktop(desk);throw new System.ComponentModel.Win32Exception(e);}
    var si=new STARTUPINFO();si.cb=Marshal.SizeOf(typeof(STARTUPINFO));si.lpDesktop=name;si.dwFlags=0x100;si.hStdOutput=log;si.hStdError=log;
    PROCESS_INFORMATION pi;
    bool ok=CreateProcess(null,new StringBuilder(commandLine),IntPtr.Zero,IntPtr.Zero,true,0x08000000,IntPtr.Zero,null,ref si,out pi);
    int error=Marshal.GetLastWin32Error();CloseHandle(log);
    if(!ok){CloseDesktop(desk);throw new System.ComponentModel.Win32Exception(error);}
    CloseHandle(pi.hThread);process=pi.hProcess;ProcessId=pi.dwProcessId;
    watcher=new Thread(Watch);watcher.IsBackground=true;watcher.Start();
  }
  public bool Wait(int ms){return WaitForSingleObject(process,(uint)ms)==0;}
  public int ExitCode(){uint c;GetExitCodeProcess(process,out c);return unchecked((int)c);}
  public string[] TakeNotes(){lock(notes){var a=notes.ToArray();notes.Clear();return a;}}
  void Note(string text){lock(notes){notes.Enqueue(text);}}
  void Watch(){
    if(!SetThreadDesktop(desk)){Note("Dialog watcher could not attach to the test desktop; a VBA error dialog there will wait until the step times out");return;}
    while(!stop){try{Scan();}catch(Exception e){Note("Dialog watcher error: "+e.Message);}Thread.Sleep(2000);}
  }
  void Scan(){
    var open=new HashSet<IntPtr>();
    EnumWindows((h,l)=>{
      if(!IsWindowVisible(h)) return true;
      var c=new StringBuilder(64);GetClassName(h,c,64);if(c.ToString()!="#32770") return true;
      open.Add(h);
      DateTime first;if(!seen.TryGetValue(h,out first)){seen[h]=DateTime.UtcNow;return true;}
      if(handled.Contains(h)) return true;
      double age=(DateTime.UtcNow-first).TotalSeconds;
      var t=new StringBuilder(512);GetWindowText(h,t,512);string title=t.ToString();
      bool vba=title.StartsWith("Microsoft Visual Basic");
      if(age<(vba?6:60)) return true;
      var texts=new List<string>();var buttons=new Dictionary<string,IntPtr>();
      EnumChildWindows(h,(ch,l2)=>{var t2=new StringBuilder(2048);var c2=new StringBuilder(64);GetWindowText(ch,t2,2048);GetClassName(ch,c2,64);
        if(c2.ToString()=="Button") buttons[t2.ToString().Replace("&","")]=ch; else if(t2.Length>0) texts.Add(t2.ToString());return true;},IntPtr.Zero);
      string body=string.Join(" ",texts).Replace("\r"," ").Replace("\n"," ");
      uint pid;GetWindowThreadProcessId(h,out pid);
      handled.Add(h);
      string press=!vba?null:(body.Contains("Compile error")?"OK":(body.Contains("Run-time error")?"End":null));
      if(press!=null && buttons.ContainsKey(press)){SendMessage(buttons[press],0x00F5,IntPtr.Zero,IntPtr.Zero);Note("Hidden VBA dialog in PID "+pid+", pressed "+press+": "+body+(press=="OK"?" (if this step now hangs, that Access is waiting in the VBA debugger; end it by PID)":""));}
      else Note("Dialog waiting on the test desktop in PID "+pid+" ("+title+"): "+body);
      return true;},IntPtr.Zero);
    foreach(var h in new List<IntPtr>(seen.Keys)){if(!open.Contains(h)){seen.Remove(h);handled.Remove(h);}}
  }
  public void Dispose(){
    stop=true;if(watcher!=null) watcher.Join(5000);
    if(process!=IntPtr.Zero){CloseHandle(process);process=IntPtr.Zero;}
    if(desk!=IntPtr.Zero){CloseDesktop(desk);desk=IntPtr.Zero;}
  }
}
'@
}

# Run a script with its arguments on a new hidden desktop, echo its output and
# return its exit code. The output is kept in a log file when the run fails.
function InvokeSwiftIsolated([string]$ScriptPath,[string[]]$Arguments){
    AddSwiftDesktopType
    $name='SwiftNavfitTests-'+$PID+'-'+[guid]::NewGuid().ToString('N').Substring(0,6)
    $log=Join-Path ([IO.Path]::GetTempPath()) ($name+'.log')
    $quoted=@($Arguments|ForEach-Object {if($_ -match '[\s"]' -or $_ -eq ''){'"'+($_ -replace '(\\*)"','$1$1\"' -replace '(\\+)$','$1$1')+'"'}else{$_}})
    $command='powershell.exe -NoProfile -File "'+$ScriptPath+'" '+($quoted -join ' ')
    $env:SWIFT_TEST_DESKTOP=$name
    try{$run=New-Object SwiftDesktop($name,$command,$log)}finally{Remove-Item Env:SWIFT_TEST_DESKTOP}
    WriteSwiftProgress ('Running on hidden desktop '+$name+' (runner PID '+$run.ProcessId+'). Set SWIFT_TEST_DESKTOP=visible to watch instead.')
    $reader=$null
    try{
        $encoding=[Text.Encoding]::GetEncoding([Globalization.CultureInfo]::CurrentCulture.TextInfo.OEMCodePage)
        $reader=New-Object IO.StreamReader((New-Object IO.FileStream($log,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)),$encoding)
        $pending=''
        do{
            $finished=$run.Wait(1000)
            $pending+=$reader.ReadToEnd()
            $parts=$pending -split "`r?`n"
            for($i=0;$i -lt $parts.Count-1;$i++){Write-Host $parts[$i]}
            $pending=$parts[$parts.Count-1]
            foreach($note in $run.TakeNotes()){WriteSwiftProgress $note}
        }until($finished)
        $pending+=$reader.ReadToEnd()
        if($pending){Write-Host $pending}
        $exit=$run.ExitCode()
    }finally{
        if($reader){$reader.Dispose()}
        $run.Dispose()
    }
    if($exit -eq 0){Remove-Item -LiteralPath $log -ErrorAction SilentlyContinue}else{WriteSwiftProgress ('Isolated run failed (exit '+$exit+'); output kept in '+$log)}
    $exit
}

# Called at the top of a workflow script with its own path and bound parameters.
# Returns $null when the script should carry on in place (already isolated, or
# SWIFT_TEST_DESKTOP=visible); otherwise reruns it isolated and returns the exit code.
function EnterSwiftIsolation([string]$ScriptPath,$Bound){
    if($env:SWIFT_TEST_DESKTOP){return $null}
    $arguments=@()
    foreach($key in $Bound.Keys){
        $value=$Bound[$key]
        if($value -is [switch]){if($value.IsPresent){$arguments+=('-'+$key)}}else{$arguments+=@(('-'+$key),[string]$value)}
    }
    InvokeSwiftIsolated $ScriptPath $arguments
}
