# Native builders may leave an empty event stub when class changes are saved/reopened.
# Remove only redundant empty stubs; preserve any nonempty procedure and its code.
function EnsureNativeUnload([string]$formName,[string]$guard){
 $app.DoCmd.OpenForm($formName,1);$form=$app.Forms.Item($formName)
 $form.HasModule=$true;$module=$form.Module
 $text=$module.Lines(1,$module.CountOfLines)
 $procedures=[regex]::Matches($text,'(?im)^Private Sub Form_Unload\(Cancel As Integer\)\s*\r?\n(?<body>[\s\S]*?)^End Sub[ \t]*')
 if($procedures.Count -gt 1){
  $nonempty=@($procedures|Where-Object {$_.Groups['body'].Value.Trim().Length -gt 0})
  if($nonempty.Count -gt 1){throw ('Multiple nonempty unload procedures: '+$formName)}
  $keep=if($nonempty.Count){$nonempty[0].Index}else{$procedures[0].Index}
  for($i=$procedures.Count-1;$i -ge 0;$i--){
   $procedure=$procedures[$i]
   if($procedure.Index -eq $keep){continue}
   $line=1+([regex]::Matches($text.Substring(0,$procedure.Index),'\n')).Count
   $count=1+([regex]::Matches($procedure.Value,'\n')).Count
   $module.DeleteLines($line,$count)
  }
 }
 try{$line=$module.ProcStartLine('Form_Unload',0)}catch{$line=$module.CreateEventProc('Unload','Form')}
 if(-not $module.Lines(1,$module.CountOfLines).Contains($guard)){$module.InsertLines($line+1,$guard)}
 $form.OnUnload='[Event Procedure]'
 $app.DoCmd.Close(2,$formName,1)
}
EnsureNativeUnload 'frmEditor' '    If Not UXEditorUnload() Then Cancel = True'
EnsureNativeUnload 'frmUXComparison' '    If Not UXGridUnload(Me) Then Cancel = True'
Write-Output 'Native unload procedures verified without duplicate empty stubs.'
