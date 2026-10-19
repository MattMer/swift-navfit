# Access COM setters can cache incompatible numeric/variant argument types in PowerShell.
# Match the existing native property type and bypass the dynamic setter call site.
function SetAccessProperty($target,[string]$name,$value){
 $property=$target.Properties.Item($name)
 $current=$property.Value
 if($null -ne $current -and $current -is [ValueType] -and $value -is [IConvertible]){
  $value=[Convert]::ChangeType($value,$current.GetType(),[Globalization.CultureInfo]::InvariantCulture)
 }
 [void]$property.GetType().InvokeMember('Value',[Reflection.BindingFlags]::SetProperty,$null,$property,[object[]]@($value))
}
