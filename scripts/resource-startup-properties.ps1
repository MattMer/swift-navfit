# Database startup properties: a plain application window for everyday users.
# Holding Shift while opening still bypasses them (AllowBypassKey stays True), so
# developers and field investigation keep full Access. The ribbon is hidden by
# UXHideAccessObjects at startup. Shortcut menus stay on for Cut/Copy/Paste.
$db=$app.CurrentDb()
function SetStartupProperty([string]$Name,[int]$Type,$Value){
    try{$db.Properties($Name).Value=$Value}catch{$db.Properties.Append($db.CreateProperty($Name,$Type,$Value))}
}
foreach($name in 'StartupShowDBWindow','AllowFullMenus','AllowSpecialKeys','AllowBuiltInToolbars','AllowBreakIntoCode'){SetStartupProperty $name 1 $false}
SetStartupProperty 'AllowBypassKey' 1 $true
SetStartupProperty 'AppTitle' 10 'Swift NAVFIT'
