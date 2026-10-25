# Database startup properties: the navigation pane stays hidden.
try{$app.CurrentDb().Properties('StartupShowDBWindow').Value=$false}catch{$db=$app.CurrentDb();$p=$db.CreateProperty('StartupShowDBWindow',1,$false);$db.Properties.Append($p)}
