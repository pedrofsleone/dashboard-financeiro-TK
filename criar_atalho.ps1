$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$env:USERPROFILE\Desktop\Dashboard TK.lnk")
$Shortcut.TargetPath = "$env:USERPROFILE\Desktop\dashboard-financeiro\iniciar.bat"
$Shortcut.WorkingDirectory = "$env:USERPROFILE\Desktop\dashboard-financeiro"
$Shortcut.Description = "Abrir Dashboard Financeiro Chinezinho"
$Shortcut.Save()
Write-Host "Atalho criado com sucesso!"
