$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("C:\Users\DELL\Desktop\Dong bo NCC MISA-Odoo.lnk")
$Shortcut.TargetPath = "D:\odoo\dong_bo_nha_cung_cap.bat"
$Shortcut.WorkingDirectory = "D:\odoo"
$Shortcut.IconLocation = "C:\Windows\System32\shell32.dll,13"
$Shortcut.Save()
Write-Host "Shortcut NCC created!"
