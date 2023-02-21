Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Force
Install-Module PSWindowsUpdate 
Import-Module PSWindowsUpdate
Install-windowsupdate -ForceInstall -AcceptAll
Set-ExecutionPolicy -ExecutionPolicy AllSigned -Force
