Set-ExecutionPolicy RemoteSigned
Uninstall-Module -Name "PSWriteColor" -AllVersions
Install-Module -Name "PSWriteColor"
Update-Module -Name "PSWriteColor"
Uninstall-Module -Name SharePointPnPPowerShellOnline -AllVersions -Force
Install-Module -Name SharePointPnPPowerShellOnline -RequiredVersion 3.23.2007.1
Get-Module SharePointPnPPowerShell* -ListAvailable | Select-Object Name,Version | Sort-Object Version -Descending