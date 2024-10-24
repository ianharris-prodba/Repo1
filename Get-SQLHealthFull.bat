cd /D "%~dp0"
powershell.exe -ExecutionPolicy Bypass -File .\Get-SQLHealth.ps1 -Config config.xml "HEALTHCHECK" "FULL" "No message"