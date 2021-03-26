if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Run this script as Administrator!"
    break
}

Write-Host "Administrator rights!"
Stop-Service spooler
Remove-Item -Path $env:windir\system32\spool\PRINTERS\*.*
Start-Service Spooler