# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/start-process?view=powershell-7.1#parameters
# https://adamtheautomator.com/powershell-escape-double-quotes/

$path = Read-Host "Enter path"

Get-ChildItem "$path" -Filter "*.pdf" | Foreach-Object {
    $name = $_.FullName

    Start-Process 'chrome' -ArgumentList "--incognito","`"$name`""
    Start-Sleep -m 50
}