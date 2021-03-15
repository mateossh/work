# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/start-process?view=powershell-7.1#parameters
# https://adamtheautomator.com/powershell-escape-double-quotes/

Get-ChildItem -Path "./*.pdf" | Foreach-Object {
    $name = $_.FullName

    Start-Process 'chrome' -ArgumentList "--incognito","`"$name`""
}