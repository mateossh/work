# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/start-process?view=powershell-7.1#parameters
# https://adamtheautomator.com/powershell-escape-double-quotes/
# https://stackoverflow.com/questions/4146635/custom-sorting-in-powershell
# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/sort-object?view=powershell-7.1
# (example 8)

$path = Read-Host "Enter path"
$name_part1 = { if ($_.Name -match '^(\d+)') { [int]$matches[1] } }

Get-ChildItem "$path" -Filter "*.pdf" | Sort-Object $name_part1 | Foreach-Object {
    $name = $_.FullName

    Start-Process 'chrome' -ArgumentList "--incognito","`"$name`""
    Start-Sleep -m 50
}