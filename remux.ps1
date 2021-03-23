$path = Read-Host "Enter path"

Get-ChildItem "$path" -Filter "*.flv" | ForEach-Object {
    $name = "$($path)\$($_.Name)"
    $outputName = ($name).Trim(".flv") + ".mp4"

    ffmpeg -i $name -codec copy $outputName
}

Write-Output "Done ;)"