Get-ChildItem -Filter "*.flv" | ForEach-Object {
    $name = $_.Name
    $outputName = ($name).Trim(".flv") + ".mp4"

    ffmpeg -i $name -codec copy $outputName
}

Write-Output "Done ;)"