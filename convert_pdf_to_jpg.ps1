# convert_pdf_to_jpg
#
# add NAPS2 installation folder to PATH
#
# save this file with "UTF-8 with BOM"
#
# running "naps2.console -i file.pdf -o file.jpg" runs scanning process
# with default profile :thinking:

$cwd = (Get-Item .).FullName
$source_path = $cwd + "\uchwały - pdf"
$dest_path = $cwd + "\uchwały - jpg"
$counter = 1

Get-ChildItem -Path $source_path -Filter *.pdf | Sort-Object { [int]$_.Name.Split(' ')[0] } | ForEach-Object {
  Write-Output $_.BaseName

  $jpg_filename = "$dest_path\$($counter).jpg"
  $qwer = $_.FullName
  
  naps2.console -i "`"$qwer`"" -o "`"$jpg_filename`"" -n 0

  $counter += 1
}
