# https://stackoverflow.com/questions/16534292/basic-powershell-batch-convert-word-docx-to-pdf
#https://superuser.com/questions/318197/how-do-i-get-get-childitem-to-filter-on-multiple-file-types
#
# save this file with "UTF-8 with BOM"

$cwd = (Get-Item .).FullName
$source_path = $cwd + "\uchwały"
$dest_path = $cwd + "\uchwały - pdf"

$word_app = New-Object -ComObject Word.Application

Get-ChildItem -Path $source_path | Where-Object { $_.extension -in ".doc",".docx",".odt" } | Sort-Object { [int]$_.Name.Split(' ')[0] } | ForEach-Object {
  Write-Output $_.BaseName

  $pdf_filename = "$dest_path\$($_.BaseName).pdf"

  $document = $word_app.Documents.Open($_.FullName)
  $document.SaveAs([ref] $pdf_filename, [ref] 17)
  $document.Close()
}

$word_app.Quit()
