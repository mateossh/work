# https://stackoverflow.com/questions/16534292/basic-powershell-batch-convert-word-docx-to-pdf
#
# save this file with "UTF-8 with BOM"

$cwd = (Get-Item .).FullName
$source_path = $cwd + "\uchwały"
$dest_path = $cwd + "\uchwały - pdf"

$word_app = New-Object -ComObject Word.Application

# This filter will find .doc as well as .docx documents
Get-ChildItem -Path $source_path -Filter *.doc? | ForEach-Object {
  Write-Output $_.BaseName

  $pdf_filename = "$dest_path\$($_.BaseName).pdf"

  $document = $word_app.Documents.Open($_.FullName)
  $document.SaveAs([ref] $pdf_filename, [ref] 17)
  $document.Close()

}

$word_app.Quit()
