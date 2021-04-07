$cdkey = (Get-WmiObject -query 'select * from SoftwareLicensingService').OA3xOriginalProductKey

Write-Output "Windows Key: $($cdkey)"