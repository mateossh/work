# Work

Work related scripts / notes

Remember about `Set-ExecutionPolicy`

---

## ip_checker.py

Log IP address every 30 seconds to `plik.txt`

Before you run: `pip install requests`

## cdkey.ps1

Get Windows Product Key from ACPI tables

## clear_spool.ps1

Force clear printer spool in Windows 7+

## open_pdf.ps1

Opens each .pdf file from directory next to script in a Chrome incognito window

## remux.ps1

Convert each .flv file to .mp4 in directory. Requires `ffmpeg` in $PATH

## tightvnc.md

How to setup projector preview on a machine connected to LAN

## debestiofikator.bas

Word macro - "Fix" documents generated by BeSTi@ (Macros don't work the way you think they would btw)

## certs.md

zxcv


---

## Export "third-party" drivers

`dism /online /export-driver /destination:c:\exported-drivers`

## .NET 3.5 installation with no internet

`dism /Online /Enable-Feature /FeatureName:NetFx3 /All /LimitAccess /Source:x:\sources\sxs` (where x: is installation drive)

or

`dism /Online /Add-Package /PackagePath:"X:\sources\sxs\Microsoft-Windows-NetFx3-OnDemand-Package.cab"` (where x: is installation drive)

## Paragraph symbol (Windows)

`ALT + 0167` -> `§`

## Count files in folder PowerShell

`(Get-ChildItem -Filter *.pdf -Recurse).FullName | measure`

similar result with unix tools

`find . -name "*pdf" | wc -l`
