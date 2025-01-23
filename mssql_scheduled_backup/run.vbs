' https://marte-it.at/en/start-powershell-script-hidden-via-task-scheduler/

command = "powershell.exe -nologo -command C:\SQLBackups\kopia.ps1"
set shell = CreateObject("WScript.Shell")
shell.Run command,0