Set WshShell = CreateObject("WScript.Shell") 
WshShell.Run chr(34) & "C:\Program Files\Common Files\System\first\starter.bat" & Chr(34), 0
Set WshShell = Nothing
