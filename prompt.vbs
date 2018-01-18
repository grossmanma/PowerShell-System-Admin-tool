Set WshShell = CreateObject("WScript.Shell" ) 
 WshShell.Run chr(34) & ".\prompt.cmd" & Chr(34), 0 
 Set WshShell = Nothing 
