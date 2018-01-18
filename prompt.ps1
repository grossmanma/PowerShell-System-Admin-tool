
 
 $cred = Get-Credential
 $pp = new-object -typename System.Management.Automation.PSCredential -argumentlist $cred
 $script = "c:\windows\System32\wscript.exe  .\invisible.vbs"
 Start-Process powershell -Credential $cred -ArgumentList ' -command &{Start-Process c:\windows\System32\wscript.exe  .\invisible.vbs  -verb runas}'

