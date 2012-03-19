'-----------------------
'--Check Java Version---
'-----------------------

Dim RegJavaVersion
Dim Jversion

RegJavaVersion = "HKEY_LOCAL_MACHINE\SOFTWARE\JavaSoft\Java Runtime Environment\CurrentVersion"
Set objShell = CreateObject("WScript.Shell")

On Error Resume Next
Jversion = objShell.RegRead (RegJavaVersion)
	if err.number <> 0 then
        WScript.Echo "Error"
    Else
    	WScript.Echo Jversion
    End If


