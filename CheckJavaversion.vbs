'-----------------------
'--Check Java Version---
'-----------------------

Dim RegJavaVersion
Dim Jversion

RegJavaVersion = "HKEY_LOCAL_MACHINE\SOFTWARE\JavaSoft\Java Runtime Environment\CurrentVersion"
Set objShell = CreateObject("WScript.Shell")
Sub install
On Error Resume Next
Jversion = objShell.RegRead (RegJavaVersion)
	if err.number <> 0 then
        objShell.Run ("jre.exe /s /v /qn")
	WScript.Sleep 30*1000
    Else
    	If Jversion < 1.6 Then
    		'WScript.Echo RegJavaVersion & "  " & Jversion
    		objShell.Run ("jre.exe /s /v /qn")
    		WScript.Sleep 30*1000
    		
    		objShell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\JavaSoft\Java Runtime Environment\CurrentVersion", Jversion,  "REG_SZ"
    	'	Jversion = objShell.RegRead (RegJavaVersion)
    	'	WScript.Echo RegJavaVersion & "  " & Jversion
    	Else 
    		WScript.Echo Jversion & " = JRE 1.6 is exis"
    	End If
    End If

End Sub	
install()
'objShell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\JavaSoft\Java Runtime Environment\CurrentVersion", Jversion , "REG_SZ"
Jversion = objShell.RegRead (RegJavaVersion)
    		WScript.Echo Jversion