'-----------------------
'--Check Java Version---
'-----------------------

Dim RegPath, RegPath64
Dim JavaVersion
Dim OsBits
RegPath = "HKEY_LOCAL_MACHINE\SOFTWARE\JavaSoft\Java Runtime Environment\CurrentVersion"
RegPath64 = "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\JavaSoft\Java Runtime Environment\CurrentVersion"
Set objShell = CreateObject("WScript.Shell")
Set service = GetObject("winmgmts:")

'== Main ==

install

'==========

Sub install
On Error Resume Next

'=== Check System ===
OsBits = OsBit()
If OsBits = "x86" Then
	JavaVersion = objShell.RegRead (RegPath)
Else 
	JavaVersion = objShell.RegRead (RegPath64)
End If 

'====== Check Java Version and install =======
	If err.number <> 0 Then
	
		'==== install JRE 1.6 if JRE 1.6 does not exist ====
        Return = objShell.Run ("jre-6u31-windows-i586.exe /s /v /qn" & WScript.ScriptFullName,1,True)
    Else
    	If JavaVersion = 1.6 Then
    	'=== Do not install JRE 1.6 ===
    	Else 
    	
    	'=== Install JRE 1.6 if jre not equal 1.6  (x < 1.6 < x) =====
    		Return = objShell.Run ("jre-6u31-windows-i586.exe /s /v /qn" & WScript.ScriptFullName ,1,True)
    		
    			'=== Registry write to old JRE version ===
    			If OsBits = "x86" Then
    				objShell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\JavaSoft\Java Runtime Environment\CurrentVersion", JavaVersion,  "REG_SZ"
    			Else 
    				objShell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\JavaSoft\Java Runtime Environment\CurrentVersion", JavaVersion,  "REG_SZ"
    			End If
    	End If
    End If
End Sub	

Function OsBit()
	const HKEY_LOCAL_MACHINE = &H80000002
	strComputer = "."

	Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
	strComputer & "\root\default:StdRegProv")

		strKeyPath = "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
		strValueName = "PROCESSOR_ARCHITECTURE"
		oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
		OsBit = strValue
End Function