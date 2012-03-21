'-----------------------
'--Check Java Version---
'-----------------------

'== Main ==

Install

'==========

Sub Install
    On Error Resume Next
    
    Dim regPath, regPath64
    Dim javaVersion
    Dim osArch
    regPath = "HKEY_LOCAL_MACHINE\SOFTWARE\JavaSoft\Java Runtime Environment\CurrentVersion"
    regPath64 = "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\JavaSoft\Java Runtime Environment\CurrentVersion"
    Set objShell = CreateObject("WScript.Shell")
    Set service = GetObject("winmgmts:")
    
    '=== Check System ===
    osArch = GetOsArch()
    If osArch = "x86" Then
        javaVersion = objShell.RegRead (regPath)
    Else 
        javaVersion = objShell.RegRead (regPath64)
    End If 
    
    '====== Check Java Version and install =======
    If err.number <> 0 Then
    
    	'==== install JRE 1.6 if JRE 1.6 does not exist ====
        Return = objShell.Run ("jre-6u31-windows-i586.exe /s /v /qn" & WScript.ScriptFullName, 1, True)
    Else
        If javaVersion <> 1.6 Then
			
            '=== Install JRE 1.6 if jre not equal 1.6  (x < 1.6 < x) =====
            Return = objShell.Run ("jre-6u31-windows-i586.exe /s /v /qn" & WScript.ScriptFullName , 1, True)
			
            '=== Registry write to old JRE version ===
            If osArch = "x86" Then
                objShell.RegWrite regPath, javaVersion, "REG_SZ"
            Else 
                objShell.RegWrite regPath64, javaVersion, "REG_SZ"
            End If
        End If
    End If
End Sub	

Function GetOsArch()
    const HKEY_LOCAL_MACHINE = &H80000002
    strComputer = "."
    
    Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
    strComputer & "\root\default:StdRegProv")
    
    strKeyPath = "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
    strValueName = "PROCESSOR_ARCHITECTURE"
    oReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue
    	
    GetOsArch = strValue
End Function