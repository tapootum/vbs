'==========================================
'VBScript: enableAutoLogon.vbs            =
'This VBScript updates the registry to    =
'enable auto-logon.  Modify the three     =
'strings in brackets, under "Define       =
'keys and values".                        =
'Courtesy of Jonathan Almquist            =
'monsterjta @ tek-tips                    =
'==========================================

Option Explicit
''
'Declarations'
''
Dim objShell
Dim RegLocAutoLogon
Dim keyDefaultDomainName
Dim valDefaultDomainName
Dim keyDefaultUserName
Dim valDefaultUserName
Dim keyDisableCAD
Dim valDisableCAD
Dim keyAutoAdminLogon
Dim valAutoAdminLogon
Dim keyForceAutoLogon
Dim valForceAutoLogon
Dim keyDefaultPassword
Dim valDefaultPassword

'''
'Define keys and values'
'''
RegLocAutoLogon = "HKLM\Software\Microsoft\" & _
"Windows NT\CurrentVersion\Winlogon\"
keyDefaultDomainName = "DefaultDomainName"
valDefaultDomainName = "xxxxxxx"
keyDefaultUserName = "DefaultUserName"
valDefaultUserName = "xxxxxx"
keyDisableCAD = "DisableCAD"
valDisableCAD = 1
keyAutoAdminLogon = "AutoAdminLogon"
valAutoAdminLogon = "1"
keyForceAutoLogon = "ForceAutoLogon"
valForceAutoLogon = "1"
keyDefaultPassword = "DefaultPassword"
valDefaultPassword = "xxxxxx"

Set objShell = CreateObject("WScript.Shell")

objShell.RegWrite RegLocAutoLogon & _
keyDefaultDomainName, 1, "REG_SZ"
objShell.RegWrite RegLocAutoLogon & _
keyDefaultDomainName, valDefaultDomainName, "REG_SZ"
objShell.RegWrite RegLocAutoLogon & _
keyDefaultUserName, 1, "REG_SZ"
objShell.RegWrite RegLocAutoLogon & _
keyDefaultUserName, valDefaultUserName, "REG_SZ"
objShell.RegWrite RegLocAutoLogon & _
keyDisableCAD, 1, "REG_DWORD"
objShell.RegWrite RegLocAutoLogon & _
keyDisableCAD, valDisableCAD, "REG_DWORD"
objShell.RegWrite RegLocAutoLogon & _
keyAutoAdminLogon, 1, "REG_SZ"
objShell.RegWrite RegLocAutoLogon & _
keyAutoAdminLogon, valAutoAdminLogon, "REG_SZ"
objShell.RegWrite RegLocAutoLogon & _
keyForceAutoLogon, 1, "REG_SZ"
objShell.RegWrite RegLocAutoLogon & _
keyForceAutoLogon, valForceAutoLogon, "REG_SZ"
objShell.RegWrite RegLocAutoLogon & _
keyDefaultPassword, 1, "REG_SZ"
objShell.RegWrite RegLocAutoLogon & _
keyDefaultPassword, valDefaultPassword, "REG_SZ"