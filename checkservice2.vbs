'---------------- Wait for openapp running  --------------
WScript.Sleep 2*60*1000  'sleep 15 min


set service = GetObject ("winmgmts:")

Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
'------- Path log -----------
pathXP = "C:\Documents and Settings\All Users\Application Data\OpenApp\"
pathOuther = "C:\ProgramData\OpenApp\"

'------- Path Registry ------------
strComputer = "."
strKey64 = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
strKey32 = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"

strEntry1a = "DisplayName"
strEntry1b = "QuietDisplayName"

Dim p
	p = 0
Dim objShell,dos
	Set objShell = WScript.CreateObject("WScript.shell")
	
Const ForAppending = 8
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	

'------- List Program ---------------------
Dim programList(13)
	programList(0) = "InfraRecorder"
	programList(1) = "Inkscape 0.48.2"
	programList(2) = "Mozilla Thunderbird (8.0)"
	programList(3) = "GIMP 2.6.11"
	programList(4) = "PDFCreator"
	programList(5) = "7-Zip 9.20"
	programList(6) = "Java(TM) 6 Update 29"
	programList(7) = "Microsoft Visual C++ 2008 Redistributable - x86 9.0.30729.17"
	programList(8) = "LibreOffice 3.3"
	programList(9) = "OpenApp"
	programList(10) = "7-Zip 9.20 (x64 edition)"
	programList(11) = "InfraRecorder 0.52 (x64 edition)"
	programList(12) = "Microsoft Visual C++ 2008 Redistributable - x64 9.0.30729.17"



'----------------------------Main-------------------------------------
	For each Process in Service.InstancesOf ("Win32_Process")
		If Process.Name = "notepad.exe" then
			i = 1
		End If
	Next
	If i = 1 Then
		mainCheck()
	Else
		objShell.Exec "shutdown -L"
	End If

'-----------------------------------------------------------------------


Sub mainCheck

For n = 1 To 999
	i=0
	For each Process in Service.InstancesOf ("Win32_Process")
		If Process.Name = "notepad.exe" then
			WScript.echo "Notepad running"
			'WScript.Sleep 1*60*1000    'sleep 1 min
			i = 1
		End If
	Next
	If i = 0 Then
		WScript.Echo "Notepad not running"
		'CheckProgramOpenapp()
	End If
	
Next

End Sub

'----------------------------------------------------------------------




'---------------  Check Windows version ------------------------

Sub CheckProgramOpenapp()
Dim objTextFile
x = OsVersion()
xp64 = "Microsoft(R) Windows(R) XP Professional x64 Edition"
xp32 = "Microsoft Windows XP Professional"
bit = Bitos()
If x = xp64 Then
	CheckProgram32xp(strKey64)
	CheckProgram32xp(strKey32)
	Sendlog(pathXP)
	'WScript.Echo x
	'WScript.Echo bit
Else 
	If x = xp32 Then
		CheckProgram32xp(strKey32)
		Sendlog(pathXP)
		'WScript.Echo x
		'WScript.Echo bit
	Else 
		If bit = "x86" Then
			CheckProgram32(strKey32)
			Sendlog(pathOuther)
			'WScript.Echo bit
		Else
			CheckProgram32(strKey32)
			CheckProgram32(strKey64)
			Sendlog(pathOuther)
			'WScript.Echo bit
		End If
	End If
	 
End If	

	WScript.Echo p
	WScript.Quit    'Exit all this script
	
End Sub	


'---------------------------------------------------------



'--------------- check Program for Windows XP -----------------
Sub CheckProgram32xp(strKey)
	
OpenAppLog = pathXP & "openapp.log" 
		Set objTextFile = objFSO.OpenTextFile _
			(OpenAppLog, ForAppending, True)
Set objReg = GetObject("winmgmts://" & strComputer & "/root/default:StdRegProv")
    objReg.EnumKey HKLM, strKey, arrSubkeys
    'WScript.Echo "Installed Applications" & VbCrLf
    For Each strSubkey In arrSubkeys
        intRet1 = objReg.GetStringValue(HKLM, strKey & strSubkey, strEntry1a, strValue1)
        If intRet1 <> 0 Then
            objReg.GetStringValue HKLM, strKey & strSubkey, strEntry1b, strValue1
        End If
        If strValue1 <> "" Then
        		For i = 0 To 12
        			If strValue1 = programList(i) Then
        				p = p + 1
        				objTextFile.WriteLine(strValue1)
        			End If
        		Next
            
        End If
    Next
    'CheckProgram32 = "0"
    'WScript.Echo p
End Sub




'------------- Check Program for Vista and 7 ----------------
Sub CheckProgram32(strKey)
	
OpenAppLog = pathOuther & "openapp.log" 
		Set objTextFile = objFSO.OpenTextFile _
			(OpenAppLog, ForAppending, True)
Set objReg = GetObject("winmgmts://" & strComputer & "/root/default:StdRegProv")
    objReg.EnumKey HKLM, strKey, arrSubkeys
    For Each strSubkey In arrSubkeys
        intRet1 = objReg.GetStringValue(HKLM, strKey & strSubkey, strEntry1a, strValue1)
        If intRet1 <> 0 Then
            objReg.GetStringValue HKLM, strKey & strSubkey, strEntry1b, strValue1
        End If
        If strValue1 <> "" Then
        		For i = 0 To 12
        			If strValue1 = programList(i) Then
        				p = p + 1
        				objTextFile.WriteLine(strValue1)
        			End If
        		Next
            
        End If
    Next
End Sub


'----------- Check bit of Windows ---------
Function Bitos()
	const HKEY_LOCAL_MACHINE = &H80000002
	strComputer = "."

	Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
	strComputer & "\root\default:StdRegProv")

		strKeyPath = "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
		strValueName = "PROCESSOR_ARCHITECTURE"
		oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
		Bitos = strValue
End Function

'-------------- Check Windows Version ------------------------
Function OsVersion()
	Dim v
	Set SystemSet = GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem") 
	For each System in SystemSet 
		' WScript.Echo System.Caption
		v = System.Caption 
	Next 
		Osversion = v
End Function


'----------  Send Log to Mail --------------
Sub  Sendlog(Path)
	OSV = OsVersion()
	BitOsv = Bitos()
	PathLog = Path & "openapp.log"
	PathCp = Path & OSV & " " & BitOsv & " Alllog.log"
	
	objFSO.CopyFile PathLog, PathCp

	Set objMessage = CreateObject("CDO.Message") 
	objMessage.Subject = "Example CDO Message" 
	objMessage.From = "tapootumchannel@gmail.com" 
	objMessage.To = "tapootumchannel@gmail.com" 
	objMessage.TextBody = "This is some sample message text."
	objMessage.AddAttachment PathCp
	'objMessage.AddAttachment "C:\ProgramData\OpenApp\tt.txt"
	'==This section provides the configuration information for the remote SMTP server.
	'==Normally you will only change the server name or IP.
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
	'Name or IP of Remote SMTP Server
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
	'Server port (typically 25)
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendusername") = "tapootumchannel"
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "tapootum15124"
	objMessage.Configuration.Fields.Update

	'==End remote SMTP server configuration section==

	objMessage.Send
End Sub
