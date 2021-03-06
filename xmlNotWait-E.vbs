'---------------- Wait for openapp running  --------------
'WScript.Sleep 1*60*1000  'sleep 2 min


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
		If Process.Name = "winoapp.exe" then
			i = 1
		End If
	Next
	If i = 1 Then
		mainCheck()
	Else
		'objShell.Exec "shutdown -l"
		mainCheck()
	End If


Sub mainCheck

'WScript.Sleep 5*60*1000

For n = 1 To 999
	i=0
	For each Process in Service.InstancesOf ("Win32_Process")
		If Process.Name = "winoapp.exe" then
			'WScript.Sleep 1*60*1000    'sleep 1 min
			i = 1
		End If
	Next
	If i = 0 Then
		CheckProgramOpenapp()
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
    CreateXML(pathXP)
	CheckProgram32xp(strKey64)
	CheckProgram32xp(strKey32)
	CloseXML(pathXP)
	Sendlog(pathXP)
Else 
	If x = xp32 Then
	   CreateXML(pathXP)
	   CheckProgram32xp(strKey32)
	   CloseXML(pathXP)		
	   Sendlog(pathXP)
	Else 
		If bit = "x86" Then
		    CreateXML(pathOuther)
			CheckProgram32(strKey32)
			CloseXML(pathOuther)
			Sendlog(pathOuther)
		Else
		    CreateXML(pathOuther)
			CheckProgram32(strKey32)
			CheckProgram32(strKey64)
			CloseXML(pathOuther)
			Sendlog(pathOuther)
		End If
	End If
End If	
	WScript.Quit    'Exit all this script
	
End Sub	
'---------------------------------------------------------



'--------------- check Program for Windows XP -----------------
Sub CheckProgram32xp(strKey)

OSV = OsVersion()
BitOsv = Bitos()
OpenAppXml = pathXP & OSV & BitOsv & "openapp.xml" 
		Set objTextFile = objFSO.OpenTextFile _
			(OpenAppXml, ForAppending, True)
'OpenAppLog = pathXP & "openapp.xml" 
'		Set objTextFile = objFSO.OpenTextFile _
'			(OpenAppLog, ForAppending, True)
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
        				objTextFile.WriteLine  vbTab & vbTab & "<ProgramList" & p & ">" & strValue1 & "</ProgramList" & p & ">"
        			End If
        		Next
            
        End If
    Next
End Sub




'------------- Check Program for Vista and 7 ----------------
Sub CheckProgram32(strKey)
OSV = OsVersion()
BitOsv = Bitos()
OpenAppXml = pathOuther & OSV & BitOsv & "openapp.xml" 
		Set objTextFile = objFSO.OpenTextFile _
			(OpenAppXml, ForAppending, True)
'OpenAppLog = pathOuther & "openapp.xml" 
'		Set objTextFile = objFSO.OpenTextFile _
'			(OpenAppLog, ForAppending, True)
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
        				objTextFile.WriteLine vbTab & vbTab & "<ProgramList" & p & ">" & strValue1 & "</ProgramList" & p & ">"
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
	Dim OSVV
	WScript.Echo OSV
	If OSV = "Microsoft Windows 7 Professional " Then
	   OSVV = "MicrosoftWindows-7-Professional"
	Else
	   If OSV = "Microsoft Windows XP Professional" Then
	       OSVV = "MicrosoftWindows-XP-Professional"
	   Else
	       If OSV = "Microsoft(R) Windows(R) XP Professional x64 Edition" Then
	           OSVV = "Microsoft(R)Windows(R)-XP-Professionalx64Edition"
	       Else
	               OSVV = "MicrosoftWindows-Vista-Business"
	       End If
	   End If        
	End If
	PathLog = Path & OSV & BitOsv & "openapp.xml" 
	PathCp = "C:\xml\" & OSVV & BitOsv & "-openapp.xml" 
	Wscript.echo PathLog
	Wscript.echo PathCp
	
	strFolder = "C:\xml"

SET objFSO = CREATEOBJECT("Scripting.FileSystemObject")

IF objFSO.FolderExists(strFolder) = FALSE THEN
	objFSO.CreateFolder strFolder
	wscript.echo "Folder Created"
ELSE
	wscript.echo "Folder already exists"
END IF
	objFSO.CopyFile PathLog, PathCp
    WScript.Echo FTPUpload("192.168.23.99","xmlftp","1234567890",PathCp,"/var/www/xml")
End Sub

Sub CreateXML(Path)
OSV = OsVersion()
BitOsv = Bitos()
OpenAppXml = Path & OSV & BitOsv & "openapp.xml" 
		Set objXMLFile = objFSO.OpenTextFile _
			(OpenAppXml, ForAppending, True)
objXMLFile.WriteLine "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
objXMLFile.WriteLine "<?xml-stylesheet type=""text/xsl""?>"
objXMLFile.WriteLine "<DOCUMENTS>"
'objXMLFile.WriteLine vbTab & "<DOCUMENT>"
End Sub

Sub CloseXML(Path)
OSV = OsVersion()
BitOsv = Bitos()
OpenAppXml = Path & OSV & BitOsv & "openapp.xml" 
		Set objXMLFile = objFSO.OpenTextFile _
			(OpenAppXml, ForAppending, True)
'objXMLFile.WriteLine vbTab & "</DOCUMENT>"
objXMLFile.WriteLine "</DOCUMENTS>" 
End Sub


Function FTPUpload(sSite, sUsername, sPassword, sLocalFile, sRemotePath)
  'This script is provided under the Creative Commons license located
  'at http://creativecommons.org/licenses/by-nc/2.5/ . It may not
  'be used for commercial purposes with out the expressed written consent
  'of NateRice.com

  Const OpenAsDefault = -2
  Const FailIfNotExist = 0
  Const ForReading = 1
  Const ForWriting = 2
  
  Set oFTPScriptFSO = CreateObject("Scripting.FileSystemObject")
  Set oFTPScriptShell = CreateObject("WScript.Shell")

  sRemotePath = Trim(sRemotePath)
  sLocalFile = Trim(sLocalFile)
  
  '----------Path Checks---------
  'Here we willcheck the path, if it contains
  'spaces then we need to add quotes to ensure
  'it parses correctly.
  If InStr(sRemotePath, " ") > 0 Then
    If Left(sRemotePath, 1) <> """" And Right(sRemotePath, 1) <> """" Then
      sRemotePath = """" & sRemotePath & """"
    End If
  End If
  
  If InStr(sLocalFile, " ") > 0 Then
    If Left(sLocalFile, 1) <> """" And Right(sLocalFile, 1) <> """" Then
      sLocalFile = """" & sLocalFile & """"
    End If
  End If

  'Check to ensure that a remote path was
  'passed. If it's blank then pass a "\"
  If Len(sRemotePath) = 0 Then
    'Please note that no premptive checking of the
    'remote path is done. If it does not exist for some
    'reason. Unexpected results may occur.
    sRemotePath = "\"
  End If
  
  'Check the local path and file to ensure
  'that either the a file that exists was
  'passed or a wildcard was passed.
  If InStr(sLocalFile, "*") Then
    If InStr(sLocalFile, " ") Then
      FTPUpload = "Error: Wildcard uploads do not work if the path contains a " & _
      "space." & vbCRLF
      FTPUpload = FTPUpload & "This is a limitation of the Microsoft FTP client."
      Exit Function
    End If
  ElseIf Len(sLocalFile) = 0 Or Not oFTPScriptFSO.FileExists(sLocalFile) Then
    'nothing to upload
    FTPUpload = "Error: File Not Found."
    Exit Function
  End If
  '--------END Path Checks---------
  
  'build input file for ftp command
  sFTPScript = sFTPScript & "USER " & sUsername & vbCRLF
  sFTPScript = sFTPScript & sPassword & vbCRLF
  sFTPScript = sFTPScript & "cd " & sRemotePath & vbCRLF
  sFTPScript = sFTPScript & "binary" & vbCRLF
  sFTPScript = sFTPScript & "prompt n" & vbCRLF
  sFTPScript = sFTPScript & "put " & sLocalFile & vbCRLF
  sFTPScript = sFTPScript & "quit" & vbCRLF & "quit" & vbCRLF & "quit" & vbCRLF


  sFTPTemp = oFTPScriptShell.ExpandEnvironmentStrings("%TEMP%")
  sFTPTempFile = sFTPTemp & "\" & oFTPScriptFSO.GetTempName
  sFTPResults = sFTPTemp & "\" & oFTPScriptFSO.GetTempName

  'Write the input file for the ftp command
  'to a temporary file.
  Set fFTPScript = oFTPScriptFSO.CreateTextFile(sFTPTempFile, True)
  fFTPScript.WriteLine(sFTPScript)
  fFTPScript.Close
  Set fFTPScript = Nothing  

  oFTPScriptShell.Run "%comspec% /c FTP -n -s:" & sFTPTempFile & " " & sSite & _
  " > " & sFTPResults, 0, TRUE
  
  Wscript.Sleep 1000
  
  'Check results of transfer.
  Set fFTPResults = oFTPScriptFSO.OpenTextFile(sFTPResults, ForReading, _
  FailIfNotExist, OpenAsDefault)
  sResults = fFTPResults.ReadAll
  fFTPResults.Close
  
  oFTPScriptFSO.DeleteFile(sFTPTempFile)
  oFTPScriptFSO.DeleteFile (sFTPResults)
  
  If InStr(sResults, "226 Transfer complete.") > 0 Then
    FTPUpload = True
  ElseIf InStr(sResults, "File not found") > 0 Then
    FTPUpload = "Error: File Not Found"
  ElseIf InStr(sResults, "cannot log in.") > 0 Then
    FTPUpload = "Error: Login Failed."
  Else
    FTPUpload = "Error: Unknown."
  End If

  Set oFTPScriptFSO = Nothing
  Set oFTPScriptShell = Nothing
End Function

