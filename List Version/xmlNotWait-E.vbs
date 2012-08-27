'---------------- Wait for openapp running  --------------
'WScript.Sleep 1*60*1000  'sleep 1 min


Set service = GetObject ("winmgmts:")

Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
'------- Path log -----------
pathXP = "C:\Documents and Settings\All Users\Application Data\OpenApp\"
pathOuther = "C:\ProgramData\OpenApp\"

'------- Path Registry ------------
strComputer = "."
strKey64 = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
strKey32 = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"

'------- Registry Variabile ---------
strEntry1a = "DisplayName"
strEntry1b = "QuietDisplayName"
strEntry1d = "DisplayVersion"

Dim p
p = 0
Dim objShell,dos
Set objShell = WScript.CreateObject("WScript.shell")

Const ForAppending = 8
Set objFSO = CreateObject("Scripting.FileSystemObject")



'------- List Program ---------------------
Dim programList(14)
programList(0) = "InfraRecorder"
programList(1) = "Inkscape 0.48.2"
programList(2) = "Mozilla Thunderbird 11.0.1 (x86 en-US)"
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
programList(13) = "OpenFonts"


'----------------------------Main-------------------------------------
For Each Process In Service.InstancesOf ("Win32_Process")
     If Process.Name = "winoapp.exe" Then
          i = 1
     End If
Next
If i = 1 Then
     MainCheck()
Else
     objShell.Exec "shutdown -l"
End If
objShell.Exec "shutdown -l"    'if program not running , computer is log off'

'-------- Sub Main ------
Sub MainCheck
     WScript.Sleep 1*60*1000
     For n = 1 To 999
          i=0
          '---- check openapp process ----
          For Each Process In Service.InstancesOf ("Win32_Process")
               If Process.Name = "winoapp.exe" Then
                    WScript.Sleep 1*60*1000    'sleep 1 min
                    i = 1
               End If
          Next
          '---if i=0 openapp is end process
          If i = 0 Then
               CheckProgramOpenapp()
          End If
     Next
End Sub
'---------------------------

'------ Check windows version --------

Sub CheckProgramOpenapp()
     Dim objTextFile
     Dim oOsVersion
     oOsVersion = OsVersion()
     xp64 = "Microsoft(R) Windows(R) XP Professional x64 Edition"
     xp32 = "Microsoft Windows XP Professional"
     oOsArch = OsArch()
     If oOsVersion = xp64 Then
          CreateXML(pathXP)
          CheckProgramXP(strKey64)
          CheckProgramXP(strKey32)
          CloseXML(pathXP)
          Sendlog(pathXP)
     Else 
          If oOsVersion = xp32 Then
               CreateXML(pathXP)
               CheckProgramXP(strKey32)
               CloseXML(pathXP)   
               Sendlog(pathXP)
          Else 
               If oOsArch = "x86" Then
                    CreateXML(pathOuther)
                    CheckProgram(strKey32)
                    CloseXML(pathOuther)
                    Sendlog(pathOuther)
               Else
                    CreateXML(pathOuther)
                    CheckProgram(strKey32)
                    CheckProgram(strKey64)
                    CloseXML(pathOuther)
                    Sendlog(pathOuther)
               End If
          End If
     End If

     objShell.Exec "shutdown -s"  'Script complete is shut down VM'
     WScript.Quit    'Exit all this script
End Sub
'----------------------------------------

'--------------- Check Program for Windows XP -----------------
Sub CheckProgramXP(strKey)
     oOsVersion = OsVersion()
     oOsArch = OsArch()
     OpenAppXml = pathXP & oOsVersion & oOsArch & "openapp.xml" 
     Set objTextFile = objFSO.OpenTextFile _
     (OpenAppXml, ForAppending, True)
     Set objReg = GetObject("winmgmts://" & strComputer & "/root/default:StdRegProv")
     objReg.EnumKey HKLM, strKey, arrSubkeys
     For Each strSubkey In arrSubkeys
          intRet1 = objReg.GetStringValue(HKLM, strKey & strSubkey, strEntry1a, strValue1)
          objReg.GetStringValue HKLM, strKey & strSubkey, strEntry1d, strValue2
               'If intRet1 <> 0 Then
               'End If
               If strValue1 <> "" Then
               'WScript.Echo strValue1
                    For i = 0 To 13
                         If strValue1 = programList(i) Then
                              If strValue1 = "InfraRecorder" Then
                                   p = p + 1
                                   vInfraRecorder = InfraRecorder()
                                   objTextFile.WriteLine vbTab & vbTab & "<ProgramList" & p & ">" & _
                                   strValue1 & " V."& vInfraRecorder & "</ProgramList" & p & ">"
                              End If
                              If strValue2 <> "" Then
                              p = p + 1
                              objTextFile.WriteLine  vbTab & vbTab & "<ProgramList" & p & ">" & _
                              strValue1 & " V." & strValue2 &"</ProgramList" & p & ">"
                              End If
                         End If
                    Next
               End If
     Next
End Sub

'------------- Check Program for Vista and 7 ----------------
Sub CheckProgram(strKey)
     oOsVersion = OsVersion()
     oOsArch = OsArch()
     OpenAppXml = pathOuther & oOsVersion & oOsArch & "openapp.xml" 
     Set objTextFile = objFSO.OpenTextFile _
     (OpenAppXml, ForAppending, True)
     Set objReg = GetObject("winmgmts://" & strComputer & "/root/default:StdRegProv")
     objReg.EnumKey HKLM, strKey, arrSubkeys
     For Each strSubkey In arrSubkeys
          intRet1 = objReg.GetStringValue(HKLM, strKey & strSubkey, strEntry1a, strValue1)
          objReg.GetStringValue HKLM, strKey & strSubkey, strEntry1d, strValue2
          'If intRet1 <> 0 Then
          'End If
          If strValue1 <> "" Then
               For i = 0 To 13
                    If strValue1 = programList(i) Then
                         If strValue1 = "InfraRecorder" Then
                              p = p + 1
                              vInfraRecorder = InfraRecorder()
                              'WScript.Echo strValue1 & vInfraRecorder
                              objTextFile.WriteLine vbTab & vbTab & "<ProgramList" & p & ">" & _
                              strValue1 & " V."& vInfraRecorder & "</ProgramList" & p & ">"
                         End If
                         If strValue2 <> "" Then
                              p = p + 1
                              'WScript.Echo strValue1 & strValue2
                              objTextFile.WriteLine vbTab & vbTab & "<ProgramList" & p & ">" & _
                              strValue1 & " V."& strValue2 & "</ProgramList" & p & ">"
                         End If
                    End If
               Next
          End If
     Next
End Sub


'----------- Check Os Arch ---------
Function OsArch()
     Const HKEY_LOCAL_MACHINE = &H80000002
     strComputer = "."
     
     Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
     strComputer & "\root\default:StdRegProv")
     
     strKeyPath = "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
     strValueName = "PROCESSOR_ARCHITECTURE"
     oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
     OsArch = strValue
End Function

'-------------- Check Windows Version ------------------------
Function OsVersion()
     Dim v
     Set SystemSet = GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem") 
     For Each System In SystemSet 
          ' WScript.Echo System.Caption
          v = System.Caption 
     Next 
     OsVersion = v
End Function

'----------  Send Log to FTP --------------
Sub  Sendlog(Path)
     oOsVersion = OsVersion()
     oOsArch = OsArch()
     Dim osVersionV
     'WScript.Echo oOsVersion
     If oOsVersion = "Microsoft Windows 7 Professional " Then
          osVersionV = "MicrosoftWindows-7-Professional"
     Else
          If oOsVersion = "Microsoft Windows XP Professional" Then
               osVersionV = "MicrosoftWindows-XP-Professional"
          Else
               If oOsVersion = "Microsoft(R) Windows(R) XP Professional x64 Edition" Then
                    osVersionV = "Microsoft(R)Windows(R)-XP-Professionalx64Edition"
               Else
                    osVersionV = "MicrosoftWindows-Vista-Business"
               End If
          End If        
     End If
     pathLog = Path & oOsVersion & oOsArch & "openapp.xml" 
     pathCp = "C:\xml\" & osVersionV & oOsArch & "-openapp.xml" 
     'WScript.echo PathLog
     'WScript.echo PathCp
     
     strFolder = "C:\xml"
     Set objFSO = CreateObject("Scripting.FileSystemObject")
     
     If objFSO.FolderExists(strFolder) = False Then
          objFSO.CreateFolder strFolder
          'WScript.echo "Folder Created"
     'Else
          'WScript.echo "Folder already exists"
     End If
     objFSO.CopyFile pathLog, pathCp
     ftp = FTPUpload("192.168.23.99","xmlftp","1234567890",PathCp,"/var/www/xml")
End Sub

Sub CreateXML(Path)
     oOsVersion = OsVersion()
     oOsArch = OsArch()
     oOpenAppXml = Path & oOsVersion & oOsArch & "openapp.xml" 
     Set objXMLFile = objFSO.OpenTextFile _
     (oOpenAppXml, ForAppending, True)
     objXMLFile.WriteLine "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
     objXMLFile.WriteLine "<?xml-stylesheet type=""text/xsl""?>"
     objXMLFile.WriteLine "<DOCUMENTS>"
     'objXMLFile.WriteLine vbTab & "<DOCUMENT>"
End Sub

Sub CloseXML(Path)
     oOsVersion = OsVersion()
     oOsArch = OsArch()
     oOpenAppXml = Path & oOsVersion & oOsArch & "openapp.xml" 
     Set objXMLFile = objFSO.OpenTextFile _
     (oOpenAppXml, ForAppending, True)
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
               "space." & vbCrLf
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
     sFTPScript = sFTPScript & "USER " & sUsername & vbCrLf
     sFTPScript = sFTPScript & sPassword & vbCrLf
     sFTPScript = sFTPScript & "cd " & sRemotePath & vbCrLf
     sFTPScript = sFTPScript & "binary" & vbCrLf
     sFTPScript = sFTPScript & "prompt n" & vbCrLf
     sFTPScript = sFTPScript & "put " & sLocalFile & vbCrLf
     sFTPScript = sFTPScript & "quit" & vbCrLf & "quit" & vbCrLf & "quit" & vbCrLf
     
     
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
     " > " & sFTPResults, 0, True
     
     WScript.Sleep 1000
     
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

'------- Get InfraRecorder version Detail------------'
Function InfraRecorder
     'Option Explicit
     ' declare global variables
     Dim aFileFullPath, aDetail
     ' set global variables
     aFileFullPath = "C:\Program Files\InfraRecorder\infrarecorder.exe"
     aDetail = "Product Version"
     ' display a message with file location and file detail
     InfraRecorder = fGetFileDetail(aFileFullPath, aDetail)
     ' make global variable happy. set them free
     Set aFileFullPath = Nothing
     Set aDetail = Nothing
End Function

' get file detail function. created by Stefan Arhip on 20111026 1000
Function fGetFileDetail(aFileFullPath, aDetail)
     ' declare local variables
     Dim pvShell, pvFileSystemObject, pvFolderName, pvFileName, pvFolder, pvFile, i
     ' set object to work with files
     Set pvFileSystemObject = CreateObject("Scripting.FileSystemObject")
     ' check if aFileFullPath provided exists
     If pvFileSystemObject.FileExists(aFileFullPath) Then
          ' extract only folder & file from aFileFullPath
          pvFolderName = pvFileSystemObject.GetFile(aFileFullPath).ParentFolder
          pvFileName = pvFileSystemObject.GetFile(aFileFullPath).Name
          ' set object to work with file details
          Set pvShell = CreateObject("Shell.Application")
          Set pvFolder = pvShell.Namespace(pvFolderName)
          Set pvFile = pvFolder.ParseName(pvFileName)
          ' in case detail is not detected...
          fGetFileDetail = "Detail not detected"
          ' parse 400 details for given file
          For i = 0 To 399
               ' if desired detail name is found, set function result to detail value
               If UCase(pvFolder.GetDetailsOf(pvFolder.Items, i)) = UCase(aDetail) Then
                    fGetFileDetail = pvFolder.GetDetailsOf(pvFile, i)
               End If
          Next
          ' if aFileFullPath provided do not exists
     Else
          fGetFileDetail = "File not found"
     End If
     ' make local variable happy. set them free
     Set pvShell = Nothing
     Set pvFileSystemObject = Nothing
     Set pvFolderName = Nothing
     Set pvFileName = Nothing
     Set pvFolder = Nothing
     Set pvFile = Nothing
     Set i = Nothing
End Function