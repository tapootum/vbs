'Dim fillAttributes(300)

'Set shell = CreateObject("Shell.Application")
'Set folder = shell.Namespace("C:\Windows")

'Set file = folder.ParseName("notepad.exe")

'For i = 0 to 299
    '    Wscript.Echo i & vbtab & fillAttributes(i) _
    '    & ": " & folder.GetDetailsOf(file, i) 
'Next

'WScript.Echo GetProductVersion("c:\Windows","notepad.exe")

' must explicitly declare all variables
Option Explicit
' declare global variables
Dim aFileFullPath, aDetail
' set global variables
aFileFullPath = "C:\Program Files\InfraRecorder\infrarecorder.exe"
aDetail = "Product Version"
' display a message with file location and file detail
WScript.Echo ("File location: " & vbTab & aFileFullPath & vbNewLine & _
aDetail & ": " & vbTab & fGetFileDetail(aFileFullPath, aDetail))
' make global variable happy. set them free
Set aFileFullPath = Nothing
Set aDetail = Nothing
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
If uCase(pvFolder.GetDetailsOf(pvFolder.Items, i)) = uCase(aDetail) Then
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


Function GetProductVersion (sFilePath, sProgram)
Dim objShell, objFolder, objFolderItem, i 
If FSO.FileExists(sFilePath & "\" & sProgram) Then
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.Namespace(sFilePath)
    Set objFolderItem = objFolder.ParseName(sProgram)
    Dim arrHeaders(300)
    For i = 0 To 300
        arrHeaders(i) = objFolder.GetDetailsOf(objFolder.Items, i)
        'WScript.Echo i &"- " & arrHeaders(i) & ": " & objFolder.GetDetailsOf(objFolderItem, i)
        If lcase(arrHeaders(i))= "product version" Then
            GetProductVersion= objFolder.GetDetailsOf(objFolderItem, i)
            Exit For
        End If
    Next
End If
End Function