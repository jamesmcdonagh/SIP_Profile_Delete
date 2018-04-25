'==========================================================================
'
'NAME:  Delete sip profile
'
'COMMENT: This script deletes the SIP profile for the user from machine if it exists 
'the purpose here is delete the cache files used by Lync 2013 and Skype for Business 2015 2016 clients
'==========================================================================

Option Explicit

Dim objShell12
Dim objUserEnv
Dim strUserPro
Dim userProfile,SipProfile
Dim proPath
Dim objFSO
Dim objStartFolder
Dim objFolder
Dim colFiles
Dim objFile
Dim Subfolder
Dim uProfile
Dim WshShell
Dim objWMIService
Dim colProcessList
Dim colProcessList2
Dim objProcess
Dim complete

'Terminate existing Lync / Skype for Business processes

Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colProcessList = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'lync.exe'")

For Each objProcess in colProcessList
	objProcess.Terminate()
Next

Set colProcessList2 = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'outlook.exe'")

For Each objProcess in colProcessList2
	objProcess.Terminate()
Next


WScript.Sleep 1000

Set objShell12=CreateObject("WScript.Shell")
Set objUserEnv=objShell12.Environment("User")

strUserPro= objShell12.ExpandEnvironmentStrings(objUserEnv("TEMP"))
userProfile = objShell12.ExpandEnvironmentStrings("%userprofile%")

'Delete SIP profile

DeleteSip strUserPro

Set objFSO = CreateObject("Scripting.FileSystemObject")

SipProfile=userProfile & "\AppData\Local\Microsoft\Office\15.0\Lync"

If (objFSO.FolderExists(SipProfile)) then
	uProfile=userProfile & "\AppData\Local\Microsoft\Office\15.0"
	LyncClear(uProfile)
End If

SipProfile = userProfile & "\AppData\Local\Microsoft\Office\16.0\Lync"

If (objFSO.FolderExists(SipProfile)) then
	uProfile=userProfile & "\AppData\Local\Microsoft\Office\16.0"
	LyncClear(uProfile)
End If

complete=MsgBox("Process complete. Please relaunch Outlook and Skype for Business.",0,"CallTower SIP Delete Profile Tool")

WScript.Quit

Sub LyncClear(objStartFolder)
	ShowSubfolders objFSO.GetFolder(objStartFolder)
	DeleteSip SipProfile
	SipProfile=SipProfile & "Sip_*"
	DeleteSip SipProfile
End Sub

Sub ShowSubFolders(Folder)
	For Each Subfolder in Folder.SubFolders
		proPath = Right(Subfolder.Path,4)
		If proPath = "Lync" Then
			DeleteSip SipProfile
		End If
    Next
End Sub

Sub DeleteSip (strSipPath)

	On Error Resume Next

	Dim objFSO
	Dim objFolder,objDir
	Dim i

	Set objFSO=CreateObject("Scripting.FileSystemObject")
	Set objFolder=objFSO.GetFolder(strSipPath)

	'delete folder

	For i=0 To 10
		For Each objDir In objFolder.SubFolders
			objDir.Delete True
		Next
	Next

	Set objFSO=Nothing
	Set objFolder=Nothing
	Set objDir=Nothing

End Sub

'==================================================================