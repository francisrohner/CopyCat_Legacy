'TODO Resolve %USERPROFILE%???

'Grab Current Directory before elevation
Dim WshShell, strCurDir
Set WshShell = CreateObject("WScript.Shell")
strCurDir = WshShell.CurrentDirectory

'Run as Admin
If Not WScript.Arguments.Named.Exists("elevate") Then
  CreateObject("Shell.Application").ShellExecute WScript.FullName _
    , WScript.ScriptFullName & " /elevate /currentDirectory " & strCurDir, "", "runas", 1
  WScript.Quit
End If

If WScript.Arguments.Named.Exists("currentDirectory")  Then
	strCurDir = Trim(WScript.Arguments(2))
End If

'Declarations
Dim copyLocation
Dim currentLine
Dim currentPath
Set logWriter = CreateObject("Scripting.FileSystemObject").OpenTextFile(strCurDir & "\\CopyCatLog.txt",2,true)
Set configFileReader = CreateObject("Scripting.FileSystemObject").OpenTextFile(strCurDir & "\\CopyCat.cfg",1)
set filesys=CreateObject("Scripting.FileSystemObject")

'Initializations
numLines = 0

currentLine = configFileReader.ReadLine()
copyLocation = Replace(Trim(Mid(currentLine, InStr(currentLine, "=") + 1)), "\", "\\")
copyLocation = resolvePaths(copyLocation)
logWriter.WriteLine("Copy Location Set To: " & copyLocation)

Redim paths(numPaths)
Redim files(numFiles)

do while not configFileReader.AtEndOfStream

	currentLine = configFileReader.ReadLine()
	currentLine = Trim(Mid(currentLine, InStr(currentLine, "=") + 1))
	'currentLine = Replace(currentLine, "\", "\\")
    currentLine = resolvePaths(currentLine)

	if InStr(currentLine, "#") Then
		logWriter.WriteLine("Not copying current path, it's probably a comment")
	elseIf filesys.FolderExists(currentLine) Then
		filesys.CopyFolder currentLine, copyLocation
		logWriter.WriteLine("Successfully copied path " & currentLine)
	elseif filesys.FileExists(currentLine) Then
		filesys.CopyFile currentLine, copyLocation, true
		logWriter.WriteLine("Successfully copied file " & currentLine)
	Else
		logWriter.WriteLine("Failed copying " & currentLine & " path doesn't exist")
	End If

loop

'Disposal
configFileReader.Close
Set configFileReader = Nothing
logWriter.Close
Set logWriter = Nothing

MsgBox("Meow, everything copied")

function resolvePaths(value)
Set oShell = CreateObject("Wscript.Shell")
strUserProfile = oShell.ExpandEnvironmentStrings("%USERPROFILE%")
strWinDir = oShell.ExpandEnvironmentStrings("%SYSTEMROOT%")
strSysDrive = oShell.ExpandEnvironmentStrings("%SYSTEMDRIVE%")
strProgramFiles = oShell.ExpandEnvironmentStrings("%PROGRAMFILES%")
strProgramFiles86 = oShell.ExpandEnvironmentStrings("%PROGRAMFILES(x86)%")
strProgramData = oShell.ExpandEnvironmentStrings("%PROGRAMDATA%")

'Replace case in-sensitive
value = Replace(value, "%USERPROFILE%", strUserProfile, 1, -1, vbTextCompare)
value = Replace(value, "%SYSTEMDRIVE%", strSysDrive, 1, -1, vbTextCompare)
value = Replace(value, "%PROGRAMFILES%", strProgramFiles, 1, -1, vbTextCompare)
value = Replace(value, "%PROGRAMFILES(x86)%", strProgramFiles86, 1, -1, vbTextCompare)
value = Replace(value, "%PROGRAMDATA%", strProgramData, 1, -1, vbTextCompare)
resolvePaths = Replace(value, "%SYSTEMROOT%", strWinDir, 1, -1, vbTextCompare)

end function
