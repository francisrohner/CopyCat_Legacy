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
currentLine = configFileReader.ReadLine()
copyLocation = Replace(Trim(Mid(currentLine, InStr(currentLine, "=") + 1)), "\", "\\")
copyLocation = resolvePaths(copyLocation)
logWriter.WriteLine("Copy Location Set To: " & copyLocation)

Redim paths(numPaths)
Redim files(numFiles)

do while not configFileReader.AtEndOfStream

	currentLine = configFileReader.ReadLine()
	currentLine = Trim(Mid(currentLine, InStr(currentLine, "=") + 1))
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
environmentVariables = Array("%USERPROFILE%", "%SYSTEMROOT%", "%SYSTEMDRIVE%", "%PROGRAMFILES%", "%PROGRAMFILES(x86)%", "%PROGRAMDATA%", "%APPDATA%")
'Replace case in-sensitive
for i = 0 to UBound(environmentVariables)
    value = Replace(value, environmentVariables(i), oShell.ExpandEnvironmentStrings(environmentVariables(i)), 1, -1, vbTextCompare)
next
resolvePaths = value
end function
