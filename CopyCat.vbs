'Author: Francis Rohner @ http://francisrohner.com

'Grab Current Directory before elevation
Dim WshShell, strCurDir
Set WshShell = CreateObject("WScript.Shell")
strCurDir = WshShell.CurrentDirectory

'Run as Admin
If Not WScript.Arguments.Named.Exists("elevate") Then
	MsgBox(Wscript.FullName)
	MsgBox(WScript.ScriptFullName)
	'MsgBox(strCurDir)
  CreateObject("Shell.Application").ShellExecute WScript.FullName _
    ,"""" & WScript.ScriptFullName & """ /elevate /currentDirectory """ & strCurDir & """", "", "runas", 1
  WScript.Quit
End If
If WScript.Arguments.Named.Exists("currentDirectory")  Then
	strCurDir = Trim(WScript.Arguments(2))
End If

'Declarations
Dim copyLocation
Dim currentLine
Dim currentPath
Dim isComment
Set filesys = CreateObject("Scripting.FileSystemObject")
Set logWriter = filesys.OpenTextFile(strCurDir & "\\CopyCatLog.txt",2,true)
Set configFileReader = filesys.OpenTextFile(strCurDir & "\\CopyCat.cfg",1)

'Initializations
currentLine = configFileReader.ReadLine()
If InStr(currentLine, "here") or InStr(currentLine, "Here") Then
    copyLocation = strCurDir
Else
    copyLocation = Replace(Trim(Mid(currentLine, InStr(currentLine, "=") + 1)), "\", "\\")
    copyLocation = resolvePaths(copyLocation)
End if
logWriter.WriteLine("Copy Location Set To: " & copyLocation)

Redim paths(numPaths)
Redim files(numFiles)

Do While Not configFileReader.AtEndOfStream

	currentLine = configFileReader.ReadLine()
	currentLine = resolveConfigLine(currentLine)
	
	If InStr(currentLine, "#") Then
		logWriter.WriteLine("Not copying current path, it's probably a comment")
	Elseif filesys.FileExists(currentLine) Then
		filesys.CopyFile currentLine, copyLocation & "\", True
		logWriter.WriteLine("Successfully copied file " & currentLine)
	ElseIf filesys.FolderExists(currentLine) Then
		Dim xcopyCommand
		xcopyCommand = "xcopy.exe " & """" & currentLine & "\*""" & " " & """" & copyLocation & "\" & getEndPath(currentLine) & "\"" /b /s /i /Y"
		logWriter.WriteLine(xcopyCommand)
		WshShell.Run xcopyCommand
		logWriter.WriteLine("Successfully copied path " & currentLine)
	Else
		logWriter.WriteLine("Failed copying " & currentLine & " path doesn't exist")
	End If

Loop

'Disposal
configFileReader.Close
Set configFileReader = Nothing
logWriter.Close
Set logWriter = Nothing
Set filesys = Nothing

'MsgBox("Meow, everything copied")

Function resolveConfigLine(value)
	If InStr(value, "#") = 0 Then
		resolveConfigLine = Trim(Mid(value, InStr(value, "=") + 1))
		resolveConfigLine = resolvePaths(resolveConfigLine)
	Else
		resolveConfigLine = "#"
	End If
End Function

Function resolvePaths(value)
	Set oShell = CreateObject("Wscript.Shell")
	environmentVariables = Array("%USERPROFILE%", "%SYSTEMROOT%", "%SYSTEMDRIVE%", "%PROGRAMFILES%", "%PROGRAMFILES(x86)%", "%PROGRAMDATA%", "%APPDATA%")
	'Replace case in-sensitive
	For i = 0 to UBound(environmentVariables)
		value = Replace(value, environmentVariables(i), oShell.ExpandEnvironmentStrings(environmentVariables(i)), 1, -1, vbTextCompare)
	Next
	resolvePaths = value
End Function

Function getEndPath(value)
	getEndPath = Trim(Mid(value, InStrRev(value, "\") + 1))
End Function

Function quote(value)
	quote = """" & """"
End Function
