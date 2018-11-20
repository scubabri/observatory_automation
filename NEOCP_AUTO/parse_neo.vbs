strScriptFile = Wscript.ScriptFullName ' D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO\parse_neo.vbs
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile(strScriptFile)
strFolder = objFSO.GetParentFolderName(objFile) ' C:

Set objShell = CreateObject("WScript.Shell")
 
strLink = "https://minorplanetcenter.net/iau/NEO/neocp.txt"
' Use strFolder to save on the same location of this script.
strSaveTo = "D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO"
 
' WGet saves file always on the actual folder. So, change the actual folder for C:\, where we want to save file
objShell.CurrentDirectory = strSaveTo
 
objShell.Run Quotes(strFolder & "\wget.exe") & " " & Quotes(strLink) & " -N",1,True

objShell.CurrentDirectory = strFolder
 
' Add Quotes to string
' http://stackoverflow.com/questions/2942554/vbscript-adding-quotes-to-a-string
Function Quotes(strQuotes)
	Quotes = chr(34) & strQuotes & chr(34)
End Function


Const ForReading = 1
Const ForAppending = 8
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFileRead = objFSO.OpenTextFile("D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO\neocp.txt", ForReading)
Set objFileWrite = objFSO.CreateTextFile("D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO\output.txt")
objFileWrite.Close
Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile("D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO\output.txt",8,true)

Do Until objFileRead.AtEndOfStream

    strLine = objFileRead.ReadLine
    
	object = Mid(strLine, 1,7)
	score = Mid(strLine, 9,3)
	vmag = Mid(strLine, 44,4)
	obs = Mid(strLine, 79,4)
	seen = Mid(strLine, 96,7)
	
    if (CSng(score) > 80) AND (CSng(vmag) < 19.6) AND (CSng(obs) > 4) AND (CSng(seen) < .8) Then
	msgbox strLine
	objFileToWrite.WriteLine(strLine)
	
	End If	
Loop

objFileRead.Close
objFileToWrite.Close
Set objFileToWrite = Nothing
