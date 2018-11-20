Set objShell = CreateObject("WScript.Shell")
strScriptFile = Wscript.ScriptFullName ' D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO\parse_neo.vbs
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile(strScriptFile)
strFolder = objFSO.GetParentFolderName(objFile) ' D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO
 
strLink = "https://minorplanetcenter.net/iau/NEO/neocp.txt" 
orbLinkBase = "https://cgi.minorplanetcenter.net/cgi-bin/showobsorbs.cgi?Obj="
orbSaveFile = "\orbits.txt"
mpcorbSaveFile = "D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO\MPCORB.dat"
strSaveTo = "D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO"    ' Use strFolder to save on the same location of this script.
 
' WGet saves file always on the actual folder. So, change the actual folder for C:\, where we want to save file
objShell.CurrentDirectory = strSaveTo
 
objShell.Run Quotes(strFolder & "\wget.exe") & " " & Quotes(strLink) & " -N",1,True 'download current neocp.txt from MPC 

objShell.CurrentDirectory = strFolder
 
' Add Quotes to string
' http://stackoverflow.com/questions/2942554/vbscript-adding-quotes-to-a-string
Function Quotes(strQuotes)
	Quotes = chr(34) & strQuotes & chr(34)
End Function

Set objFileRead = objFSO.OpenTextFile("D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO\neocp.txt", 1) ' change path for input file from wget 
'Set objFileWrite = objFSO.CreateTextFile("D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO\output.txt")  ' change path for output directory
if objFSO.FileExists(mpcorbSaveFile) then
            objFSO.DeleteFile mpcorbSaveFile
        end if
if objFSO.FileExists(strFolder+"\output.txt") then
            objFSO.DeleteFile strFolder+"\output.txt"
        end if

Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile("D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO\output.txt",8,true)  ' create temporary output 
Set orbFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile("D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO\MPCORB.dat",8,true)  ' MPCORB.dat output


Do Until objFileRead.AtEndOfStream

    strLine = objFileRead.ReadLine
    
	object = Mid(strLine, 1,7)
	score = Mid(strLine, 9,3)
	dec = Mid(strLine, 35,7)
	vmag = Mid(strLine, 44,4)
	obs = Mid(strLine, 79,4)
	seen = Mid(strLine, 96,7)
	
    if (CSng(score) >= 80) AND (CSng(dec) >= 0) AND (CSng(vmag) <= 19.6) AND (CSng(obs) >= 4) AND (CSng(seen) <= .8) Then
	'if (CSng(score) >= 80)  Then
		msgbox strLine
		objFileToWrite.WriteLine(strLine)
		objShell.Run Quotes(strFolder & "\wget.exe") & " " & Quotes(orbLinkBase) & object & "&orb=y -O" & " " & Quotes(strSaveTo) & orbSaveFile,1,True ' run wget to get orbits from NEOCP 
		
		Set objRegEx = CreateObject("VBScript.RegExp")
		objRegEx.Pattern = "NEOCPNomin"
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objFile = objFSO.OpenTextFile(strSaveTo+orbSaveFile, 1)
		Do Until objFile.AtEndOfStream
			strSearchString = objFile.ReadLine
			Set colMatches = objRegEx.Execute(strSearchString)
			If colMatches.Count > 0 Then
				For Each strMatch in colMatches
					Wscript.Echo strSearchString
					orbFileToWrite.WriteLine(strSearchString)
				Next
			End If
		Loop

		objFile.Close
		
	End If	
Loop

objFileRead.Close
objFileToWrite.Close
Set objFileToWrite = Nothing
