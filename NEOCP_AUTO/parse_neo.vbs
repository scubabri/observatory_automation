Set objShell = CreateObject("WScript.Shell")

strScriptFile = Wscript.ScriptFullName 													' D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO\parse_neo.vbs
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile(strScriptFile)
strFolder = objFSO.GetParentFolderName(objFile) 										' D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO
 
neocpLink = "https://minorplanetcenter.net/iau/NEO/neocp.txt" 							' minorplanetcenter URL, shouldnt need to change this
neocpFile = "D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO\neocp.txt"								' where to put the downloaded neocp.txt, adjust as required.

orbLinkBase = "https://cgi.minorplanetcenter.net/cgi-bin/showobsorbs.cgi?Obj="			' base url to get NEOCPNomin orbit elements, shouldnt need to change this
orbSaveFile = "\orbits.txt"
neocpOutputFile = "D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO\output.txt"						' where to put output of selected NEOCP objects for further parsing

mpcorbSaveFile = "D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO\MPCORB.dat"						' the final (almost) MPCORB.dat 
strSaveToDir = "D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO"    								' base directory for all operations
 
' WGet saves file always on the actual folder. So, change the actual folder for C:\, where we want to save file
objShell.CurrentDirectory = strSaveToDir

' Add Quotes to string
' http://stackoverflow.com/questions/2942554/vbscript-adding-quotes-to-a-string
Function Quotes(strQuotes)
	Quotes = chr(34) & strQuotes & chr(34)
End Function
 
if objFSO.FileExists(mpcorbSaveFile) then
	objFSO.DeleteFile mpcorbSaveFile
end if

if objFSO.FileExists(strFolder+"\output.txt") then
	objFSO.DeleteFile strFolder+"\output.txt"
end if

objShell.Run Quotes(strFolder & "\wget.exe") & " " & Quotes(neocpLink) & " -N",1,True 		'download current neocp.txt from MPC 
 
Set objFileRead = objFSO.OpenTextFile(neocpFile, 1) 	' change path for input file from wget 

Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(neocpOutputFile,8,true)  ' create temporary output 
Set orbFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(mpcorbSaveFile,8,true)  ' MPCORB.dat output

Do Until objFileRead.AtEndOfStream
    strLine = objFileRead.ReadLine
	object = Mid(strLine, 1,7)
	score = Mid(strLine, 9,3)
	dec = Mid(strLine, 35,7)
	vmag = Mid(strLine, 44,4)
	obs = Mid(strLine, 79,4)
	seen = Mid(strLine, 96,7)
	
    'if (CSng(score) >= 80) AND (CSng(dec) >= 0) AND (CSng(vmag) <= 19.6) AND (CSng(obs) >= 4) AND (CSng(seen) <= .8) Then
	if (CSng(score) >= 100)  Then
		'msgbox strLine														' output selected neocp for testing only
		objFileToWrite.WriteLine(strLine)									' append neocp object to output.txt
		
		objShell.Run Quotes(strFolder & "\wget.exe") & " " & Quotes(orbLinkBase) & object & "&orb=y -O" & " " & Quotes(strSaveToDir) & orbSaveFile,1,True ' run wget to get orbits from NEOCP 
		
		Set objRegEx = CreateObject("VBScript.RegExp")
		objRegEx.Pattern = "NEOCPNomin"
        Set objFile = objFSO.OpenTextFile(strSaveToDir+orbSaveFile, 1)
		Do Until objFile.AtEndOfStream
			strSearchString = objFile.ReadLine
			Set colMatches = objRegEx.Execute(strSearchString)
			
			If colMatches.Count > 0 Then
				For Each strMatch in colMatches
					'Wscript.Echo strSearchString							'echo selected MPCORB element for testing only
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
