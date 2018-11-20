Set objShell = CreateObject("WScript.Shell")

strScriptFile = Wscript.ScriptFullName 													' D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO\parse_neo.vbs
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile(strScriptFile)
strFolder = objFSO.GetParentFolderName(objFile) 										' D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO
strSaveToDir = "D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO"    								' base directory for all operations
 
neocpLink = "https://minorplanetcenter.net/iau/NEO/neocp.txt" 							' minorplanetcenter URL, shouldnt need to change this
neocpFile = strSaveToDir+"\neocp.txt"								' where to put the downloaded neocp.txt, adjust as required.
neocpOutputFile = strSaveToDir+"\output.txt"						' where to put output of selected NEOCP objects for further parsing

orbLinkBase = "https://cgi.minorplanetcenter.net/cgi-bin/showobsorbs.cgi?Obj="			' base url to get NEOCPNomin orbit elements, shouldnt need to change this
orbSaveFile = "\orbits.txt"

mpcorbSaveFile = strSaveToDir+"\MPCORB.dat"						' the final (almost) MPCORB.dat 
 
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

if objFSO.FileExists(neocpOutputFile) then
	objFSO.DeleteFile neocpOutputFile
end if

objShell.Run Quotes(strFolder & "\wget.exe") & " " & Quotes(neocpLink) & " -N",1,True 		'download current neocp.txt from MPC 
 
Set objFileRead = objFSO.OpenTextFile(neocpFile, 1) 	' change path for input file from wget 

Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(neocpOutputFile,8,true)  ' create temporary output 
Set orbFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(mpcorbSaveFile,8,true)  ' MPCORB.dat output

Do Until objFileRead.AtEndOfStream
    strLine = objFileRead.ReadLine						' its probably a good idea NOT to touch the positions as they are fixed position.
	object = Mid(strLine, 1,7)							' temporary object designation
	score = Mid(strLine, 9,3)							' neocp desirablility score from 0 to 100, 100 being most desirable.
	dec = Mid(strLine, 35,7)							' declination 
	vmag = Mid(strLine, 44,4)							' if you dont know what this is, change hobbies
	obs = Mid(strLine, 79,4)							' how many observations has it had
	seen = Mid(strLine, 96,7)							' when was the object last seen
	
    if (CSng(score) >= 80) AND (CSng(dec) >= 0) AND (CSng(vmag) <= 19.6) AND (CSng(obs) >= 4) AND (CSng(seen) <= .8) Then
	'if (CSng(score) >= 100)  Then 											' for testing, comment out the above line and uncomment this one to get more objects.
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
					orbFileToWrite.WriteLine(strSearchString+"           "+object)				'write elemets to MPCORB.dat
				Next
			End If
			
		Loop

		objFile.Close
		
	End If	
Loop

objFileRead.Close
objFileToWrite.Close
Set objFileToWrite = Nothing