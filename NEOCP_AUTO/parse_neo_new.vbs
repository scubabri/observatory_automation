Set objShell = CreateObject("WScript.Shell")

Sub Include(file)
	On Error Resume Next

	Dim FSO
	Set FSO = CreateObject("Scripting.FileSystemObject")
	ExecuteGlobal FSO.OpenTextFile(file & ".vbs", 1).ReadAll()
	Set FSO = Nothing

	If Err.Number <> 0 Then
		If Err.Number = 1041 Then
			Err.Clear
		Else
			WScript.Quit 1
		End If
	End If
End Sub

Include "VbsJson"	

Dim json, neocpStr, jsonDecoded
Set json = New VbsJson

strScriptFile = Wscript.ScriptFullName 													' D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO\parse_neo.vbs
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile(strScriptFile)
strFolder = objFSO.GetParentFolderName(objFile) 										' D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO
baseDir = "D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO"    								' base directory for all operations

scoutLink = "https://ssd-api.jpl.nasa.gov/scout.api?tdes="
scoutSaveFile = "\scout.json"
 
neocpLink = "https://minorplanetcenter.net/iau/NEO/neocp.txt" 							' minorplanetcenter URL, shouldnt need to change this
neocpFile = baseDir+"\neocp.txt"								' where to put the downloaded neocp.txt, adjust as required.
objectsSaveFile = baseDir+"\output.txt"						' where to put output of selected NEOCP objects for further parsing

orbLinkBase = "https://cgi.minorplanetcenter.net/cgi-bin/showobsorbs.cgi?Obj="			' base url to get NEOCPNomin orbit elements, shouldnt need to change this
orbSaveFile = "\orbits.txt"

mpcorbSaveFile = baseDir+"\MPCORB.dat"						' the final (almost) MPCORB.dat 
 
' WGet saves file always on the actual folder. So, change the actual folder for C:\, where we want to save file
objShell.CurrentDirectory = baseDir

' Add Quotes to string
' http://stackoverflow.com/questions/2942554/vbscript-adding-quotes-to-a-string
Function Quotes(strQuotes)
	Quotes = chr(34) & strQuotes & chr(34)
End Function
 
if objFSO.FileExists(mpcorbSaveFile) then
	objFSO.DeleteFile mpcorbSaveFile
end if

if objFSO.FileExists(objectsSaveFile) then
	objFSO.DeleteFile objectsSaveFile
end if

objShell.Run Quotes(strFolder & "\wget.exe") & " " & Quotes(neocpLink) & " -N",1,True 		'download current neocp.txt from MPC 
 
Set neocpFileRead = objFSO.OpenTextFile(neocpFile, 1) 	' change path for input file from wget 

Set objectsFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(objectsSaveFile,8,true)  ' create output.txt
Set MPCorbFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(mpcorbSaveFile,8,true)  ' MPCORB.dat output

objectsFileToWrite.WriteLine("Object    score    ra   dec    vmag      nobs      lastSeen       rate")
objectsFileToWrite.WriteLine("-------------------------------------------------------------------------")

Do Until neocpFileRead.AtEndOfStream
    strLine = neocpFileRead.ReadLine						' its probably a good idea NOT to touch the positions as they are fixed position.
	object = Mid(strLine, 1,7)							' temporary object designation
	score = Mid(strLine, 9,3)							' neocp desirablility score from 0 to 100, 100 being most desirable.
	dec = Mid(strLine, 35,7)							' declination 
	vmag = Mid(strLine, 44,4)							' if you dont know what this is, change hobbies
	obs = Mid(strLine, 79,4)							' how many observations has it had
	seen = Mid(strLine, 96,7)							' when was the object last seen
	
    if (CSng(score) >= 0) AND (CSng(dec) >= 0) AND (CSng(vmag) <= 22) AND (CSng(obs) >= 3) AND (CSng(seen) <= 2) Then
	'if (CSng(score) >= 1)  Then 											' for testing, comment out the above line and uncomment this one to get more objects.
		'msgbox strLine														' output selected neocp for testing only
		'objectsFileToWrite.WriteLine(strLine)									' append neocp object to output.txt
		
		objShell.Run Quotes(strFolder & "\wget.exe") & " " & Quotes(orbLinkBase) & object & "&orb=y -O" & " " & Quotes(baseDir) & orbSaveFile,1,True ' run wget to get orbits from NEOCP
		
		Set objRegEx = CreateObject("VBScript.RegExp")
		objRegEx.Pattern = "NEOCPNomin"
        Set objFile = objFSO.OpenTextFile(baseDir+orbSaveFile, 1)
		Do Until objFile.AtEndOfStream
			strSearchString = objFile.ReadLine
			Set colMatches = objRegEx.Execute(strSearchString)
			
			If colMatches.Count > 0 Then
				For Each strMatch in colMatches
					'Wscript.Echo strSearchString							'echo selected MPCORB element for testing only
					MPCorbFileToWrite.WriteLine(strSearchString+"           "+object)				'write elemets to MPCORB.dat
				Next
			End If
			
		Loop
		objFile.Close
		objShell.Run Quotes(strFolder & "\wget.exe") & " " & Quotes(scoutLink) & object & " -O" & " " & Quotes(baseDir) & scoutSaveFile,1,True ' Get NEOCP from Scout
		scoutStr = objFSO.OpenTextFile(baseDir+scoutSaveFile).ReadAll
		Set jsonDecoded = json.Decode(scoutStr)
	   
		scoutobject  = jsonDecoded("objectName")					' temporary object designation
		scoutscore   = jsonDecoded("neoScore")					    ' neocp desirablility score from 0 to 100, 100 being most desirable.
		scoutdec     = jsonDecoded("dec")							' declination 
		scoutra      = jsonDecoded("ra")
		scoutvmag    = jsonDecoded("Vmag")							' if you dont know what this is, change hobbies
		scoutobs     = jsonDecoded("nObs")							' how many observations has it had
		scoutrate    = jsonDecoded("rate")

		objectsFileToWrite.WriteLine(scoutobject+ "     " + score + "  " + scoutra + "  " + scoutdec + "    " + scoutvmag + "      " + scoutobs + "     " + seen + "      " + scoutrate)		
		
	End If	
Loop

neocpFileRead.Close
objectsFileToWrite.Close
MPCorbFileToWrite.Close

if objFSO.FileExists(baseDir+orbSaveFile) then
	objFSO.DeleteFile basedir+orbSaveFile
end if
if objFSO.FileExists(neocpFile) then
	objFSO.DeleteFile neocpFile
end if
if objFSO.FileExists(baseDir+scoutSaveFile) then
	objFSO.DeleteFile baseDir+scoutSaveFile
end if
Set objectsFileToWrite = Nothing