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

Function Quotes(strQuotes)																' Add Quotes to string
	Quotes = chr(34) & strQuotes & chr(34)												' http://stackoverflow.com/questions/2942554/vbscript-adding-quotes-to-a-string
End Function

Include "VbsJson"																		' include VBScript Jason parse funtions

strScriptFile = Wscript.ScriptFullName 													' D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO\parse_neo.vbs
Set objFSO = CreateObject("Scripting.FileSystemObject")
objFile = objFSO.GetFile(strScriptFile)
strFolder = objFSO.GetParentFolderName(objFile) 										' D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO
baseDir = "D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO\VBS_JSON"    								' base directory for all operations
 
scoutLink = "https://ssd-api.jpl.nasa.gov/scout.api"
scoutSaveFile = "\scout.json"
neocpOutputFile = baseDir+"\output.txt"						' where to put output of selected NEOCP objects for further parsing

orbLinkBase = "https://cgi.minorplanetcenter.net/cgi-bin/showobsorbs.cgi?Obj="			' base url to get NEOCPNomin orbit elements, shouldnt need to change this
orbSaveFile = "\orbits.txt"

mpcorbSaveFile = baseDir+"\MPCORB.dat"						' the final (almost) MPCORB.dat 
 
' WGet saves file always on the actual folder. So, change the actual folder for C:\, where we want to save file
objShell.CurrentDirectory = baseDir
 
if objFSO.FileExists(mpcorbSaveFile) then
	objFSO.DeleteFile mpcorbSaveFile
end if

if objFSO.FileExists(neocpOutputFile) then
	objFSO.DeleteFile neocpOutputFile
end if

objShell.Run Quotes(strFolder & "\wget.exe") & " " & Quotes(scoutLink) & " -O" & " " & Quotes(baseDir) & scoutSaveFile,1,True ' Get NEOCP from Scout

Dim json, neocpStr, jsonDecoded
Set json = New VbsJson
neocpStr = objFSO.OpenTextFile(baseDir+scoutSaveFile).ReadAll
Set jsonDecoded = json.Decode(neocpStr)
objCount = jsonDecoded("count")
'Wscript.Echo objCount

Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(neocpOutputFile,8,true)  ' create temporary output 
Set orbFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(mpcorbSaveFile,8,true)  ' MPCORB.dat output

Dim counter
counter = 0 

Do Until counter = (objCount-1)

	object  = jsonDecoded("data")(counter)("objectName")					' temporary object designation
	score   = jsonDecoded("data")(counter)("neoScore")					    ' neocp desirablility score from 0 to 100, 100 being most desirable.
	dec     = jsonDecoded("data")(counter)("dec")							' declination 
	ra      = jsonDecoded("data")(counter)("ra")
	vmag    = jsonDecoded("data")(counter)("Vmag")							' if you dont know what this is, change hobbies
	obs     = jsonDecoded("data")(counter)("nObs")							' how many observations has it had
	lastRun = jsonDecoded("data")(counter)("lastRun")						' when was the object last lastRun
	rate    = jsonDecoded("data")(counter)("rate")
		
    if (CSng(score) >= 80) AND (CSng(dec) >= 0) AND (CSng(vmag) <= 20) AND (CSng(obs) > 3)AND DateDiff("h",lastRun,FormatDateTime(Now)) <= 11 Then
	'if (CSng(score) >= 100)  Then 											' for testing, comment out the above line and uncomment this one to get more objects.
																
		objFileToWrite.WriteLine("Object    score    ra   dec    vmag      nobs      lastRun              rate")
		objFileToWrite.WriteLine("-------------------------------------------------------------------------")
		objFileToWrite.WriteLine(object+ "     " + score + "  " + ra + "  " + dec + "    " + vmag + "      " + obs + "     " + lastRun + "      " + rate)									' append neocp object to output.txt
		
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
					orbFileToWrite.WriteLine(strSearchString+"           "+object)				'write elemets to MPCORB.dat
				Next
			End If
			
		Loop

		objFile.Close
		

		End If	
	 counter=(counter+1)
Loop
if objFSO.FileExists(baseDir+orbSaveFile) then
	objFSO.DeleteFile basedir+orbSaveFile
end if
if objFSO.FileExists(baseDir+scoutSaveFile) then
	objFSO.DeleteFile baseDir+scoutSaveFile
end if
objFileToWrite.Close
Set objFileToWrite = Nothing