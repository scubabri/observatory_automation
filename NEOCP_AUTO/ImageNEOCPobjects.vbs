
Option Explicit

'Global Objects
Dim objTheSkyChart
Dim objTheSkyInfo
Dim objTheSky
Dim objTele
Dim objCam

'Global User Variables see InitGlobalUserVariables()
Dim PathToTargetsFile
Dim expTime
Dim bIgnoreErrors
Dim tgtname
Dim status
Dim imagecount
Dim imagetaken
Dim imageScale
Dim objectMotion

'This is where the work starts
Call InitGlobalUserVariables()
Call CreateObjects()
Call ConnectObjects()
'Call ParkScope()
Call TargetLoop()
Call DisconnectObjects()
Call DeleteObjects()

Sub InitGlobalUserVariables()

	PathToTargetsFile = "C:\Users\brians\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO\output.txt"
	
	imageScale = 1.95
	
	'If you want your script to run all night regardless of errors, Set bIgnoreErrors = True
	bIgnoreErrors = False

End Sub

Sub GetExposureData()
	
	expTime = (60*(imageScale/objectMotion))
	
	If (expTime >= 30) AND (expTime < 45) Then
		expTime = 30 
		'imagecount = 120
	ElseIf	(expTime >= 45) AND (expTime < 60) Then 
		expTime = 45 
		'imagecount = 90
	ElseIf expTime >= 60 Then 
		expTime = 60 
		'imagecount = 60
	End If
	
	imagecount = round((60*(60/expTime)),0)

End Sub

Sub GetDataFromTextFile(LineFromFile, szTargetName, vMag, objectMotion)

	szTargetName= Mid(LineFromFile,1,10)
	tgtname = "MPL " + szTargetName 
	vMag = Mid(LineFromFile,32,4)
	objectMotion = Mid(LineFromFile,59,5)

End Sub


Sub GetUpdatedCoordinates(tgtname, dRa, dDec)

	Set objTheSkyChart = CreateObject("TheSkyX.sky6StarChart") 
	status = objTheSkyChart.Find (tgtname)
	Set objTheSkyInfo = CreateObject("TheSkyX.sky6ObjectInformation") 
	objTheSkyInfo.Index = 0 
	status = objTheSkyInfo.Property (54) 
	dRa = objTheSkyInfo.ObjInfoPropOut 
	status = objTheSkyInfo.Property (55)
	dDec = objTheSkyInfo.ObjInfoPropOut

End Sub

Sub PromptOnError(bErrorOccurred)
	Dim bExitScript
	
	bErrorOccurred = False
	bExitScript = False

	if (bIgnoreErrors = True) then 
		'Ignore all errors except when the user Aborts
		if (CStr(Hex(Err.Number)) = "800404BC") then 
			'Do nothing and let the user abort
		else
			Err.Clear
		end if
	end if

	if (Err.Number) then 
		bErrorOccurred = True
		bExitScript = MsgBox ("An error has occurred running this script.  Error # " & CStr(Hex(Err.Number)) & " " & Err.Description + CRLF + CRLF + "Exit Script?", vbYesNo + vbInformation)
	end if

	If bExitScript = vbYes Then
		WScript.Quit
	End if 
End Sub

Sub TargetLoop()
	On Error Resume Next
	Dim TxtFile
	Dim fso
	Const ForReading = 1
	Dim dRa
	Dim dDec
	Dim bErrorOccurred
	Dim szTargetName

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set TxtFile = fso.OpenTextFile(PathToTargetsFile, ForReading)

	Do While (TxtFile.AtEndOfStream <> True)
	    	
	    imagetaken = 0
		Call GetDataFromTextFile(TxtFile.ReadLine, szTargetName, vMag, objectMotion)

		Call GetExposureData()
		
		msgbox (tgtname & " " & vMag & " " & objectMotion & " " & imagecount & " " & expTime)

		Do While (imagetaken <= imagecount)
			Call GetUpdatedCoordinates(tgtname, dRa, dDec)

			Err.Clear 'Clear the error object
			bErrorOccurred = FALSE 'No error has occurred
	
			if (bErrorOccurred = False) then
				Call objTele.SlewToRaDec(dRa, dDec, tgtname)
				'Call PromptOnError(bErrorOccurred)
			end if
		
			if (bErrorOccurred = False) then
				objCam.ExposureTime = expTime
				objCam.Frame = 1 
				objCam.ImageReduction = 0
				objCam.AutoSaveOn = 1 
				status = objCam.TakeImage()
				Call PromptOnError(bErrorOccurred)
				imagetaken = imagetaken + 1
			end if
		Loop
	Loop
End Sub

Sub CreateObjects()
	Set objTheSky 	= CreateObject("TheSkyX.Application") 
	Set objTele		= CreateObject("TheSkyX.sky6RASCOMTele") 
	Set objCam 		= CreateObject("TheSkyX.ccdsoftCamera") 
End Sub

Sub ConnectObjects()
	objTele.Connect()
	objCam.Connect()
End Sub

Sub ParkScope()
	status = objTele.Park()
End Sub

Sub DisconnectObjects()
	objTele.Disconnect()
	'objCam.Disconnect()
End Sub 

Sub DeleteObjects()
	Set objTheSky = Nothing
	Set objTele = Nothing
	Set objCam = Nothing
End Sub 