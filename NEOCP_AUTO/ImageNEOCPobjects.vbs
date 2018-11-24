
'Option Explicit

'Global Objects
Dim objTheSkyChart
Dim objTheSkyInfo
Dim objTheSky
Dim objTele
Dim objCam

'Global User Variables see InitGlobalUserVariables()
Dim PathToTargetsFile
Dim PathToWeatherFile
Dim expTime
Dim ignoreErrors
Dim targetName
Dim status
Dim imageCount
Dim imageTaken
Dim imageScale
Dim objectMotion
Dim objectAlt
Dim cameraTemp

'This is where the work starts
Call InitGlobalUserVariables()
Call checkWeather()
'Call CreateObjects()
'Call ConnectObjects()
'Call setCamTemp()
'Call UnParkScope
'Call TargetLoop()
'Call ParkScope()
'Call DisconnectObjects()
'Call DeleteObjects()

Sub InitGlobalUserVariables()

	PathToTargetsFile = "C:\Users\brians\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO\output.txt"
	PathToWeatherFile = "C:\Users\brians\Dropbox\ASTRO\weatherdata.txt"
	
	imageScale = 1.95
	cameraTemp = -10
	enableWeather = 1
	
	'If you want your script to run all night regardless of errors, Set ignoreErrors = True
	ignoreErrors = False
	
End Sub

Sub GetExposureData()
	
	expTime = (60*(imageScale/objectMotion))
	
	If (expTime >= 30) AND (expTime < 45) Then
		expTime = 30 
	ElseIf	(expTime >= 45) AND (expTime < 60) Then 
		expTime = 45 
	ElseIf expTime >= 60 Then 
		expTime = 60 
	End If
	
	imageCount = round((60*(60/expTime)),0)

End Sub

Sub GetDataFromTextFile(LineFromFile, szTargetName, vMag, objectMotion)

	szTargetName= Mid(LineFromFile,1,10)
	targetName = "MPL " + szTargetName 
	vMag = Mid(LineFromFile,32,4)
	objectMotion = Mid(LineFromFile,59,5)

End Sub


Sub GetUpdatedCoordinates(targetName, dRa, dDec, objectAlt)

	Set objTheSkyChart = CreateObject("TheSkyX.sky6StarChart") 
	status = objTheSkyChart.Find (targetName)
	Set objTheSkyInfo = CreateObject("TheSkyX.sky6ObjectInformation") 
	objTheSkyInfo.Index = 0 
	status = objTheSkyInfo.Property (54) 
	dRa = objTheSkyInfo.ObjInfoPropOut 
	status = objTheSkyInfo.Property (55)
	dDec = objTheSkyInfo.ObjInfoPropOut
	status = objTheSkyInfo.Property (59)
	objectAlt = objTheSkyInfo.ObjInfoPropOut

End Sub

Sub PromptOnError(bErrorOccurred)
	Dim bExitScript
	
	bErrorOccurred = False
	bExitScript = False

	if (ignoreErrors = True) then 
		'Ignore all errors except when the user Aborts
		if (CStr(Hex(Err.Number)) = "800404BC") then 
			'Do nothing and let the user abort
		else
			Err.Clear
		end if
	end if

	if (Err.Number) then 
		bErrorOccurred = True
		'bExitScript = MsgBox ("An error has occurred running this script.  Error # " & CStr(Hex(Err.Number)) & " " & Err.Description + CRLF + CRLF + "Exit Script?", vbYesNo + vbInformation)
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
	
		Err.Clear 'Clear the error object
		bErrorOccurred = False 'No error has occurred
			
		Call GetDataFromTextFile(TxtFile.ReadLine, szTargetName, vMag, objectMotion)
		Call GetExposureData()
		
		msgbox (targetName & " " & vMag & " " & objectMotion & " " & imageCount & " " & expTime)
		imageTaken = 0
		
		Do While (imageTaken <= imageCount)
			
			Call GetUpdatedCoordinates(targetName, dRa, dDec, objectAlt)

			if (bErrorOccurred = False) then
			
				Call checkObjectElev()
				Call objTele.SlewToRaDec(dRa, dDec, targetName)
				Call PromptOnError(bErrorOccurred)
				
				if (objectAlt < 0) Then
					imageTaken = imageCount
				End If
			
			End If
		
			if (bErrorOccurred = False) then
				Call takeImage()
				msgbox "calling takeImage"
				Call PromptOnError(bErrorOccurred)
				imageTaken = imageTaken + 1
			else 
				imageTaken = imageCount+1
			
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

Sub setCamTemp()

	objCam.Connect()
	objCam.TemperatureSetpoint() = cameraTemp					
	objCam.RegulateTemperature() = 1
	currentTemp = objCam.Temperature
	
	while currentTemp < (cameraTemp - 1)									'
		currentTemp = objCam.Temperature
		wscript.Sleep 10000
	Wend
	
End Sub

Sub UnParkScope()
	status = objTele.UnPark()
End Sub

Sub ParkScope()
	status = objTele.Park()
End Sub

Sub checkObjectElev()
		Do while (objectAlt < 0) 
			CreateObject("WScript.Shell").Popup "objectAlt is " & round(objectAlt,0) & " sleeping for 60 seconds", 10, "Title"
			Wscript.Sleep 60000
		Loop
End Sub

Sub takeImage()
	objCam.ExposureTime = expTime
	objCam.Frame = 1 
	objCam.ImageReduction = 0
	objCam.AutoSaveOn = 1 
	status = objCam.TakeImage()
End Sub

Sub checkWeather()
	
	Dim WeatherFile
	Dim fso
	Const ForReading = 1
	Dim cloudCover
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set WeatherFile = fso.OpenTextFile(PathToWeatherFile, ForReading)
	Call GetWeatherfromFile(WeatherFile.ReadLine, cloudCover, rainFlag)
	
	if (cloudCover >=0) OR (rainFlag >=1) Then
		msgbox (cloudCover & " " & rainFlag)
	End If

End Sub

Sub GetWeatherfromFile(LineFromFile, cloudCover, rainFlag)

	cloudCover= Mid(LineFromFile,94,1)
	rainFlag= Mid(LineFromFile,98,1)

	
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