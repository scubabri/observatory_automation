
'this is based on the vbscript examples from bisque.com 

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
Dim TargetArray()
Dim ignoreErrors
Dim objectName
Dim status

Dim expTime
Dim imageCount
Dim imagesTaken
Dim imageScale
Dim maxObjMove
Dim SlewDelay
Dim SlewCountDown

Dim objectAlt
Dim minSlewElevation
Dim objectRate
Dim skySeeing
Dim objectHa
Dim objectTransit

Dim cameraTemp
Dim currentTemp

Dim enableWeather
Dim rainFlag

'This is where the work starts
Call InitGlobalUserVariables()
Call connectObjects()
Call createTargetList()
Call checkWeather()
Call setCamTemp()
Call UnparkScope
Call targetLoop()
Call parkScope()
Call DisconnectObjects()

Sub InitGlobalUserVariables()
	PathToTargetsFile = "D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO\output.txt"
	PathToWeatherFile = "D:\Dropbox\ASTRO\weatherdata.txt"
		
	imageScale = 1.95
	cameraTemp = -10
	enableWeather = 0
	minSlewElevation = 0
	maxObjMove = 8 						' max object move before closed loop slew is 8 arcmin
	skySeeing = 4
	
	'If you want your script to run all night regardless of errors, Set ignoreErrors = True
	ignoreErrors = False	
End Sub

Sub connectObjects()
	Set objTheSky 	= CreateObject("TheSkyX.Application") 
	Set objTele		= CreateObject("TheSkyX.sky6RASCOMTele") 
	Set objCam 		= CreateObject("TheSkyX.ccdsoftCamera") 
	objTele.Connect()
	objCam.Connect()
End Sub

Sub DisconnectObjects()
	objTele.Disconnect()
	objCam.Disconnect()
	Set objTheSky = Nothing
	Set objTele = Nothing
	Set objCam = Nothing
End Sub 

Sub createTargetList()
	Dim fso
	Dim TargetsFile
	Dim imageTime
	Const ForReading = 1	
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set TargetsFile = fso.OpenTextFile(PathToTargetsFile, ForReading)
	Do While (TargetsFile.AtEndOfStream <> True)
		targetString = TargetsFile.ReadLine
		ReDim Preserve TargetArray(i)
		TargetArray(i) = targetString
		'Wscript.Echo TargetArray(i)
		i=i+1
	Loop
	Set fso = Nothing
	TargetsFile.Close()
	
	For i = LBound(TargetArray) to UBound(TargetArray)
		For j = LBound(TargetArray) to UBound(TargetArray)
			If j <> UBound(TargetArray) Then
				If CDbl(Mid(TargetArray(j),18,7)) < CDbl(Mid(TargetArray(j + 1),18,7)) Then
					TempValue = TargetArray(j + 1)
					TargetArray(j + 1) = TargetArray(j)
					TargetArray(j) = TempValue
				End If
			End If
		Next
	Next
		'szobjectName = Mid(targetString,1,10)
		'objectName = "MPL " + szobjectName 
		'Call getCoordinates(objectName, objectRa, objectDec, objectAlt, objectRate)
		'Call getObjectTimes(objectName, objectRa, objectDec, objectTransit)
		'decimalTime = round(Hour(Now()) + (Minute(Now())/60 + (Second(Now())/3600)),4)	
End Sub

Sub GetExposureData()
	Call getCoordinates(objectName, objectRa, objectDec, objectAlt, objectRate)
	expTime = round((60*(imageScale/objectRate)*skySeeing),0)
	
	If (expTime >= 30) AND (expTime < 45) Then
		expTime = 30 
	ElseIf	(expTime >= 45) AND (expTime < 60) Then 
		expTime = 45 
	ElseIf expTime >= 60 Then 
		expTime = 60 
	End If
	
	imageCount = round((60*(60/expTime)),0)
	slewDelay  = round((60*maxObjMove)/objectRate,0)						' FOV is 23'X15', 8' to keep object in FOV, delay this many images before slew
End Sub

Sub GetObjectFromArray(szobjectName, vMag)
	szobjectName = Mid(targetArray(index),1,10)
	objectName = "MPL " + szobjectName 
	vMag = Mid(targetArray(index),37,4)
	
End Sub

Sub GetWeatherfromFile(LineFromFile, cloudCover, rainFlag)
	cloudCover= Mid(LineFromFile,94,1)
	rainFlag= Mid(LineFromFile,98,1)
End Sub

Sub checkWeather()
	
	Dim WeatherFile
	Dim fso
	Const ForReading = 1
	Dim cloudCover
	
	If enableWeather = 1 Then
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set WeatherFile = fso.OpenTextFile(PathToWeatherFile, ForReading)
		Call GetWeatherfromFile(WeatherFile.ReadLine, cloudCover, rainFlag)
	
		Do While (cloudCover >=2)
			if (cloudCover = 2) Then
				call objTele.SetTracking(0,1,0,0)
				cloudType = " Light "
				CreateObject("WScript.Shell").Popup cloudType& " clouds detected  sleeping for 60 seconds", 10, "Title"
				Wscript.Sleep 60000
			else
				cloudType = " Heavy "
				CreateObject("WScript.Shell").Popup cloudType& " clouds detected  exiting...", 10, "Title"
				Call parkScope()
				Call DisconnectObjects()
				Call DeleteObjects()
				WScript.Quit
			End If
			Call objTele.SetTracking(1,1,0,0)
			slewCountDown = slewDelay
		Loop
		WeatherFile.Close 
		Set WeatherFile = Nothing
		Set fso = Nothing
	End If
	
End Sub

Sub getCoordinates(objectName, objectRa, objectDec, objectAlt, objectRate)
	Set objTheSkyChart = CreateObject("TheSkyX.sky6StarChart") 
	status = objTheSkyChart.Find (objectName)
	Set objTheSkyInfo = CreateObject("TheSkyX.sky6ObjectInformation") 
	objTheSkyInfo.Index = 0 
	status = objTheSkyInfo.Property (54) 
	objectRa = objTheSkyInfo.ObjInfoPropOut 
	status = objTheSkyInfo.Property (55)
	objectDec = objTheSkyInfo.ObjInfoPropOut
	status = objTheSkyInfo.Property (59)
	objectAlt = objTheSkyInfo.ObjInfoPropOut
	status = objTheSkyInfo.Property (77)
	objectRateRA = objTheSkyInfo.ObjInfoPropOut
	status = objTheSkyInfo.Property (78)
	objectRateDEC = objTheSkyInfo.ObjInfoPropOut
	objectRate = round(sqr((objRateRA*objRateRA)+(objectRateDEC*objectRateDEC))*60,2)
	Set objTheSkyChart = Nothing
	Set objTheSkyInfo = Nothing
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

Sub closedLoopSlew(objectName)
	cdLight = 1														'Constant for frame type emumeration
	cdAutoDark = 3													'Constant for image reduction enumeration
	Set objTheSkyChart = CreateObject("TheSkyX.sky6StarChart") 
	set objCls = CreateObject("TheSkyX.closedLoopSlew")				'Create object for closedLoopSlew class
	objCam.ExposureTime = 10.0										'Set the exposure time
	objCam.Delay = 5.0												'Set an exposure delay
	objCam.Frame = cdLight											'Set a frame type
	objCam.ImageReduction = cdAutoDark								'Set for autodark
	objCam.Asynchronous = False										'Set for synchronous imaging (wait until done)
	status = objTheSkyChart.Find (objectName)						'Find objectName
	CreateObject("WScript.Shell").Popup "executing closed loop slew to " & objectName, 10, "closedLoopSlew"
	status = objCls.exec()
	Set objTheSkyChart = Nothing
	Set objCls = Nothing
End Sub

Sub targetLoop() 										' This is where the majority of the work takes place 
	'On Error Resume Next								' yes, we really want to do this to get to the error trap
	Dim bErrorOccurred
	Dim szobjectName
	Dim targetIndex
	targetIndex = 0
    	
	Do while targetIndex < UBound(targetArray)+1
		'Err.Clear 'Clear the error object
		'bErrorOccurred = False 'No error has occurred
				
		Call GetObjectFromArray(szobjectName, vMag)
		Call GetExposureData()
		
		CreateObject("WScript.Shell").Popup objectName & " vMag " & vMag & " Motion " & objectRate & " Imgcount " & imageCount & " Exptime " & expTime, 10, "Title"
		'Wscript.Quit
		imagesTaken = 0
		slewCountDown = slewDelay						' set this for initial closed loop slew
		
		Do While (imagesTaken <= imageCount)
			Call getCoordinates(objectName, objectRa, objectDec, objectAlt,objectRate)
			
			if (bErrorOccurred = False) then	
				Call checkWeather()
				Call checkObjectElev()
				
				if (slewCountDown >= slewDelay) Then
					Call closedLoopSlew(objectName)
					slewCountDown = 0 
				End If
				slewCountDown = slewCountDown + 1
				Call PromptOnError(bErrorOccurred)
				
				if (objectAlt < minSlewElevation) Then
					imagesTaken = imageCount
				End If
			End If
		
			if (bErrorOccurred = False) then
				Call takeImage()
				Call PromptOnError(bErrorOccurred)
				imagesTaken = imagesTaken + 1
			else 
				imagesTaken = imageCount + 1
			
			end if
		Loop
	Loop
End Sub

Sub setCamTemp()
	objCam.TemperatureSetpoint() = cameraTemp					
	objCam.RegulateTemperature() = 1
	currentTemp = objCam.Temperature	
	
	Do while (currentTemp >= (cameraTemp + 1))									'
		currentTemp = objCam.Temperature
		wscript.Sleep 10000
	Loop	

End Sub

Sub UnparkScope()
	status = objTele.UnPark()
End Sub

Sub parkScope()
	status = objTele.Park()
End Sub

Sub checkObjectElev()
		Do while (objectAlt < minSlewElevation)
			'Call objTele.SetTracking(0,1,0,0)
			Call getCoordinates(objectName, objectRa, objectDec, objectAlt, objectRate)
			CreateObject("WScript.Shell").Popup "objectAlt is " & round(objectAlt,0) & " sleeping for 5  minutes", 20, "Title"
			Wscript.Sleep 300000
		Loop
		'Call objTele.SetTracking(1,1,0,0)
End Sub

Sub takeImage()
	objCam.Delay = 0
	objCam.ExposureTime = expTime
	objCam.Frame = 1 
	objCam.ImageReduction = 0
	objCam.AutoSaveOn = 1 
	CreateObject("WScript.Shell").Popup "taking image " & imagesTaken + 1 & " of " & imageCount & " for " & objectName, 3, "takeImage"
	status = objCam.TakeImage()
End Sub

Sub getObjectTimes(objectName, objectRa, objectDec, objectTransit)
	Set objTheSkyUtils = CreateObject("TheSkyX.sky6Utils") 
	status = objTheSkyUtils.ComputeLocalSiderealTime()	
	myLst = objTheSkyUtils.dOut0
	status = objTheSkyUtils.ComputeHourAngle (objectRa) 
	objectHa = objTheSkyUtils.dOut0
	status = objTheSkyUtils.ComputeRiseTransitSetTimes (objectRa, objectDec) 
	objectTransit = objTheSkyUtils.dOut1
	Set objTheSkyUtils = Nothing
End Sub
