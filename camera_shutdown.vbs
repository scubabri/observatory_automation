Set objShell = CreateObject("WScript.Shell")

'set myCamera = CreateObject("CCDSoft2XAdaptor.CCDSoft5Camera")
set myCamera = CreateObject("TheSkyX.ccdsoftCamera")
'c.DeviceType="Camera"
'c.DeviceType="Dome"
'id = c.Choose("SoftwareBisque")
'id = c.Choose("POTH.Dome")
'Set oc = CreateObject("ASCOM.Boltwood.ObservingConditions")
'Set sm = CreateObject("ASCOM.Boltwood.OkToOpen.SafetyMonitor")

'Set myCamera = CreateObject("CCDSoft2XAdaptor.CCDSoft5Camera")               'The CCD camera object (and guider) in Maxim 
'myCamera.Disconnect()    							 'Link CCD camera to Maxim




 								 
	myCamera.Connect()
	myCamera.TemperatureSetpoint() = -10					' set warmup setpoint to ambient temp + 5c 
	myCamera.RegulateTemperature() = 1
	Temp = myCamera.Temperature

	while Temp > -10									' when ccd temp reaches ambient - 5c, its probably safe to shut down the camera
		Temp = myCamera.Temperature
		wscript.Sleep 10000
	Wend
	
	myCamera.TemperatureSetpoint() = 15					' set warmup setpoint to ambient temp + 5c 
	myCamera.RegulateTemperature() = 1
	
	while Temp < 15								' when ccd temp reaches ambient - 5c, its probably safe to shut down the camera
		Temp = myCamera.Temperature
		wscript.Sleep 10000
	Wend
	
	myCamera.Disconnect()
	Set myCamera = Nothing
	
	
'dim isCameraConnected
'dim ccd
'set ccd = CreateObject("TheSkyX.ccdsoftCamera")
'isCameraConnected = (0 = ccd.Connect())

'if isCameraConnected then
'    ccd.TemperatureSetPoint() = 10
'    ccd.RegulateTemperature() = true
'end if

'if isCameraConnected then
'    [Statements]
'end if

'if ccd.RegulateTemperature <> 0 then
'    [Statements]
'end if