Set objShell = CreateObject("WScript.Shell")

Set oc = CreateObject("ASCOM.Boltwood.ObservingConditions")
oc.Connected = True

tempc = oc.Temperature				' get ambient temperature for ccd warmup setpoint
oc.Connected = False

Set oc = Nothing

Set scope = CreateObject("ASCOM.SoftwareBisque.Telescope")
scope.Connected = true

if scope.AtPark = False Then
	scope.Park
End If

scope.Connected = false

'objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.8 i 2",0, True ' power off the mount

Set scope = Nothing

Set roof=CreateObject("ASCOM.SkyRoofHub.Dome")
roof.Connected = True        		 						 'Assign the variable "roof" to the ASCOM driver object
'objShell.Popup "Closing Roof...", Timeout, PopUp_Title      'Status message

if roof.shutterstatus <> 1 Then									 ' Check to see if roof is opened before trying to close
	roof.closeshutter 										     'Close the roof
	while roof.shutterstatus <> 1                         		 'Loop until the driver reports the roof is closed
	wend

End If

'objShell.Popup "Roof Closed", Timeout, PopUp_Title          'Status message
roof.connected = false                                       'Disconnect from driver

Set roof = Nothing                                             'Dispose object

Set myCamera = CreateObject("MaxIm.CCDCamera")               'The CCD camera object (and guider) in Maxim
myCamera.LinkEnabled() = True    							 'Link CCD camera to Maxim

If myCamera.LinkEnabled Then
	myCamera.GuiderStop()        								 'Stop the guider
	wscript.Sleep 5000           								 'wait 5 seconds 

	MyCamera.CoolerOn = True
	MyCamera.TemperatureSetpoint = (tempc + 5)					' set warmup setpoint to ambient temp + 5c 
	Temp = MyCamera.Temperature()

	while Temp < (tempc - 5)									' when ccd temp reaches ambient - 5c, its probably safe to shut down the camera
		Temp = MyCamera.Temperature()
		wscript.Sleep 10000
	Wend
	
	myCamera.CoolerOn = False    								'Turn the cooler off
End If

wscript.Sleep 5000

If myCamera.LinkEnabled Then
	MyCamera.LinkEnabled = False    							'Disconnect the camera from Maxim (if connected)
End If

Set myCamera = Nothing

'objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.6 i 2",0, True 'power off focuser 
objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.7 i 2",0, True 'power off camera
objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.4 i 2",0, True 'power off fan, just in case


'objShell.Run "shutdown.exe /R /T 5 /C ""Rebooting your computer now!"" "

Set objShell = Nothing
