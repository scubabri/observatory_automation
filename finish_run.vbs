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
Set scope = Nothing

'wscript.Sleep 1800000										' sleep for a long time to allow me to kill the script.

set myCamera = CreateObject("TheSkyX.ccdsoftCamera")         
    								 
myCamera.Connect()
myCamera.TemperatureSetpoint() = (tempc + 5)					' set warmup setpoint to ambient temp + 5c 
myCamera.RegulateTemperature() = 1
Temp = myCamera.Temperature

while Temp < (tempc - 5)									' when ccd temp reaches ambient - 5c, its probably safe to shut down the camera
	Temp = myCamera.Temperature
	wscript.Sleep 10000
Wend
	
myCamera.RegulateTemperature() = 0
	
wscript.Sleep 5000
myCamera.Disconnect()    							'Disconnect the camera from Maxim (if connected)
Set myCamera = Nothing

objShell.Run "taskkill.exe /IM CCDCommander.exe"
wscript.Sleep 30000
objShell.Run "taskkill.exe /IM TheSkyX.exe"
objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.8 i 2",0, True ' power off the mount

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
objShell.Run "taskkill.exe /IM SkyRoof.exe" 
Set roof = Nothing                                             'Dispose object


'objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.6 i 2",0, True 'power off focuser 
'objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.7 i 2",0, True 'power off camera
objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.6 i 2",0, True 'power off fan, just in case

'objShell.Run "taskkill.exe /IM FocusMax.exe" 
'objShell.Run "taskkill.exe /IM CCDAutoPilot5.exe" 

'objShell.Run "shutdown.exe /S /F" 

Set objShell = Nothing
