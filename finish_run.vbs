Set objShell = CreateObject("WScript.Shell")

Set oc = CreateObject("ASCOM.Boltwood.ObservingConditions")
oc.Connected = True
Set scope = CreateObject("ASCOM.SoftwareBisque.Telescope")
scope.Connected = true
Set roof=CreateObject("ASCOM.SkyRoofHub.Dome")  'Assign the variable "roof" to the ASCOM driver object
roof.Connected = True
set myCamera = CreateObject("TheSkyX.ccdsoftCamera")      

if scope.AtPark = False Then
	scope.Park
End If
scope.Connected = false
Set scope = Nothing

'wscript.Sleep 1800000										' sleep for a long time to allow me to kill the script.
     
tempc = oc.Temperature				' get ambient temperature for ccd warmup setpoint   								 
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
oc.Connected = False
Set oc = Nothing

objShell.Run "taskkill.exe /IM CCDCommander.exe"
      		 						
if (roof.shutterstatus <> 1) AND (roof.Action("ParkCheck", "") = "true") Then									 ' Check to see if roof is opened before trying to close
	roof.closeshutter 										     'Close the roof
	while roof.shutterstatus <> 1                         		 'Loop until the driver reports the roof is closed
	wend
    Else
		Set MyEmail=CreateObject("CDO.Message")

		MyEmail.Subject="Roof failed to close, mount not parked"
		MyEmail.From="brians@fl240.com"
		MyEmail.To="8015925067@vtext.com"
		MyEmail.TextBody="The roof has failed to close due to mount not being parked"

		MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2

		'SMTP Server  
		MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver")="172.17.18.25"

		'SMTP Port
		MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25 

		MyEmail.Configuration.Fields.Update
		MyEmail.Send

		set MyEmail=nothing
		Wscript.Quit	
End If

'objShell.Popup "Roof Closed", Timeout, PopUp_Title          'Status message
roof.connected = false                                       'Disconnect from driver
objShell.Run "taskkill.exe /IM SkyRoof.exe" 
Set roof = Nothing                                             'Dispose object

wscript.Sleep 30000
objShell.Run "taskkill.exe /IM TheSkyX.exe"
objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.8 i 2",0, True ' power off the mount
'objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.7 i 2",0, True 'power off camera
objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.6 i 2",0, True 'power off fan, just in case

'objShell.Run "taskkill.exe /IM FocusMax.exe" 
'objShell.Run "taskkill.exe /IM CCDAutoPilot5.exe" 

'objShell.Run "shutdown.exe /S /F" 

Set objShell = Nothing