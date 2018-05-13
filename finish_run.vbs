Set objShell = CreateObject("WScript.Shell")

Set scope = CreateObject("ASCOM.Celestron.Telescope")
'scope.Connected = true
'scope.Park
scope.Connected = false

objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.8 i 2",0, True

Set scope = Nothing

Set roof=CreateObject("ASCOM.SkyRoofHub.Dome")        		 'Assign the variable "roof" to the ASCOM driver object
'objShell.Popup "Closing Roof...", Timeout, PopUp_Title      'Status message
roof.closeshutter 'Close the roof
while roof.shutterstatus <> 1                         		 'Loop until the driver reports the roof is closed
wend
'objShell.Popup "Roof Closed", Timeout, PopUp_Title          'Status message
roof.connected = false                                       'Disconnect from driver

Set roof=Nothing                                             'Dispose object

Set myCamera = CreateObject("MaxIm.CCDCamera")               'The CCD camera object (and guider) in Maxim
myCamera.LinkEnabled() = True    							 'Link CCD camera to Maxim

If myCamera.LinkEnabled Then
myCamera.GuiderStop()        								 'Stop the guider
wscript.Sleep 5000           								 'wait 5 seconds 

MyCamera.CoolerOn = True
MyCamera.TemperatureSetpoint = 10
Temp = MyCamera.Temperature()

while Temp < 0
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

objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.6 i 2",0, True
objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.7 i 2",0, True

Set objShell = Nothing
