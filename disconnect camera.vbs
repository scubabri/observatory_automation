Set myCamera = CreateObject("MaxIm.CCDCamera")       'The CCD camera object (and guider) in Maxim
myCamera.LinkEnabled() = True    'Link CCD camera to Maxim

If myCamera.LinkEnabled Then
myCamera.GuiderStop()        'Stop the guider
wscript.Sleep 5000           'wait 5 seconds 

MyCamera.CoolerOn = True
MyCamera.TemperatureSetpoint = 10

Temp = MyCamera.Temperature()
while Temp < 0
	Temp = MyCamera.Temperature()
	wscript.Sleep 10000
Wend

myCamera.CoolerOn = False    'Turn the cooler off
End If

wscript.Sleep 5000

If myCamera.LinkEnabled Then
MyCamera.LinkEnabled = False    'Disconnect the camera from Maxim (if connected)
End If

Set myCamera = Nothing
