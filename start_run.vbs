Set objShell = CreateObject("WScript.Shell")

Dim safe
Dim counter
counter = 1

objShell.run "C:\Users\brians\Dropbox\ASTRO\Software\sunwait.exe wait set offset +01:00:00 40N 111W",0, True

Set sm = CreateObject("ASCOM.Boltwood.OkToOpen.SafetyMonitor")
Set oc = CreateObject("ASCOM.Boltwood.ObservingConditions")
sm.Connected = True
oc.Connected = True

safe = False

Do Until safe = True

    'safe = sm.IsSafe
	clouds = oc.CloudCover
	rain = oc.RainRate
	msgBox clouds
	msgBox rain
	
	If (clouds < 60) AND (rain = 0) Then 
			safe = true
		else
			safe = false
	End If
	
	If safe = True Then
	   sm.Connected = False
	   oc.Connected = False
	   Exit Do
	Else						' not safe to contnue, lets wait up to 30 minutes
		If counter >= 30 Then 
		    sm.Connected = False
			Exit Do
		Else
			counter = counter + 1
			wscript.sleep(60000)
		End If
	
	End If
		
Loop

If safe = True Then 
	'msgBox "Safe to open, continuing..."
	sm.Connected = False
	oc.Connected = False
	
	Else	
	msgBox "Not safe to open, exiting..."
	sm.Connected = False
	oc.Connected = False
	Wscript.Quit
	
End If
 
Set roof=CreateObject("ASCOM.SkyRoofHub.Dome")              'Assign the variable "roof" to the ASCOM driver object
roof.connected = true                                       'Connect to the driver
wscript.sleep(3000)                                         'Wait a few seconds for connection to driver
roof.openshutter                                            'Open the roof

while roof.shutterstatus <> 0                               'Loop until the driver reports the roof is open
' Need to add timeout here.
wend

roof.connected = false

objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.6 i 1",0, True
objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.7 i 1",0, True
objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.8 i 1",0, True

' need to put checks to verify devices powered on.

wscript.sleep(60000) 								        'Sleep for 60 seconds for things to settle 
objShell.run """C:\Program Files (x86)\CCDWare\CCDAutoPilot5\CCDAutoPilot5.exe""",0, False

Set objShell = Nothing
Set roof = Nothing