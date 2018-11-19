Set objShell = CreateObject("WScript.Shell")

Dim safe
Dim counter
Dim rcounter
counter = 1
rcounter = 1

'objShell.run "C:\Users\brians\Dropbox\ASTRO\Software\sunwait.exe wait astronomical set offset +02:00:00 40N 111W",0, True

Set sm = CreateObject("ASCOM.Boltwood.OkToOpen.SafetyMonitor")
Set oc = CreateObject("ASCOM.Boltwood.ObservingConditions")
sm.Connected = True
oc.Connected = True

safe = False

Do Until safe = True

	'safe = sm.IsSafe         ' this doesnt work if I need to open and its still light out.
	clouds = oc.CloudCover
	rain = oc.RainRate
	
	If (clouds < 50) AND (rain = 0) Then 
			safe = true
		else
			If (counter = 1) Then
				safe = false
				Set MyEmail=CreateObject("CDO.Message")

				MyEmail.Subject="Weather is not safe, waiting"
				MyEmail.From="brians@fl240.com"
				MyEmail.To="8015925067@vtext.com"
				MyEmail.TextBody="Weather conditions are not safe to open"
				MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2

				'SMTP Server  
				MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver")="172.17.18.25"
				'SMTP Port
				MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25 

				MyEmail.Configuration.Fields.Update
				MyEmail.Send

				set MyEmail=nothing
			Else 
				safe = false
		End If 
	End If
	
	If safe = True Then
	   sm.Connected = False
	   oc.Connected = False
	   Exit Do
	Else						' not safe to contnue, lets wait up to 60 minutes
		If counter >= 360 Then 
		    sm.Connected = False
			oc.Connected = False
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

If roof.shutterstatus = 1 Then

	roof.openshutter   
																   'Open the roof
	Do Until roof.shutterstatus = 0									'Loop until the driver reports the roof is open
	
		If rcounter >= 30 Then
			roof.closeshutter
			Set MyEmail=CreateObject("CDO.Message")

			MyEmail.Subject="Roof failed to open"
			MyEmail.From="brians@fl240.com"
			MyEmail.To="8015925067@vtext.com"
			MyEmail.TextBody="The roof has failed to open"

			MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2

			'SMTP Server  
			MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver")="172.17.18.25"

			'SMTP Port
			MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25 

			MyEmail.Configuration.Fields.Update
			MyEmail.Send

			set MyEmail=nothing
			Wscript.Quit
			Exit Do
		Else
			rcounter = rcounter + 1
			wscript.sleep(1000)
		End If
	
	Loop
	
	Set MyEmail=CreateObject("CDO.Message")

    MyEmail.Subject="Roof opening"
    MyEmail.From="brians@fl240.com"
    MyEmail.To="8015925067@vtext.com"
    MyEmail.TextBody="The roof has been opened"

    MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2

    'SMTP Server  
    MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver")="172.17.18.25"

    'SMTP Port
    MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25 

    MyEmail.Configuration.Fields.Update
    MyEmail.Send

	set MyEmail=nothing
End If

If roof.shutterstatus = 0 Then

	objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.6 i 1",0, True 'power on Fan to aid cooling
	objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.7 i 1",0, True 'power on the camera
	objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.8 i 1",0, True 'power on the mount

	' need to put checks to verify devices powered on.
	' need to put checks to verify scope, camera, focuser and roof status.

	'wscript.sleep(1800000) 								        'Sleep for 60 minutes before running ccdap.
	'lets change this to a loop to keep an eye on clouds/rain after opening the roof
	'objShell.run """C:\Program Files (x86)\CCDWare\CCDAutoPilot5\CCDAutoPilot5.exe""",0, False
    objShell.run """C:\Program Files (x86)\Software Bisque\TheSkyX Professional Edition\TheSkyX.exe""",4, False
	wscript.sleep(15000)
	'objShell.run "C:\CCD Commander\CCDCommander.exe ""AutoRun "C:\CCD Commander\Actions\NEOCP_11_18_2018_.act"""",4, False
	'objShell.run """C:\CCD Commander\CCDCommander.exe AutoRun "C:\CCD Commander\Actions\NEOCP_11_18_2018_.act"""",4, False
	
	Dim strPath1, strAttr1, strAttr2

    strPath1 = """C:\CCD Commander\CCDCommander.exe"""
    strAttr1 = " AutoRun "
    strAttr2 = """C:\CCD Commander\Actions\NEOCP_11_18_2018.act"""

	objShell.Run strPath1 & strAttr1 & strAttr2 

Else 

    MyEmail.Subject="Roof failure"
    MyEmail.From="brians@fl240.com"
    MyEmail.To="8015925067@vtext.com"
    MyEmail.TextBody="The roof didnt open, aborting startup"

    MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2

    'SMTP Server  
    MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver")="172.17.18.25"

    'SMTP Port
    MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25 

    MyEmail.Configuration.Fields.Update
    MyEmail.Send

	set MyEmail=nothing
	
End If

roof.connected = false
Set objShell = Nothing
Set roof = Nothing
