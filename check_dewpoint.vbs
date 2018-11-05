
Set oc = CreateObject("ASCOM.Boltwood.ObservingConditions")
Set roof = CreateObject("ASCOM.SkyRoofHub.Dome")             
Set objShell = CreateObject("WScript.Shell")

oc.Connected = True
roof.connected = true 

wscript.sleep(3000)                                       

If roof.shutterstatus = 0 Then
	
	Do Until roof.shutterstatus = 1
		tempc = oc.Temperature
		dewpointc = oc.Dewpoint
		spreadc = (tempc - dewpointc)
		if (spreadc < 5) Then
			msgbox spreadc
			objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.4 i 1",0, True 'power on dew heater
		End If
		wscript.sleep(60000)
	Loop
	
	objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.4 i 0",0, True 'power off dew heater

	End If



oc.Connected = False
roof.connected = False

Set oc = Nothing
Set roof = Nothing