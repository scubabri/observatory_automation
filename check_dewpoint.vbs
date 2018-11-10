
Set oc = CreateObject("ASCOM.Boltwood.ObservingConditions")
Set roof = CreateObject("ASCOM.SkyRoofHub.Dome")             
Set objShell = CreateObject("WScript.Shell")

oc.Connected = True
roof.connected = true 

wscript.sleep(3000)                                       

If roof.shutterstatus = 0 Then
	
	    tempc = oc.Temperature
		dewpointc = oc.Dewpoint
		spreadc = (tempc - dewpointc)
		humidity = oc.Humidity
		
		If (spreadc < 12) OR (humidity > 50) Then
			'msgbox spreadc
			objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.4 i 1",0, True 'power on dew heater
		Else
			If (spreadc > 12) AND (humidity < 50) Then
				objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.4 i 2",0, True 'power off dew heater
			End If
		End If
		
End If

If roof.shutterstatus = 1 Then
	objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.4 i 2",0, True 'power off dew heater
End If

oc.Connected = False
roof.connected = False

Set oc = Nothing
Set roof = Nothing