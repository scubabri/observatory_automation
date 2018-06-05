
Set objShell = WScript.CreateObject("WScript.Shell")
DIM returnValue
Const Timeout = 3   
Const PopUp_Title = "SkyRoof Driver Script" 

'need to put checks to see if scope is powered up   

Set roof=CreateObject("ASCOM.SkyRoofHub.Dome")        'Assign the variable "roof" to the ASCOM driver object
roof.connected = true 
set scope = CreateObject("ASCOM.Celestron.Telescope")
scope.Connected = true

if (roof.shutterstatus = 0) And (scope.AtPark = True) Then
	
	'MsgBox "Press enter to unpark the scope ", 0, "Press enter to unpark scope"
	scope.UnPark
	scope.FindHome
	'Msgbox "Scope unparked and tracking"
	objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.4 i 2",0, True 'power off fan

Else 

	objShell.Popup "Roof is not open, aborting.", Timeout, PopUp_Title 
	
End If

set scope = Nothing
set roof = Nothing

