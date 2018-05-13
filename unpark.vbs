'set c = CreateObject("ASCOM.Utilities.Chooser")
'c.DeviceType="Telescope"
'id = c.Choose("ASCOM.Celestron.Telescope")

Set objShell = WScript.CreateObject("WScript.Shell")
DIM returnValue
Const Timeout = 3   
Const PopUp_Title = "SkyRoof Driver Script"   

Set roof=CreateObject("ASCOM.SkyRoofHub.Dome")        'Assign the variable "roof" to the ASCOM driver object
roof.connected = true 

if roof.shutterstatus = 0 Then

set scope = CreateObject("ASCOM.Celestron.Telescope")
scope.Connected = true

'MsgBox "Press enter to unpark the scope ", 0, "Press enter to unpark scope"
scope.UnPark
scope.FindHome
'Msgbox "Scope unparked and tracking"

Else 

objShell.Popup "Roof is not open, aborting.", Timeout, PopUp_Title 

End If

set scope = Nothing
set roof = Nothing

