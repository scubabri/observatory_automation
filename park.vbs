'set c = CreateObject("ASCOM.Utilities.Chooser")
'c.DeviceType="Telescope"
'id = c.Choose("ASCOM.Celestron.Telescope")
set scope = CreateObject("ASCOM.SoftwareBisque.Telescope")
scope.Connected = true

MsgBox "Press enter to park scope ", 0, "Press enter to park scope"
scope.Park
Msgbox "Scope Parked."

'msgBox id
scope.Connected = false