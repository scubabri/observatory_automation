set objShell = CreateObject("WScript.Shell")
'set c = CreateObject("ASCOM.Utilities.Chooser")
'c.DeviceType="SafetyMonitor"
'id = c.Choose("ASCOM.SafetyMonitor")
Set oc = CreateObject("ASCOM.Boltwood.ObservingConditions")
Set sm = CreateObject("ASCOM.Boltwood.OkToOpen.SafetyMonitor")

oc.Connected = True
sm.Connected = True
Dim cover
Dim safe
cover = oc.CloudCover
safe =  sm.IsSafe
msgbox safe
oc.Connected = False
sm.Connected = False




