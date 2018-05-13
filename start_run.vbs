Set objShell = CreateObject("WScript.Shell")
objShell.run "C:\Users\brians\Dropbox\ASTRO\Software\sunwait.exe wait set offset +01:00:00 40N 111W",0, True
 
Set roof=CreateObject("ASCOM.SkyRoofHub.Dome")              'Assign the variable "roof" to the ASCOM driver object
'Set objShell = WScript.CreateObject("WScript.Shell")       'Shell for PopUp messages
'Const Timeout = 3                                           'Constant for PopUp message display time
'Const PopUp_Title = "SkyRoof Driver Script"                'PopUp message title
roof.connected = true                                       'Connect to the driver
wscript.sleep(3000)                                         'Wait a few seconds for connection to driver
roof.openshutter                                            'Open the roof
'objShell.Popup "Opening Roof...", Timeout, PopUp_Title     'Status message
while roof.shutterstatus <> 0                               'Loop until the driver reports the roof is open
											    ' Need to add timeout here.
wend
'objShell.Popup "Roof Open", Timeout, PopUp_Title           'Roof is open
roof.connected = false

objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.6 i 1",0, True
objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.7 i 1",0, True
objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.8 i 1",0, True

wscript.sleep(60000) 								        'Sleep for 60 seconds for things to settle 
objShell.run """C:\Program Files (x86)\CCDWare\CCDAutoPilot5\CCDAutoPilot5.exe""",0, False

Set objShell = Nothing
Set root = Nothing