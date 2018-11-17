
Set objShell = WScript.CreateObject("WScript.Shell")
DIM returnValue
Const Timeout = 3   
Const PopUp_Title = "SkyRoof Driver Script" 

'need to put checks to see if scope is powered up   

Set roof=CreateObject("ASCOM.SkyRoofHub.Dome")        'Assign the variable "roof" to the ASCOM driver object
roof.connected = true 
set scope = CreateObject("ASCOM.SoftwareBisque.Telescope")
scope.Connected = true

if roof.shutterstatus = 0 Then
	
	'MsgBox "Press enter to unpark the scope ", 0, "Press enter to unpark scope"
	scope.UnPark()
	scope.FindHome()
	'Msgbox "Scope unparked and tracking"
	objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.6 i 2",0, True 'power off fan
	objShell.run """C:\Users\brians\AppData\Local\Apps\2.0\1VAJZAH0.0TT\NRRZANR1.0QO\skyr..tion_d2275ca0e4e6fd85_0001.0000_fde1d24b8ff66f56\SkyRoof.exe""",4, False

Else 

	objShell.Run "taskkill.exe /IM CCDAutoPilot5.exe" 
	objShell.Run "taskkill.exe /IM TheSkyX.exe" 
	'objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.6 i 2",0, True 'power off focuser 
    objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.7 i 2",0, True 'power off camera
    objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.6 i 2",0, True 'power off fan, just in case
	objShell.run "C:\usr\bin\snmpset.exe -v 1 -c private bs-obspdu.fl240.com PowerNet-MIB::sPDUOutletCtl.8 i 2",0, True 'power on the mount
	'objShell.Popup "Roof is not open, aborting.", Timeout, PopUp_Title 
	
	Set MyEmail=CreateObject("CDO.Message")

    MyEmail.Subject="Failed to find home, roof closed"
    MyEmail.From="brians@fl240.com"
    MyEmail.To="8015925067@vtext.com"
    MyEmail.TextBody="The mount was not homed due to the roof not being open"

    MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2

    'SMTP Server  
    MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver")="172.17.18.25"

    'SMTP Port
    MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25 

    MyEmail.Configuration.Fields.Update
    MyEmail.Send

	set MyEmail=nothing
	
End If

set scope = Nothing
set roof = Nothing

