Set roof=CreateObject("ASCOM.SkyRoofHub.Dome")              'Assign the variable "roof" to the ASCOM driver object
roof.connected = true                                       'Connect to the driver
wscript.sleep(3000)    

'roof.openshutter 
roof.closeshutter