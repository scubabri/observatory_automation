# observatory_automation
These are scripts that I've cobbled together to work around limitations in my mount, roof, or otherwise.

Power switching is provided via an APC SNMP enabled PDU, https://www.amazon.com/APC-AP7900-Switched-Surge-Protector/dp/B0000AAAYH

I relied on https://tobinsramblings.wordpress.com/2011/05/03/snmp-tutorial-apc-pdus/ on setting up snmp to be able to power on/off ports. 

Roof control is handled by http://interactiveastronomy.com/skyroof.html

Weather monitoring is handled by http://interactiveastronomy.com/skyalertindex.html

imaging automation is handled by http://ccdware.com/products/ccdap5/ and/or ccd commander http://ccdcommander.com/

Scripts are dependent on ASCOM 6.3 https://ascom-standards.org/

Boltwood safety driver/Observing conditions https://ascom-standards.org/Downloads/SafetyMonitorDrivers.htm

Ascom Driver Access (reference for scripting) http://www.ascom-standards.org/Help/Developer/html/N_ASCOM_DriverAccess.htm

imaging and camera control is handled by MaximDL5 http://www.diffractionlimited.com/help/maximdl/MaxIm-DL.htm, see "Scripting" section or TheSkyX http://www.bisque.com/scriptTheSkyX/




