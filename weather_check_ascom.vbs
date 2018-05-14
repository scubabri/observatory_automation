Dim safe
Dim counter
counter = 1

Set sm = CreateObject("ASCOM.Boltwood.OkToOpen.SafetyMonitor")
sm.Connected = True

Do Until safe = True

    safe = sm.IsSafe
	msgBox safe			
	If safe = True Then
	   Exit Do
	Else						' not safe to contnue, lets wait up to 10 minutes
		If counter >= 11 Then 
			Exit Do
		Else
			counter = counter + 1
			wscript.sleep(60000)
		End If
	
	End If
		
Loop
	

If safe = True Then 
	msgBox "Sky is clear, resuming"
	sm.Connected = False
	
	Else	
	msgBox "Sky is cloudy, not continuing"
	sm.Connected = False
	Wscript.Quit
	
End If