Dim cloudy
cloudy = 1
Dim counter
counter = 1

Do Until cloudy = 0

    On Error Resume Next
	Set fso = CreateObject("Scripting.FileSystemObject")
    set src = fso.OpenTextFile("c:\Users\brians\Dropbox\ASTRO\weatherdata.txt",1)  ' read "boltwood II from SkyAlert
	
	Dim strSearchString
	strSearchString = src.readall()
	
	If InStr(1, strSearchString, "1 1 1 1 0 0") > 0 then			' No clouds, rain, alerts and dark
		cloudy = "0"												' safe to continue 
		msgBox "1 1 1 1 0 0"
	
	ElseIf InStr(1, strSearchString, "1 1 1 2 0 0") > 0 then		' no clouds, rain, alerts and dim
		cloudy = "0"												' safe to continue
		msgBox "1 1 1 2 0 0"									
	
	ElseIf InStr(1, strSearchString, "1 1 1 3 0 0") > 0 then		' no clouds, rain, alerts and day
		cloudy = "0"												' safe to continue
		msgBox "1 1 1 3 0 0"
	
	Else
		cloudy = "1"												' not safe to contnue, lets wait up to 10 minutes
		msgBox strSearchString
		wscript.sleep(60000)
		
    If counter >= 11 Then Exit Do
	counter = counter + 1
	
	End If
		
Loop
	

If cloudy = 0 Then 
	msgBox "Sky is clear, resuming"
	
	Else	
	msgBox "Sky is cloudy, not continuing"
	Wscript.Quit
	
End If

MsgBox cloudy