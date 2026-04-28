''Note: This executable file used to avoid locking the machine while automation execution
set WshShell = CreateObject("WScript.Shell")
Do
	WshShell.SendKeys "{NUMLOCK}"
	WshShell.SendKeys "{NUMLOCK}"
	WScript.Sleep (60*1000)
Loop
set WshShell =nothing
