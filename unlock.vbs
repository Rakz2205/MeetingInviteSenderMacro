start = now()
count=0
do
	set wshShell=CreateObject("Wscript.Shell")
	wshShell.SendKeys "{SCROLLLOCK}"
	Wscript.Sleep 1
	wshShell.SendKeys "{SCROLLLOCK}"
	Wscript.Sleep 100000
	count=count+1
loop Until Count=0
msgbox start&"-----"&now()