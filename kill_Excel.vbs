Option Explicit
Dim objWMIService, objProcess, colProcess
Dim strComputer, strProcessKill
strComputer = "."
strProcessKill = "'EXCEL.exe'"
 
Set objWMIService = GetObject("winmgmts:\\"&strComputer&"\root\cimv2")
 
Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name = " & strProcessKill )
For Each objProcess in colProcess
objProcess.Terminate()
Next
msgbox "Done"