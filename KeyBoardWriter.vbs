strPassword="RK House , BUnglow no 1 , Geminin Co-op Housing Society , Anushakti Nagar , Mankhurd - east , Mumbai 400088"
Set WshShell = CreateObject("WScript.Shell")
WScript.sleep 2000
For nPassword = 1 To Len(strPassword)
			strTemp=Mid(strPassword,nPassword,1)
			WshShell.SendKeys strTemp
			WScript.sleep 5
		Next
 
