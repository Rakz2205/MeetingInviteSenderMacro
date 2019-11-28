characterLength=200
StartChar="Rakesh"
EndChar="Gupta"




















































































































































































































messageToEnter=StartChar & generateRandomString((characterLength-len(StartChar)-len(EndChar)),5)  & EndChar

'msgbox messageToEnter
'strPassword="RK House , BUnglow no 1 , Geminin Co-op Housing Society , Anushakti Nagar , Mankhurd - east , Mumbai 400088"
Set WshShell = CreateObject("WScript.Shell")
WScript.sleep 2000
For nPassword = 1 To Len(messageToEnter)
	strTemp=Mid(messageToEnter,nPassword,1)
	WshShell.SendKeys strTemp
	'WScript.sleep 5
Next
 

Public Function randomGenerator(rangeval)
    randomGenerator = Int(rangeval * Rnd) + 1
End Function

Function generateRandomString(ByVal strLen, ByVal intVal)
	
	Dim str
	Dim strLetters
	Select Case intVal
		Case 1		'Capital alphabets
			 strLetters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
		Case 2		'Small case alphabets
			 strLetters = "abcdefghijklmnopqrstuvwxyz"
		Case 3		'Numbers
			 strLetters = "0123456789"
		Case 4		'Special characters
			 strLetters = "!*=?"
		Case 5		'Alpha-numeric and special characters
			 strLetters = "abcdetungd1234567!*="
	End Select
	
    For i = 1 to strLen
        str = str & Mid(strLetters, randomGenerator(Len(strLetters)), 1)
    Next
    generateRandomString = str
End Function