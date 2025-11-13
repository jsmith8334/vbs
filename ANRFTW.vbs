Set WShell = CreateObject("WScript.Shell")

Do
WShell.SendKeys "{RIGHT 2}"
WShell.SendKeys "{LEFT 2}"
WScript.Sleep(60*1000)
Loop
