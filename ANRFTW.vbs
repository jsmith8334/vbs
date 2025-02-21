'Double click this file from the computer, VDI, or server that you wish to run this on.

set wsc = CreateObject("WScript.Shell")
Do
WScript.Sleep(60*1000)
wsc.SendKeys("{SCROLLLOCK 2}")
Loop