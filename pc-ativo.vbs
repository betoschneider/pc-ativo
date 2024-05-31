Set wsc = CreateObject("WScript.shell")
Do
'Five minutes
Wscript.sleep(5*60*1000)
wsc.sendkeys("{F13}")
Loop