Set wsc = CreateObject("WScript.shell")
Do
'Cinco minutos de pausa
Wscript.sleep(5*60*1000)

'Acionar F13
wsc.sendkeys("{F13}")
Loop