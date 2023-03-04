do
Set wshshell = wscript.CreateObject("WScript.Shell")
Wshshell.run "Notepad"
wscript.sleep 100
wshshell.sendkeys "H"
wscript.sleep 100
wshshell.sendkeys "A"
wscript.sleep 100
wshshell.sendkeys "P"
wscript.sleep 100
wshshell.sendkeys "P"
wscript.sleep 100
wshshell.sendkeys "Y"
loop
md 666
md 666
cd 666
cd 666