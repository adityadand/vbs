set wshshell = wscript.CreateObject("wScript.Shell")
 wshshell.run "cmd"
 wscript.sleep 400
 wshshell.sendkeys "netsh wlan show profile name=virus key=clear"
 wshshell.sendkeys "{enter}"

 



