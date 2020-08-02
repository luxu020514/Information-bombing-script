msgbox("开始了吗？")
set wshshell=wscript.createobject("wscript.shell")
WshShell.AppActivate"测试"
for i=0 to 100
WScript.Sleep 250
WshShell.Sendkeys"^v"
WshShell.Sendkeys i
WshShell.Sendkeys"%s"
Next
