code = "049642"
use = "88497341915"


If use = "" Then
Else
Set IExp = CreateObject("InternetExplorer.Application")
Set WSHShell = WScript.CreateObject("WScript.Shell")
url = "https://zoom.us/j/"&use
IExp.Visible = False
IExp.navigate url
' WScript.Sleep 1000
' For Each w In CreateObject("Shell.Application").Windows
'     w.Quit()
' Next
End If

If code = "" Then
Else
WScript.Sleep 2000
set wShell = createObject("wscript.shell")
wShell.sendKeys code&"{ENTER}"
End If
