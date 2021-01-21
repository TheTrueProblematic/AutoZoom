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






















' today = Date
' snd = "1/1/3000"
' week = Weekday(Date)
' hr = Hour(Time)
' min = Minute(Time)
' tm = hr&":"&min
' nicetime = TimeValue(tm)
'
' tim = InputBox("What time is class "&ii,"ZoomLauncher","Use Military Time (Example 13:45)")
' Dim dow
' dow = Array("Monday","Tuesday","Wednessday","Thursday","Friday")
' build = ""
' For i = 0 to 4
' tday = dow(i)
' intAnswer = _
'     MsgBox("Do you have class on "&tday&"?", _
'         vbYesNo, "ZoomLauncher")
' If intAnswer = vbYes Then
'     ' Msgbox "You answered yes."
'     If i<4 Then
'       build = build&"1:"
'     Else
'       build = build&"1"
'     End If
' Else
'     ' Msgbox "You answered no."
'     If i<4 Then
'       build = build&"0:"
'     Else
'       build = build&"0"
'     End If
' End If
'
' Next
'
' tm = tim&":"&build
'
' MsgBox tm


' ' ft = tm&":"&weektime
'
' ssnd = DateDiff("d",today,snd)
'
' Dim weektime
' weektime = Array("0:0:0:1","0:0:1:0","0:0:1:1","0:1:0:0","0:1:0:1","0:1:1:0","0:1:1:1","1:0:0:0","1:0:0:1","1:0:1:0","1:0:1:1","1:1:0:0","1:1:0:1","1:1:1:0","1:1:1:1","0:0:0:0")
' Dim strHTML
' Dim IE


' strHTML = "<HTML>" & "<HEAD>"
' strHTML = strHTML & vbCrlf & "<TITLE>Help Box</TITLE>"
' strHTML = strHTML & vbCrlf & "<SCRIPT TYPE=""text/vbscript"">"
' strHTML = strHTML & vbCrlf & "Sub subOK()"
' strHTML = strHTML & vbCrlf & "    Msgbox ""test"""
' strHTML = strHTML & vbCrlf & "End Sub"
' strHTML = strHTML & vbCrlf & "</SCRIPT>"
' strHTML = strHTML & vbCrlf & "</HEAD>"
' strHTML = strHTML & vbCrlf & "<BODY>"
' strHTML = strHTML & vbCrlf & "Error occured. If you want"
' strHTML = strHTML & vbCrlf & "<BR> help, click the link below."
' strHTML = strHTML & vbCrlf & "<BR><A HREF=""http://www.help.com"">http://www.help.com</A>"
' strHTML = strHTML & vbCrlf & "<BR><CENTER><INPUT TYPE=""button"" name=""cmdOK"" Value=""OK"" onClick=""subOK""></CENTER>"
' strHTML = strHTML & vbCrlf & "</BODY></HTML>"
'
' 'MsgBox strHTML
' Set IE = WScript.CreateObject("InternetExplorer.Application")
' IE.Navigate "about:blank"
' IE.AddressBar = 0
' IE.menubar = 0
' IE.ToolBar = 0
' IE.StatusBar = 0
' IE.width = 400
' IE.height = 150
' IE.resizable = 0
' IE.visible = True

' ie.Document.Body.InnerHTML = strHTML

' MsgBox weektime(0)
' use = 1
' Set IExp = CreateObject("InternetExplorer.Application")
' Set WSHShell = WScript.CreateObject("WScript.Shell")
' url = "https://zoom.us/j/"&use
' IExp.Visible = False
' IExp.navigate url
'
' For Each w In CreateObject("Shell.Application").Windows
'     w.Quit()
' Next
