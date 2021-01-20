today = Date
snd = "1/1/3000"
week = Weekday(Date)
hr = Hour(Time)
min = Minute(Time)
tm = hr&":"&min
nicetime = TimeValue(tm)


' ft = tm&":"&weektime

ssnd = DateDiff("d",today,snd)

Dim weektime
weektime = Array("0:0:0:1","0:0:1:0","0:0:1:1","0:1:0:0","0:1:0:1","0:1:1:0","0:1:1:1","1:0:0:0","1:0:0:1","1:0:1:0","1:0:1:1","1:1:0:0","1:1:0:1","1:1:1:0","1:1:1:1","0:0:0:0")
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

MsgBox weektime(0)
