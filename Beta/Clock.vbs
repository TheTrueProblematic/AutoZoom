today = Date
snd = "1/1/3000"
week = Weekday(Date)
hr = Hour(Time)
min = Minute(Time)
tm = hr&":"&min
nicetime = TimeValue(tm)
ft = tm&":"&week



'msgbox nicetime



c = 0
c = c-1
zoom0 = ""
zoom1 = ""
zoom3 = ""
zoom2 = ""
zoom4 = ""
zoom6 = ""
zoom5 = ""
zoom7 = ""
zoom8 = ""
zoom9 = ""
time0 = ""
time1 = ""
time2 = ""
time3 = ""
time4 = ""
time5 = ""
time6 = ""
time7 = ""
time8 = ""
time9 = ""
nnd = ""
Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile("Vars.txt",1)
Dim strLine
do while not objFileToRead.AtEndOfStream
     strLine = objFileToRead.ReadLine()
     'Do something with the line
     x = strLine
     If c = 0 Then
     zoom0 = x
     ElseIf c = 1 Then
     zoom1 = x
     ElseIf c = 2 Then
     zoom2 = x
     ElseIf c = 3 Then
     zoom3 = x
     ElseIf c = 4 Then
     zoom4 = x
     ElseIf c = 5 Then
     zoom5 = x
     ElseIf c = 6 Then
     zoom6 = x
     ElseIf c = 7 Then
     zoom7 = x
     ElseIf c = 8 Then
     zoom8 = x
     ElseIf c = 9 Then
     zoom9 = x
     ElseIf c = 10 Then
     time0 = x
     ElseIf c = 11 Then
     time1 = x
     ElseIf c = 12 Then
     time2 = x
     ElseIf c = 13 Then
     time3 = x
     ElseIf c = 14 Then
     time4 = x
     ElseIf c = 15 Then
     time5 = x
     ElseIf c = 16 Then
     time6 = x
     ElseIf c = 17 Then
     time7 = x
     ElseIf c = 18 Then
     time8 = x
     ElseIf c = 19 Then
     time9 = x
     ElseIf c = 20 Then
     nnd = x
     Else
     End If

     c = c+1
loop
objFileToRead.Close
Set objFileToRead = Nothing


ssnd = DateDiff("d",today,nnd)
If ssnd<0 Then
Wscript.Quit
Else
End If


use = ""

If time0 = ft Then
use = zoom0
ElseIf time1 = ft Then
use = zoom1
ElseIf time2 = ft Then
use = zoom2
ElseIf time3 = ft Then
use = zoom3
ElseIf time4 = ft Then
use = zoom4
ElseIf time5 = ft Then
use = zoom5
ElseIf time6 = ft Then
use = zoom6
ElseIf time7 = ft Then
use = zoom7
ElseIf time8 = ft Then
use = zoom8
ElseIf time9 = ft Then
use = zoom9
Else
End If

If use = "" Then

Else

Set IExp = CreateObject("InternetExplorer.Application")
Set WSHShell = WScript.CreateObject("WScript.Shell")
IExp.Visible = False
IExp.navigate "https://zoom.us/j/"+use

End If


WScript.Sleep 60000
CreateObject("WScript.Shell").Run WScript.ScriptFullName
