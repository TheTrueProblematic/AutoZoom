today = Date
snd = "1/1/3000"
week = Weekday(Date)
hr = Hour(Time)
min = Minute(Time)
tm = hr&":"&min
nicetime = TimeValue(tm)

Dim weektime

If week = 2 Then
' Monday
weektime = Array("1:0:0:0:1","1:0:0:1:0","1:0:0:1:1","1:0:1:0:0","1:0:1:0:1","1:0:1:1:0","1:0:1:1:1","1:1:0:0:0","1:1:0:0:1","1:1:0:1:0","1:1:0:1:1","1:1:1:0:0","1:1:1:0:1","1:1:1:1:0","1:1:1:1:1","1:0:0:0:0")
ElseIf week = 3 Then
' Tuesday
weektime = Array("0:1:0:0:1","0:1:0:1:0","0:1:0:1:1","0:1:1:0:0","0:1:1:0:1","0:1:1:1:0","0:1:1:1:1","1:1:0:0:0","1:1:0:0:1","1:1:0:1:0","1:1:0:1:1","1:1:1:0:0","1:1:1:0:1","1:1:1:1:0","1:1:1:1:1","0:1:0:0:0")
ElseIf week = 4 Then
' Wednessday
weektime = Array("0:0:1:0:1","0:0:1:1:0","0:0:1:1:1","0:1:1:0:0","0:1:1:0:1","0:1:1:1:0","0:1:1:1:1","1:0:1:0:0","1:0:1:0:1","1:0:1:1:0","1:0:1:1:1","1:1:1:0:0","1:1:1:0:1","1:1:1:1:0","1:1:1:1:1","0:0:1:0:0")
ElseIf week = 5 Then
' Thursday
weektime = Array("0:0:0:1:1","0:0:1:1:0","0:0:1:1:1","0:1:0:1:0","0:1:0:1:1","0:1:1:1:0","0:1:1:1:1","1:0:0:1:0","1:0:0:1:1","1:0:1:1:0","1:0:1:1:1","1:1:0:1:0","1:1:0:1:1","1:1:1:1:0","1:1:1:1:1","0:0:0:1:0")
ElseIf week = 6 Then
' Friday
weektime = Array("0:0:0:1:1","0:0:1:0:1","0:0:1:1:1","0:1:0:0:1","0:1:0:1:1","0:1:1:0:1","0:1:1:1:1","1:0:0:0:1","1:0:0:1:1","1:0:1:0:1","1:0:1:1:1","1:1:0:0:1","1:1:0:1:1","1:1:1:0:1","1:1:1:1:1","0:0:0:0:1")
Else
End If



'msgbox nicetime





c = 0
c = c-1
Dim zooms = Array("","","","","","","","","","")
Dim times = Array("","","","","","","","","","")
' zoom0 = ""
' zoom1 = ""
' zoom3 = ""
' zoom2 = ""
' zoom4 = ""
' zoom6 = ""
' zoom5 = ""
' zoom7 = ""
' zoom8 = ""
' zoom9 = ""
' time0 = ""
' time1 = ""
' time2 = ""
' time3 = ""
' time4 = ""
' time5 = ""
' time6 = ""
' time7 = ""
' time8 = ""
' time9 = ""
nnd = ""
Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile("Vars.txt",1)
Dim strLine
do while not objFileToRead.AtEndOfStream
     strLine = objFileToRead.ReadLine()
     'Do something with the line
     x = strLine
     If c = 0 Then
     zooms(0) = x
     ElseIf c = 1 Then
     zooms(1) = x
     ElseIf c = 2 Then
     zooms(2) = x
     ElseIf c = 3 Then
     zooms(3) = x
     ElseIf c = 4 Then
     zooms(4) = x
     ElseIf c = 5 Then
     zooms(5) = x
     ElseIf c = 6 Then
     zooms(6) = x
     ElseIf c = 7 Then
     zooms(7) = x
     ElseIf c = 8 Then
     zooms(8) = x
     ElseIf c = 9 Then
     zooms(9) = x
     ElseIf c = 10 Then
     times(0) = x
     ElseIf c = 11 Then
     times(1) = x
     ElseIf c = 12 Then
     times(2) = x
     ElseIf c = 13 Then
     times(3) = x
     ElseIf c = 14 Then
     times(4) = x
     ElseIf c = 15 Then
     times(5) = x
     ElseIf c = 16 Then
     times(6) = x
     ElseIf c = 17 Then
     times(7) = x
     ElseIf c = 18 Then
     times(8) = x
     ElseIf c = 19 Then
     times(9) = x
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

For i = 0 to 15
wt = weektime(i)
ft = tm&":"&wt

If time(0) = ft Then
use = zoom(0)
ElseIf time(1) = ft Then
use = zoom(1)
ElseIf time(2) = ft Then
use = zoom(2)
ElseIf time(3) = ft Then
use = zoom(3)
ElseIf time(4) = ft Then
use = zoom(4)
ElseIf time(5) = ft Then
use = zoom(5)
ElseIf time(6) = ft Then
use = zoom(6)
ElseIf time(7) = ft Then
use = zoom(7)
ElseIf time(8) = ft Then
use = zoom(8)
ElseIf time(9) = ft Then
use = zoom(9)
Else
End If

Next

If use = "" Then

Else

Set IExp = CreateObject("InternetExplorer.Application")
Set WSHShell = WScript.CreateObject("WScript.Shell")
url = "https://zoom.us/j/"&use
IExp.Visible = False
IExp.navigate url

For Each w In CreateObject("Shell.Application").Windows
    w.Quit()
Next


End If


WScript.Sleep 60000
CreateObject("WScript.Shell").Run WScript.ScriptFullName
