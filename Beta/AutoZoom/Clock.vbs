today = Date
snd = "1/1/3000"
week = Weekday(Date)
hr = Hour(Time)
min = Minute(Time)
tm = hr&":"&min
nicetime = TimeValue(tm)
term = ""

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



neg = 0
neg = neg-1
c = neg
Dim zooms
zooms = Array("","","","","","","","","","")
Dim times
times = Array("","","","","","","","","","")
Dim codes
codes = Array("","","","","","","","","","")
Dim endtims
endtims = Array("","","","","","","","","","")

nnd = ""
Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile("AutoZoom\Vars.txt",1)
Dim strLine
do while not objFileToRead.AtEndOfStream
     strLine = objFileToRead.ReadLine()
     'Do something with the line
     x = strLine
     If c = neg Then
     term = x
     Else
     End If

     If term = "false" Then
     Wscript.Quit
     Else
     End If

     If c>neg AND c<10 Then
     zooms(c) = x
     ElseIf c=20 Then
     nnd = x
     ElseIf c<21 Then
     times(c) = x
     ElseIf c<31 Then
     codes(c) = x
     ElseIf c<41 Then
     endtims(c) = x
     Else
     End If

     ' If c = 0 Then
     ' zooms(0) = x
     ' ElseIf c = 1 Then
     ' zooms(1) = x
     ' ElseIf c = 2 Then
     ' zooms(2) = x
     ' ElseIf c = 3 Then
     ' zooms(3) = x
     ' ElseIf c = 4 Then
     ' zooms(4) = x
     ' ElseIf c = 5 Then
     ' zooms(5) = x
     ' ElseIf c = 6 Then
     ' zooms(6) = x
     ' ElseIf c = 7 Then
     ' zooms(7) = x
     ' ElseIf c = 8 Then
     ' zooms(8) = x
     ' ElseIf c = 9 Then
     ' zooms(9) = x
     ' ElseIf c = 10 Then
     ' times(0) = x
     ' ElseIf c = 11 Then
     ' times(1) = x
     ' ElseIf c = 12 Then
     ' times(2) = x
     ' ElseIf c = 13 Then
     ' times(3) = x
     ' ElseIf c = 14 Then
     ' times(4) = x
     ' ElseIf c = 15 Then
     ' times(5) = x
     ' ElseIf c = 16 Then
     ' times(6) = x
     ' ElseIf c = 17 Then
     ' times(7) = x
     ' ElseIf c = 18 Then
     ' times(8) = x
     ' ElseIf c = 19 Then
     ' times(9) = x
     ' ElseIf c = 20 Then
     ' nnd = x
     ' ElseIf c = 21 Then
     ' codes(0) = x
     ' ElseIf c = 22 Then
     ' codes(1) = x
     ' ElseIf c = 23 Then
     ' codes(2) = x
     ' ElseIf c = 24 Then
     ' codes(3) = x
     ' ElseIf c = 25 Then
     ' codes(4) = x
     ' ElseIf c = 26 Then
     ' codes(5) = x
     ' ElseIf c = 27 Then
     ' codes(6) = x
     ' ElseIf c = 28 Then
     ' codes(7) = x
     ' ElseIf c = 29 Then
     ' codes(8) = x
     ' ElseIf c = 30 Then
     ' codes(9) = x
     ' Else
     ' End If

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
code = ""
endtim = ""

For i = 0 to 9
wt = weektime(i)
ft = tm&":"&wt

If endtims(i) = tm Then

Else
End If

If times(i) = ft Then
code = code(i)
use = zooms(i)
Else
End If


' If times(0) = ft Then
' code = code(0)
' use = zooms(0)
' ElseIf times(1) = ft Then
' code = code(1)
' use = zooms(1)
' ElseIf times(2) = ft Then
' code = code(2)
' use = zooms(2)
' ElseIf times(3) = ft Then
' code = code(3)
' use = zooms(3)
' ElseIf times(4) = ft Then
' code = code(4)
' use = zooms(4)
' ElseIf times(5) = ft Then
' code = code(5)
' use = zooms(5)
' ElseIf times(6) = ft Then
' code = code(6)
' use = zooms(6)
' ElseIf times(7) = ft Then
' code = code(7)
' use = zooms(7)
' ElseIf times(8) = ft Then
' code = code(8)
' use = zooms(8)
' ElseIf times(9) = ft Then
' code = code(9)
' use = zooms(9)
' Else
' End If
Next

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


WScript.Sleep 30000
CreateObject("WScript.Shell").Run WScript.ScriptFullName
