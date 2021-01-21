version = "AutoZoom 0.2.0"


n = InputBox("How many Zoom classes?",version,"Type Here")

If n="" Then
  Wscript.Quit
Else
nn = n-1
End If

'Loop for Class Code and Time prompts!
For i = 0 to nn
ii = i+1


'Class Code Prompt
cl = InputBox("Zoom code for Class "&ii,version,"Type Here")

If cl="" Then
  Wscript.Quit
Else
End If

Set shell = CreateObject("WScript.Shell")
x = "AutoZoom\Vars.txt "
y = "zoom"&i&" "
z = cl
shell.Run "AutoZoom\Replace.vbs " & x & y & z


'Time Prompts
tim = InputBox("What time is class "&ii,version,"Use Military Time (Example 13:45)")
Dim dow
dow = Array("Monday","Tuesday","Wednessday","Thursday","Friday")
build = ""
For o = 0 to 4
tday = dow(o)
intAnswer = _
    MsgBox("Do you have class "&ii&" on "&tday&"?", _
        vbYesNo, version)
If intAnswer = vbYes Then
    ' Msgbox "You answered yes."
    If o<4 Then
      build = build&"1:"
    Else
      build = build&"1"
    End If
Else
    ' Msgbox "You answered no."
    If o<4 Then
      build = build&"0:"
    Else
      build = build&"0"
    End If
End If
Next

inAnswer = _
    MsgBox("Does class "&ii&" have a passcode?", _
        vbYesNo, version)
If inAnswer = vbYes Then
    ' Msgbox "You answered yes."
    cod = InputBox("What is the passcode for class "&ii&"?",version,"Type Here")

    Set shell = CreateObject("WScript.Shell")
    x = "AutoZoom\Vars.txt "
    y = "code"&i&" "
    z = cod
    shell.Run "AutoZoom\Replace.vbs " & x & y & z

Else
    ' Msgbox "You answered no."

End If


tm = tim&":"&build

If tm="" Then
  Wscript.Quit
Else
End If

Set shell = CreateObject("WScript.Shell")
x = "AutoZoom\Vars.txt "
y = "time"&i&" "
z = tm
shell.Run "AutoZoom\Replace.vbs " & x & y & z

'End of loop
Next

snd = InputBox("What date does the semester end?",version,"month/day/year")
If snd="" Then
  Wscript.Quit
Else
End If

Set shell = CreateObject("WScript.Shell")
x = "AutoZoom\Vars.txt "
y = "end "
z = snd
shell.Run "AutoZoom\Replace.vbs " & x & y & z

Set shell = CreateObject("WScript.Shell")
x = "AutoZoom\Vars.txt "
y = "false "
z = "true"
shell.Run "AutoZoom\Replace.vbs " & x & y & z

Set shell = CreateObject("WScript.Shell")
'shell.CurrentDirectory = "C:\Users\js\Desktop\createIndex"
shell.Run "AutoZoom\Config.bat"
