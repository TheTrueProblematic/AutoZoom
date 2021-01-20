n = InputBox("How many Zoom classes?","ZoomLauncher","Type Here")

If n="" Then
  Wscript.Quit
Else
nn = n-1
End If

'Loop for Class Code and Time prompts!
For i = 0 to nn
ii = i+1


'Class Code Prompt
cl = InputBox("Zoom code for Class "&ii,"ZoomLauncher","Type Here")

If cl="" Then
  Wscript.Quit
Else
End If

Set shell = CreateObject("WScript.Shell")
x = "Files\Vars.txt "
y = "zoom"&i&" "
z = cl
shell.Run "Files\Replace.vbs " & x & y & z


'Time Prompts
tim = InputBox("What time is class "&ii,"ZoomLauncher","Use Military Time (Example 13:45)")
Dim dow
dow = Array("Monday","Tuesday","Wednessday","Thursday","Friday")
build = ""
For o = 0 to 4
tday = dow(o)
intAnswer = _
    MsgBox("Do you have class "&ii&" on "&tday&"?", _
        vbYesNo, "ZoomLauncher")
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

tm = tim&":"&build

If tm="" Then
  Wscript.Quit
Else
End If

Set shell = CreateObject("WScript.Shell")
x = "Files\Vars.txt "
y = "time"&i&" "
z = tm
shell.Run "Files\Replace.vbs " & x & y & z

'End of loop
Next

snd = InputBox("What date does the semester end?","ZoomLauncher","month/day/year")
If snd="" Then
  Wscript.Quit
Else
End If

Set shell = CreateObject("WScript.Shell")
x = "Files\Vars.txt "
y = "end "
z = snd
shell.Run "Files\Replace.vbs " & x & y & z

Set shell = CreateObject("WScript.Shell")
x = "Files\Vars.txt "
y = "false "
z = "true"
shell.Run "Files\Replace.vbs " & x & y & z

Set shell = CreateObject("WScript.Shell")
'shell.CurrentDirectory = "C:\Users\js\Desktop\createIndex"
shell.Run "Files\Config.bat"
