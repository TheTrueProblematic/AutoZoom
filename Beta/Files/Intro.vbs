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
x = "Vars.txt "
y = "zoom"&i&" "
z = cl
shell.Run "Replace.vbs " & x & y & z


'Time Prompts
tm = ""               'Prompt User For Time and Day of Week

If tm="" Then
  Wscript.Quit
Else
End If

Set shell = CreateObject("WScript.Shell")
x = "Vars.txt "
y = "time"&i&" "
z = tm
shell.Run "Replace.vbs " & x & y & z

'End of loop
Next

snd = InputBox("What date does the semester end?","ZoomLauncher","month/day/year")
If snd="" Then
  Wscript.Quit
Else
End If

Set shell = CreateObject("WScript.Shell")
x = "Vars.txt "
y = "end "
z = snd
shell.Run "Replace.vbs " & x & y & z

Set shell = CreateObject("WScript.Shell")
x = "Vars.txt "
y = "false "
z = "true"
shell.Run "Replace.vbs " & x & y & z

Set shell = CreateObject("WScript.Shell")
'shell.CurrentDirectory = "C:\Users\js\Desktop\createIndex"
shell.Run "Config.bat"
