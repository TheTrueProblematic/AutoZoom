n = InputBox("How many Zoom classes?","ZoomLauncher","Type Here")

If n="" Then
  Wscript.Quit
Else
nn = n-1
End If

For i = 0 to nn

ii = i+1
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
