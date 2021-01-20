c = 0
introComplete = False

Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile("Files\Vars.txt",1)
Dim strLine
do while not objFileToRead.AtEndOfStream
     strLine = objFileToRead.ReadLine()
     'Do something with the line
     x = strLine
     If c = 0 Then
      If x = "false" Then
        introComplete = False
      Else
        introComplete = True
      End If

     Else
     End If

     c = c+1
loop
objFileToRead.Close
Set objFileToRead = Nothing




If introComplete = True Then
  CreateObject("WScript.Shell").Run("Files\Clock.vbs")
Else
  CreateObject("WScript.Shell").Run("Files\Intro.vbs")
End If
