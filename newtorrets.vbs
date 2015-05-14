do

Dim objShell, intCount, intRnd, intBase, intMax, intMin, intSleep
Set objShell = CreateObject("WScript.Shell")

randomize

intCount = 1
intBase = 5
intMax = 10
intMin = 3

WScript.Sleep 10000

Do While intCount < 3
intSleep = ((intBase)+Rnd+(Rnd*0.1))
Dim oPlayer
Set oPlayer = CreateObject("WMPlayer.OCX")

' Play audio
oPlayer.URL = "C:\Windows\Media\notify.wav"
oPlayer.controls.play 
While oPlayer.playState <> 1 ' 1 = Stopped
  WScript.Sleep 100
Wend

' Release the audio file
oPlayer.close
WScript.Sleep (CInt(intSleep)) 
intCount = intCount + 1
Loop

WScript.Quit
Loop