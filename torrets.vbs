do
Dim oPlayer
Set oPlayer = CreateObject("WMPlayer.OCX")

Set WshShell = CreateObject("WScript.Shell")
WshShell.SendKeys(chr(175))

' Play audio
oPlayer.URL = "C:\Windows\Media\notify.wav"
oPlayer.controls.play 
While oPlayer.playState <> 1 ' 1 = Stopped
  WScript.Sleep 100
Wend

' Release the audio file
oPlayer.close
loop