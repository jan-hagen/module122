'oShell definieren'
Dim oShell : Set oShell = CreateObject("WScript.Shell")

'Startet Wordpad'
oShell.Run "wordpad"

'Warten bis Wordpad gestartet'
WScript.Sleep 3000

' Beendet Wordpad'
oShell.Run "taskkill /im wordpad.exe", , True