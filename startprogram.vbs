dim strProgram, x, ObjShell
strProgram = "WinWord.exe"

set ObjShell = CreateObject("WScript.Shell")

ObjShell.RUN strProgram

WScript.Sleep 3000

ObjShell.SendKeys("blablabla1")
ObjShell.SendKeys("{Enter}")
ObjShell.SendKeys("2. Zeile")

WScript.Sleep 3000

ObjShell.SendKeys("")

WScript.Sleep 3000

ObjShell.SendKeys("erstes Dokument")
ObjShell.SendKeys("{Enter}")

WScript.Sleep 3000

ObjShell.SendKeys("{Enter}")

WScript.Sleep 3000

ObjShell.SendKeys("%{F4}")
