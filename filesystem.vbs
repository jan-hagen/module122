option explicit

dim FS
dim MyWinDir
dim ausgabe

set FS = CreateObject("Scripting.FileSystemObject")
Set MyWinDir = FS.GetSpecialFolder(0)

ausgabe.popup "WINIDIR " & MyWinDir, 0, "WINDIR"
