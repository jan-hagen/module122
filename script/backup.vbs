Dim Fso
Dim Directory
Dim Modified
Dim Files
Dim MyDate
Dim OutputFile
MyDate = Replace(Date, "/", "-")

OutputFile = "backup-" & mydate & ".rar"
Set objShell = WScript.CreateObject("WScript.Shell")
objShell.Run "C:\Windows\WinRAR.exe a C:\backups\" & OutputFile & "C:\Users\minos\repos\"

Set Fso = CreateObject("Scripting.FileSystemObject")
Set Directory = Fso.GetFolder("C:\backups")
Set Files = Directory.Files
For Each Modified in Files
    If DateDiff("D", Modified.DateLastModified, Now) > 15 Then Modified.Delete
Next
