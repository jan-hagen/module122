option explicit

dim ausgabe
dim heute
set ausgabe = WScript.CreateObject("Wscript.Shell")
heute = Date
ausgabe.popup "Heute ist der " & heute, 0,"today"
