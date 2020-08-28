dim objWord, objDocument

set ObjWord = CreateObject("Word.Application")
objWord.Visible = True
Set objDocument = objWord.Documents.Open("C:\Users\minos\repos\module122\makros\Geschaeftsbrief.dotm")


'WScript.Sleep 3000

'ObjShell.SendKeys("blablabla1")


