m = 10000

x = MsgBox("Herzlich Willkommen zu Casino.vbs!" & vbCrLf & "Ihr aktuelles Geld: 10 000 $",1+48,"Virtual Casino")
If x=1 then
	Do
		y = MsgBox("'OK' druecken, um zu drehen" & vbCrLf & "Kostet 100$",1+64,"Money:" & m & "$")
		If y = 1 then
			m = m-100
			a = Int((RND*10)+1)
			b = Int((RND*10)+1)
			c = Int((RND*10)+1)
			MsgBox a & "|" & b & "|" & c,0,"Einarmiger Bandit"
			If a = b AND b = c then
					m = m+600
					MsgBox "Gratuliere! +600$!" & vbCrLf & m & "$"
					v = MsgBox("Nochmal?",4)
					If v = 6 then
						f = 100
					else
						f = 1
					end if
				else
					g = MsgBox("Leider nicht! :-(" & vbCrLf & "Nochmal?" & vbCrLf & m & "$",4)
					If g = 6 then
						If m=0 then
							MsgBox "Du hast leider kein Geld mehr!!"
							f = 1
						else
							f = 100
						end if
					else
						f = 1
					end if
			end if
		else
			f = 1
		end if
	loop until f=1
else
	f = 1
end if
