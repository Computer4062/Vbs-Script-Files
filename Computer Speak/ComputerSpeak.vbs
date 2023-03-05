Dim userinput 
userinput = inputbox ("Type bellow to here your computer speak:")
set sapi = wscript.createobject ("SAPI.Spvoice")
Sapi.speak userinput
if userinput = "Mihan" Then 
x = msgbox ("Hy Mihan",2+16,"Hy Mihan")
End if