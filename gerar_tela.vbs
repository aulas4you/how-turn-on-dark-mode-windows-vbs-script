Dim WshShell, BtnCode
Set WshShell = WScript.CreateObject("WScript.Shell")

Dim objShell
Set objShell = Wscript.CreateObject("WScript.Shell")


Dim objShell2
Set objShell2 = Wscript.CreateObject("WScript.Shell")


BtnCode = WshShell.Popup("Quer ligar o modo escuro", 7, "Clique para alterar as configurações", 4 + 32)
WScript.Echo "botao"&BtnCode
Select Case BtnCode
   case 6      WScript.Echo "Dark thema enable"
	objShell.Run "ligar_dark.vbs" 
   case 7      WScript.Echo "dark thema disable"
       objShell.Run "desligar_dark.vbs" 
	
   case -1     WScript.Echo "existe alguem ai?"
End Select


