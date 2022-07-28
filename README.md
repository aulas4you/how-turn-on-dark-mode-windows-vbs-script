# how-turn-on-dark-mode-windows-vbs-script
simple vbs script to windows 10 to turn on and turn off DARK MODE thema of colors.
You can erase the "echo" parts to avoid printf popups

gerar_tela.vbs   start a pop up option to turn on or turn off dark mode
desligar_dark.vbs    change the registry of windows to turn off dark mode
ligar_dark.vbs   change the registry of windows to turn on dark mode 


How this work?
Dark thema off
HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize  SystemUsesLightTheme D_WORD 1
HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize AppsUseLightTheme"   D_WORD 1 

Dark thema on
HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize  SystemUsesLightTheme D_WORD 0
HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize AppsUseLightTheme"   D_WORD 0 


gerar_tela.vbs generates the popup and select the option
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


 the vbs ligar_dark change this with
 Set WshShell = WScript.CreateObject("WScript.Shell")
  Set WshShell2 = WScript.CreateObject("WScript.Shell")
 WshShell2. RegWrite"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme", 1, "REG_DWORD"
 WshShell. RegWrite"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\SystemUsesLightTheme", 1, "REG_DWORD"
 
 
 the vbs desligarligar_dark change this with
 
 Set WshShell = WScript.CreateObject("WScript.Shell")
Set WshShell2 = WScript.CreateObject("WScript.Shell") 
 WshShell2. RegWrite"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme", 0, "REG_DWORD"
 WshShell. RegWrite"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\SystemUsesLightTheme", 0, "REG_DWORD"
 
 
 WScript.Echo can be deleted ,is only for debug
 
