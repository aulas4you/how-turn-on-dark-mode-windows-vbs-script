
Dim WshShell, bKey
Dim WshShell2


Set WshShell = WScript.CreateObject("WScript.Shell")

Set WshShell2 = WScript.CreateObject("WScript.Shell")



WScript.Echo "SystemUsesLightTheme" & WshShell.RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\SystemUsesLightTheme") 

WScript.Echo "AppsUseLightTheme" & WshShell2.RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme")


WshShell. RegWrite"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\SystemUsesLightTheme", 0, "REG_DWORD"

WshShell2. RegWrite"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme", 0, "REG_DWORD"


WScript.Echo "SystemUsesLightTheme" & WshShell.RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\SystemUsesLightTheme") 

WScript.Echo "AppsUseLightTheme" & WshShell2.RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme")
