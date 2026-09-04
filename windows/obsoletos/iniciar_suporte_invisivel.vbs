Dim WshShell, fso, batPath, esperaSegundos

batPath = "C:\Users\hiago TI\PortalSuporte\ligar_portal.bat"
esperaSegundos = 30

Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

WScript.Sleep esperaSegundos * 1000

If fso.FileExists(batPath) Then
    WshShell.Run Chr(34) & batPath & Chr(34), 0, False
End If

Set fso = Nothing
Set WshShell = Nothing