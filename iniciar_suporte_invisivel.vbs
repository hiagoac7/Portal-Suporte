' ============================================================
'  Inicia o Portal Suporte de forma invisivel ao ligar o PC.
'  Espera o sistema ficar pronto antes de rodar o portal.
' ============================================================

Dim WshShell, fso, batPath, esperaSegundos

batPath = "C:\Users\hiago TI\PortalSuporte\ligar_portal.bat"
esperaSegundos = 30   ' aumente para 60 se o PC for mais lento

Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Espera o sistema terminar de inicializar
WScript.Sleep esperaSegundos * 1000

' So executa se o .bat realmente existir
If fso.FileExists(batPath) Then
    WshShell.Run Chr(34) & batPath & Chr(34), 0, False
End If

Set fso = Nothing
Set WshShell = Nothing