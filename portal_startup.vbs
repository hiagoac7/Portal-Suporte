' ============================================================
'  Portal de Suporte - inicializacao automatica (Windows)
'  Inicia o Streamlit em segundo plano, OCULTO (sem janela).
'  Coloque este arquivo (ou um atalho dele) na pasta:
'      shell:startup   (Win+R -> digite shell:startup -> Enter)
' ============================================================

Dim sh, base, cmd
Set sh = CreateObject("WScript.Shell")

base = "C:\Users\hiago TI\PortalSuporte"
sh.CurrentDirectory = base

cmd = """" & base & "\venv\Scripts\pythonw.exe"" -m streamlit run """ & base & "\app.py"" --server.port 8501 --server.address 0.0.0.0 --server.headless true"

' 0 = janela oculta ; False = nao esperar (o script sai imediatamente)
sh.Run cmd, 0, False

Set sh = Nothing
