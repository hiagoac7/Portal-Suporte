Dim sh, base, cmd
Set sh = CreateObject("WScript.Shell")

base = "C:\Users\hiago TI\PortalSuporte"
sh.CurrentDirectory = base

cmd = """" & base & "\venv\Scripts\pythonw.exe"" -m streamlit run """ & base & "\app.py"" --server.port 8501 --server.address 0.0.0.0 --server.headless true"

sh.Run cmd, 0, False

Set sh = Nothing
