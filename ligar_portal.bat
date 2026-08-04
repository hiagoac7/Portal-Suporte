@echo off
cd /d "C:\Users\hiago TI\PortalSuporte"

if not exist ".\venv\Scripts\pythonw.exe" (
    echo [ERRO] Nao encontrei .\venv\Scripts\pythonw.exe
    echo Verifique se a pasta venv existe.
    pause
    exit /b
)

echo Limpando bytes nulos do app.py (saneamento anti-corrupcao)...
".\venv\Scripts\python.exe" "sanitize_app.py"

echo Iniciando o Portal de Suporte em segundo plano...

:: pythonw.exe roda SEM console, entao o site NAO depende desta janela.
:: Fechar esta janela nao derruba mais o site.
:: --server.address 0.0.0.0 garante que o site aceite conexoes de outros PCs da rede.
start "" ".\venv\Scripts\pythonw.exe" -m streamlit run app.py --server.port 8501 --server.address 0.0.0.0 --server.headless true

echo.
echo [OK] Portal no ar: http://localhost:8501
echo.
echo  - Pode FECHAR esta janela; o site continua rodando.
echo  - Para PARAR o portal: Gerenciador de Tarefas (Ctrl+Shift+Esc),
echo    aba Detalhes, finalize o processo "pythonw.exe".
echo.
echo  - Acesso de OUTROS PCs na mesma rede (Wi-Fi/cabo), use o(s) IP(s) abaixo:
for /f "tokens=2 delims=:" %%A in ('ipconfig ^| findstr /c:"IPv4"') do (
    echo      http://%%A:8501
)
echo    (se nao abrir de outro PC, libere a porta 8501 no Firewall do Windows)
echo.
pause