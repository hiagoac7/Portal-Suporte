@echo off
chcp 65001 >nul
title Portal de Suporte TLM
:: Trabalha SEMPRE a partir da pasta deste arquivo (%~dp0). Antes o caminho
:: estava fixo no .bat; se a pasta mudasse de lugar ou de usuario, quebrava.
cd /d "%~dp0"

if not exist ".\venv\Scripts\pythonw.exe" (
    echo [ERRO] Nao encontrei .\venv\Scripts\pythonw.exe
    echo Verifique se a pasta venv existe.
    pause
    exit /b
)

echo Limpando bytes nulos do app.py ^(saneamento anti-corrupcao^)...
".\venv\Scripts\python.exe" "app\sanitize_app.py"

echo Aplicando a marca TLM no HTML do Streamlit ^(titulo e favicon^)...
".\venv\Scripts\python.exe" "app\marca_streamlit.py"

echo Iniciando o Portal de Suporte em segundo plano...

:: pythonw.exe roda SEM console, entao o site NAO depende desta janela.
start "" ".\venv\Scripts\pythonw.exe" -m streamlit run "app\app.py" --server.port 8501 --server.address 0.0.0.0 --server.headless true

echo.
echo [OK] Portal no ar: http://localhost:8501
echo.
echo  - Pode FECHAR esta janela; o site continua rodando.
echo  - Para PARAR: Gerenciador de Tarefas ^(Ctrl+Shift+Esc^),
echo    aba Detalhes, finalize o processo "pythonw.exe".
echo.
echo  - Acesso de OUTROS PCs na mesma rede, use o(s) IP(s) abaixo:
for /f "tokens=2 delims=:" %%A in ('ipconfig ^| findstr /c:"IPv4"') do (
    echo      http://%%A:8501
)
echo    (se nao abrir de outro PC, libere a porta 8501 no Firewall do Windows)
echo.
pause
