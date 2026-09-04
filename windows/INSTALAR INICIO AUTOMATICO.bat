@echo off
chcp 65001 >nul
title Instalar inicio automatico do Portal de Suporte
echo ============================================================
echo   Portal de Suporte - iniciar junto com o Windows
echo ============================================================
echo.
echo Metodo: atalho na pasta de Inicializacao do seu usuario.
echo Nao precisa de administrador. Roda oculto, ~20s apos o login.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0criar_atalho_inicio.ps1"
