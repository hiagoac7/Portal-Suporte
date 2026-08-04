@echo off
chcp 65001 >nul
title Configurar Portal para iniciar com o Windows
echo ============================================================
echo   Configurando o Portal para iniciar junto com o Windows
echo ============================================================
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0criar_atalho_inicio.ps1"
