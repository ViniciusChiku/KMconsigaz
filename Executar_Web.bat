@echo off
title Consigaz - Auditoria de KM Web
cd /d "C:\Users\vchiku\Desktop\automacoes\Fechamento"

netstat -ano | findstr :8501 | findstr LISTENING >nul
if %errorlevel% equ 0 (
    echo O sistema ja esta em execucao! Abrindo no navegador...
    start http://localhost:8501
    exit /b
)

echo Iniciando o sistema web local...
start http://localhost:8501
python -m streamlit run app_web_km.py --server.port 8501
if %errorlevel% neq 0 (
    echo Ocorreu um erro ao executar.
    pause
)
