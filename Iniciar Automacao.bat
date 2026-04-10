@echo off
title Automacao de Baixa NF - Dealer.net
color 0A

echo ============================================================
echo    AUTOMACAO DE BAIXA DE NF - DEALER.NET
echo ============================================================
echo.
echo  Iniciando servidor...
echo.

cd /d "%~dp0"
set PATH=C:\Users\igorferreira\AppData\Local\Programs\Python\Python312;C:\Users\igorferreira\AppData\Local\Programs\Python\Python312\Scripts;%PATH%

REM Inicia o servidor em background
start /B "" python servidor.py

REM Espera o servidor subir
timeout /t 3 /nobreak >nul

REM Abre o navegador padrao
echo  Abrindo painel de controle no navegador...
start http://localhost:5000

echo.
echo ============================================================
echo  Servidor rodando em: http://localhost:5000
echo.
echo  NAO FECHE ESTA JANELA enquanto usar a automacao!
echo  Para encerrar, feche esta janela ou pressione CTRL+C
echo ============================================================
echo.

REM Mantem a janela aberta
pause >nul
