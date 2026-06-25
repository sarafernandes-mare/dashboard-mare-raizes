@echo off
chcp 65001 >nul
REM Versao SILENCIOSA para o Agendador de Tarefas do Windows (sem pausas).
REM Roda o script, regenera xlsx + dashboard. Grava log em _log_dashboard.txt
cd /d "%~dp0"
echo [%date% %time%] Iniciando atualizacao >> _log_dashboard.txt
py "gerar_base_mare.py" >> _log_dashboard.txt 2>&1
echo [%date% %time%] Finalizado (codigo %ERRORLEVEL%) >> _log_dashboard.txt
echo. >> _log_dashboard.txt
