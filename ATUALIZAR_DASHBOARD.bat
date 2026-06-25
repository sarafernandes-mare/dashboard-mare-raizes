@echo off
chcp 65001 >nul
title Mare Educacao - Atualizar Base e Dashboard
cd /d "%~dp0"

echo ============================================================
echo   MARE EDUCACAO - Atualizacao da Base + Dashboard
echo ============================================================
echo.
echo  Passo 1: Exporte o relatorio de vendas da Hotmart
echo           (com Atrasados incluidos) para a pasta Downloads.
echo           O nome precisa conter a palavra "venda".
echo.
echo  Passo 2: Feche o arquivo base_mare_parcelas_projecao.xlsx
echo           se ele estiver aberto no Excel.
echo.
echo  Pressione uma tecla para gerar tudo...
pause >nul
echo.
echo  Processando... (pode levar ~1 minuto)
echo.

py "gerar_base_mare.py"
set ERR=%ERRORLEVEL%

echo.
if %ERR%==0 (
  echo ============================================================
  echo   CONCLUIDO COM SUCESSO!
  echo ============================================================
  echo   - Planilha:  base_mare_parcelas_projecao.xlsx
  echo   - Dashboard: dashboard_mare.html
  echo.
  echo   Abrindo o dashboard...
  start "" "dashboard_mare.html"
) else (
  echo ============================================================
  echo   OCORREU UM ERRO ^(codigo %ERR%^)
  echo ============================================================
  echo   Verifique se:
  echo   - O arquivo exportado da Hotmart esta na pasta Downloads
  echo     e o nome contem "venda"
  echo   - O base_mare_parcelas_projecao.xlsx nao esta aberto
  echo   - O Python esta instalado ^(comando: py --version^)
)
echo.
echo  Pressione uma tecla para fechar.
pause >nul
