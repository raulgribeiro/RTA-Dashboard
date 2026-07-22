@echo off
chcp 65001 >nul
setlocal

REM ============================================================
REM  Atualiza o dashboard RTA (dados novos da planilha) e ja
REM  publica a atualizacao no GitHub / GitHub Pages.
REM
REM  Este arquivo deve ficar na MESMA PASTA do script
REM  atualizar_dashboard_rta.py (ex: App RTA).
REM ============================================================

REM Garante que estamos rodando a partir da pasta onde este .bat esta salvo,
REM nao importa de onde ele foi clicado.
cd /d "%~dp0"

echo ============================================================
echo  1/2 - Atualizando o HTML com os dados das planilhas...
echo ============================================================
python atualizar_dashboard_rta.py
if errorlevel 1 (
    echo.
    echo ============================================================
    echo  ERRO: o script de atualizacao falhou. Nada foi enviado ao
    echo  GitHub. Revise a mensagem de erro acima.
    echo ============================================================
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  2/2 - Enviando a atualizacao para o GitHub...
echo ============================================================

git add .
git commit -m "Atualizacao automatica dos dados RTA - %date% %time%"
if errorlevel 1 (
    echo.
    echo Nada novo para enviar ao GitHub ^(os dados nao mudaram desde a ultima vez^).
    echo.
    pause
    exit /b 0
)

git push
if errorlevel 1 (
    echo.
    echo ============================================================
    echo  ERRO ao enviar para o GitHub. O HTML local foi atualizado,
    echo  mas a versao publicada no GitHub Pages AINDA NAO reflete essa
    echo  atualizacao. Revise a mensagem de erro acima ^(rede, login,
    echo  conflito de arquivos, etc^) e rode este .bat de novo.
    echo ============================================================
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Tudo certo! Dados atualizados e enviados ao GitHub.
echo  O site no GitHub Pages deve refletir a mudanca em 1-2 minutos.
echo ============================================================
pause
