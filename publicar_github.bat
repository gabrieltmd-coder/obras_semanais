@echo off
echo ============================================
echo   TMS - Publicar no GitHub
echo ============================================
echo.
echo Repositorio: https://github.com/gabrieltmd/obras-semanais
echo.
set "GIT=C:\Program Files\Git\bin\git.exe"
cd /d "%~dp0"

"%GIT%" add .
"%GIT%" status --short

echo.
set /p MSG="Mensagem do commit (Enter para usar 'atualizacao'): "
if "%MSG%"=="" set MSG=atualizacao

"%GIT%" commit -m "%MSG%"

echo.
echo Enviando para GitHub...
echo Quando solicitado, use:
echo   Usuario: gabrieltmd
echo   Senha:   seu Personal Access Token (PAT)
echo.
"%GIT%" push -u origin main

echo.
if %ERRORLEVEL%==0 (
    echo [OK] Publicado com sucesso!
    echo Acesse: https://github.com/gabrieltmd/obras-semanais
) else (
    echo [ERRO] Falha no push. Verifique o PAT e tente novamente.
)
echo.
pause
