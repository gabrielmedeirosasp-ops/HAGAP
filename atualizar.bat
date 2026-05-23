@echo off
chcp 65001 > nul
echo ============================================================
echo   HAGAP — Atualizador de projetos
echo ============================================================

:: ── CONFIGURE AQUI ──────────────────────────────────────────
set REPO=C:\Users\user\Desktop\PDF AUTOMATICO\HAGAP_WEB
set URL_SITE=https://seu-app.onrender.com
:: ────────────────────────────────────────────────────────────

cd /d "%REPO%"

echo.
echo [1/5] Baixando db.json atual do servidor (preserva edicoes feitas no site)...
curl -s "%URL_SITE%/api/dados" -o db_servidor.json
if %errorlevel% neq 0 (
    echo     AVISO: nao foi possivel baixar do servidor. Usando db.json local.
) else (
    echo     OK - db_servidor.json baixado
    copy /Y db_servidor.json db.json > nul
    echo     OK - db.json atualizado com dados do servidor
    del db_servidor.json
)

echo.
echo [2/5] Rodando gerador (merge com os novos PDFs)...
python gerador.py
if %errorlevel% neq 0 (
    echo     ERRO no gerador! Abortando.
    pause
    exit /b 1
)

echo.
echo [3/5] Copiando PDFs para o repositorio...
xcopy "C:\HAGAP\AES\*"       "pdfs\aes\"       /Y /Q 2>nul
xcopy "C:\HAGAP\MEDICOES\*"  "pdfs\medicoes\"  /Y /Q 2>nul
xcopy "C:\HAGAP\PROJETOS\*"  "pdfs\projetos\"  /E /Y /Q 2>nul
echo     OK

echo.
echo [4/5] Adicionando arquivos ao Git...
git add db.json pdfs/ resultado.xlsx
git status --short

echo.
echo [5/5] Enviando para o GitHub (Render atualiza automaticamente)...
git commit -m "atualiza projetos %date% %time:~0,5%"
git push
if %errorlevel% neq 0 (
    echo     AVISO: nao havia nada novo para enviar.
)

echo.
echo ============================================================
echo   PRONTO! O site sera atualizado em ~1 minuto.
echo   Acesse: %URL_SITE%
echo ============================================================
pause
