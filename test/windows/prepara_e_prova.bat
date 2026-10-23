@echo off
rem Prepara Windows ed esegue tutti i test di FileGUY (come test\linux\prepara_e_prova.sh su Linux).
rem Uso:  test\windows\prepara_e_prova.bat [cartella dei programmi, es. C:\prova]
rem   Nella cartella dei programmi ci sono gli .exe, config.dat e i modelli: si provano anche gli .exe.
rem Per provare anche una stampa vera:  set FILEGUY_STAMPA_VERA=1  (e se serve  set FILEGUY_STAMPANTE=nome)
setlocal
pushd "%~dp0..\.."
set "PROVA=%~1"
if "%PROVA%"=="" set "PROVA=C:\prova"
set "FILEGUY_PROVA=%PROVA%"

echo === Ambiente Python (.venv-windows) ===
if not exist ".venv-windows\Scripts\python.exe" (
    py -3 -m venv .venv-windows 2>nul || python -m venv .venv-windows
)
if not exist ".venv-windows\Scripts\python.exe" (
    echo Python non trovato: installalo da python.org ^(spunta "Add python.exe to PATH"^)
    popd
    exit /b 1
)
call ".venv-windows\Scripts\activate.bat"
python -m pip install --upgrade pip >nul
pip install -r requirements.txt
if errorlevel 1 (
    echo Installazione delle librerie non riuscita
    popd
    exit /b 1
)

set ESITO=0
for %%T in (test_windows test_moduli test_editor_modelli test_editor_excel test_excel_resa) do (
    echo.
    echo === test\%%T.py ===
    python -W ignore test\%%T.py || set ESITO=1
)

echo.
echo === test\esegui_test.py (servizio) ===
if exist "%PROVA%\modelli" (
    python -W ignore test\esegui_test.py SingleGUY "%PROVA%\modelli" || set ESITO=1
) else (
    python -W ignore test\esegui_test.py SingleGUY || set ESITO=1
)

echo.
echo === test\test_eseguibili.py (gli .exe di %PROVA%) ===
if exist "%PROVA%\SingleGUY.exe" (
    python -W ignore test\test_eseguibili.py "%PROVA%" || set ESITO=1
) else (
    echo Nessun .exe in %PROVA%: prova degli eseguibili saltata
)

echo.
if "%ESITO%"=="0" (
    echo TUTTO OK. L'immagine con i font da guardare: test\_run\windows\font_windows.png
) else (
    echo CI SONO ERRORI: vedi sopra le righe [ERR]
)
popd
exit /b %ESITO%
