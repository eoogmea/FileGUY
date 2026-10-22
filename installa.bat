@echo off
rem Installa tutto quello che serve a FileGUY su Windows (doppio clic oppure dal prompt).
rem   installa.bat              ambiente Python e librerie
rem   installa.bat --compila    in piu' crea gli eseguibili in dist\
rem Se Python non c'e' lo installa con winget (Windows 10/11), poi esegue installa.py.
setlocal
pushd "%~dp0"

set "PY="
py -3 --version >nul 2>&1 && set "PY=py -3"
if not defined PY python --version >nul 2>&1 && set "PY=python"
if not defined PY (
    echo Python non trovato: lo installo con winget...
    winget install --id Python.Python.3.14 -e --scope user --accept-package-agreements --accept-source-agreements
    if errorlevel 1 (
        echo.
        echo Installazione di Python non riuscita. Installalo da https://www.python.org/downloads/
        echo spuntando "Add python.exe to PATH", poi rilancia installa.bat
        goto fine_errore
    )
    rem Il PATH nuovo vale solo nelle finestre aperte dopo: si usa il percorso dove winget lo mette
    set PY="%LOCALAPPDATA%\Programs\Python\Python314\python.exe"
    if not exist "%LOCALAPPDATA%\Programs\Python\Python314\python.exe" (
        echo Python installato: chiudi questa finestra e rilancia installa.bat
        goto fine_errore
    )
)

%PY% installa.py %*
if errorlevel 1 goto fine_errore
popd
echo.
pause
exit /b 0

:fine_errore
popd
echo.
pause
exit /b 1
