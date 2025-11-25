@echo off
REM Script di installazione per mk2docx (Windows)

echo === Installazione dipendenze per mk2docx ===
echo.

REM Verifica se Python è installato
python --version >nul 2>&1
if errorlevel 1 (
    echo ERRORE: Python non trovato!
    echo Installa Python da https://www.python.org/downloads/
    echo Assicurati di selezionare "Add Python to PATH" durante l'installazione
    pause
    exit /b 1
)

REM Installa pypandoc
echo 1. Installazione della libreria Python pypandoc...
python -c "import pypandoc" >nul 2>&1
if errorlevel 1 (
    echo    Installazione via pip...
    python -m pip install --upgrade pip
    python -m pip install pypandoc
) else (
    echo    pypandoc e' già installato
)

echo.
echo 2. Installazione di Pandoc...

REM Verifica se pandoc è già installato
pandoc --version >nul 2>&1
if not errorlevel 1 (
    echo    Pandoc e' già installato
    goto :done
)

REM Verifica se winget è disponibile (Windows 10/11)
winget --version >nul 2>&1
if not errorlevel 1 (
    echo    Installazione via winget...
    winget install --id JohnMacFarlane.Pandoc -e --silent
    goto :done
)

REM Verifica se chocolatey è disponibile
choco --version >nul 2>&1
if not errorlevel 1 (
    echo    Installazione via Chocolatey...
    choco install pandoc -y
    goto :done
)

REM Se nessun package manager è disponibile
echo.
echo    ATTENZIONE: Nessun package manager trovato (winget/chocolatey)
echo.
echo    Opzioni per installare Pandoc:
echo    1. Scarica l'installer da: https://pandoc.org/installing.html
echo    2. Oppure installa Chocolatey da: https://chocolatey.org/install
echo       e poi esegui di nuovo questo script
echo.
echo    Dopo l'installazione manuale di Pandoc, riesegui questo script.
pause
exit /b 1

:done
echo.
echo === Installazione completata con successo! ===
echo.
echo Puoi ora utilizzare mk2docx con:
echo   python mk2docx.py input.md -o output.docx
echo.
pause
