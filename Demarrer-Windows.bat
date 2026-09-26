@echo off
setlocal
cd /d "%~dp0"
node -v >nul 2>&1
if errorlevel 1 (
  echo.
  echo Node.js n'est pas installe sur cet ordinateur.
  echo Installe-le depuis https://nodejs.org (version LTS), puis relance ce fichier.
  echo.
  pause
  exit /b
)
if not exist node_modules (
  echo Premiere fois : installation des composants (Internet requis)...
  call npm install
  if errorlevel 1 (
    echo.
    echo L'installation a echoue. Verifie ta connexion Internet et reessaie.
    pause
    exit /b
  )
)
echo.
echo Quiz demarre. Laisse cette fenetre OUVERTE pendant la partie.
echo Sur cet ordinateur, ouvre http://localhost:3000
echo.
node server.js
echo.
echo (Le serveur s'est arrete.)
pause
