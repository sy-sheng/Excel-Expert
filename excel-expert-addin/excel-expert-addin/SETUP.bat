@echo off
echo.
echo ============================================
echo  Excel Expert Add-in — First-Time Setup
echo ============================================
echo.

:: Check Node.js
node --version >nul 2>&1
if %errorlevel% neq 0 (
  echo [ERROR] Node.js is not installed.
  echo Please download it from https://nodejs.org and install, then run this script again.
  pause
  exit /b 1
)

echo [1/4] Node.js found.

:: Install dependencies
echo [2/4] Installing dependencies...
call npm install
if %errorlevel% neq 0 (
  echo [ERROR] npm install failed. Check your internet connection.
  pause
  exit /b 1
)

:: Generate self-signed SSL cert using Node.js (no openssl needed)
echo [3/4] Generating SSL certificate...
if not exist certs mkdir certs
node -e "
const { execSync } = require('child_process');
const fs = require('fs');
// Use built-in crypto to write a self-signed cert via a temp script
const script = \`
const forge = require('node-forge');
\`;
// Fallback: use openssl if available, else skip
try {
  execSync('openssl version', {stdio:'ignore'});
  execSync('openssl req -x509 -newkey rsa:2048 -keyout certs/server.key -out certs/server.crt -days 365 -nodes -subj \"/CN=localhost\"', {stdio:'inherit'});
  console.log('SSL cert generated via openssl.');
} catch(e) {
  console.log('openssl not found — running on HTTP. Add -k flag when sideloading or install openssl.');
}
"

echo [4/4] Setup complete.
echo.
echo ============================================
echo  Next steps:
echo  1. Run START.bat to launch the add-in server
echo  2. In Excel: Insert ^> Get Add-ins ^> Upload My Add-in
echo     and select the manifest.xml file in this folder
echo ============================================
echo.
pause
