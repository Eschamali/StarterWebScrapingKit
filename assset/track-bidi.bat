@echo off
setlocal
rem Save the chromium-bidi mapperTab.js matching a Chrome for Testing channel.
rem Usage: track-bidi.bat [Stable|Beta|Dev|Canary]   (no argument: choose from a menu)
rem If npm has no build for that Chrome version, fall back to the
rem chromium-bidi version pinned by the latest puppeteer-core.
cd /d "%~dp0"

set "CHANNEL=%~1"
if not "%CHANNEL%"=="" goto :run

echo   1. Stable
echo   2. Beta
echo   3. Dev
echo   4. Canary
choice /c 1234 /n /m "Channel [1-4]: "
if errorlevel 4 (set "CHANNEL=Canary") else if errorlevel 3 (set "CHANNEL=Dev") else if errorlevel 2 (set "CHANNEL=Beta") else (set "CHANNEL=Stable")
rem Run with the chosen channel, then keep the window open when double-clicked
call "%~f0" %CHANNEL%
pause
exit /b

:run

rem 1. Chrome version of the channel -> npm version (last "." becomes "-")
for /f "usebackq delims=" %%a in (`powershell -NoProfile -Command "(Invoke-RestMethod https://googlechromelabs.github.io/chrome-for-testing/last-known-good-versions.json).channels.%CHANNEL%.version"`) do set "CHROME=%%a"
if "%CHROME%"=="" (
  echo Failed to read Chrome %CHANNEL% version
  exit /b 1
)
for /f "usebackq delims=" %%a in (`powershell -NoProfile -Command "'%CHROME%' -replace '\.(\d+)$','-$1'"`) do set "BIDI=%%a"
set "SOURCE=cft"

rem 2. Fall back to puppeteer-core if npm does not have that build
curl -fsS -o nul "https://registry.npmjs.org/chromium-bidi/%BIDI%" 2>nul
if errorlevel 1 (
  echo chromium-bidi %BIDI% is not on npm, using puppeteer-core instead
  for /f "usebackq tokens=1,2" %%a in (`powershell -NoProfile -Command "$j = Invoke-RestMethod https://registry.npmjs.org/puppeteer-core/latest; $j.version + ' ' + $j.dependencies.'chromium-bidi'"`) do (
    set "PPTR=%%a"
    set "BIDI=%%b"
  )
  set "SOURCE=puppeteer-core"
)
if "%SOURCE%"=="puppeteer-core" set "SOURCE=puppeteer-core %PPTR%"

echo Chrome %CHANNEL% %CHROME% / chromium-bidi %BIDI% (%SOURCE%)

rem 3. Skip if unchanged
set "STATE=%CHANNEL%-version.txt"
set "PREV="
if exist "%STATE%" set /p PREV=<"%STATE%"
if "%PREV%"=="%BIDI%" (
  echo No change
  exit /b 0
)

rem 4. Download mapperTab.js and record
curl -fsS -o "mapperTab.js" "https://unpkg.com/chromium-bidi@%BIDI%/out/Default/gen/src/mapperTab.js" || exit /b 1

>"%STATE%" echo %BIDI%
>>history.log echo %date% %time:~0,8% %CHANNEL% Chrome %CHROME% chromium-bidi %BIDI% (%SOURCE%)
echo Updated: %PREV% -^> %BIDI%
