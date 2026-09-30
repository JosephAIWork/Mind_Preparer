@echo off
rem Mind Ready -- FastAPI backend + the Figma-built React front-end, one URL.
rem Build the front-end first (once, or after UI changes):
rem     cd "..\FigmaOutput" && npm install && npm run build
cd /d "%~dp0"
set MIND_READY_PORT=8600
set PY=python
if exist ".venv\Scripts\python.exe" set PY=.venv\Scripts\python.exe

rem Already running (started earlier, or from VS Code)? Then only open it.
powershell -NoProfile -Command "try { Invoke-WebRequest -UseBasicParsing http://127.0.0.1:%MIND_READY_PORT%/api/health -TimeoutSec 2 | Out-Null; exit 0 } catch { exit 1 }"
if not errorlevel 1 (
  echo Mind Ready is already running. Opening the browser.
  start "" http://localhost:%MIND_READY_PORT%
  exit /b 0
)

echo Starting Mind Ready on http://localhost:%MIND_READY_PORT% ...
rem The server runs in its own window (kept open on error so the message can be read).
start "Mind Ready server - close this window to stop" cmd /k "%PY% -m app.web.server"

rem Open the browser only once the server answers (up to 90 s). The check goes to
rem 127.0.0.1, the address the server binds: through "localhost" the first request
rem of a process can spend 2 s on IPv6 before falling back, longer than the timeout.
powershell -NoProfile -Command "for ($i = 0; $i -lt 90; $i++) { try { Invoke-WebRequest -UseBasicParsing http://127.0.0.1:%MIND_READY_PORT%/api/health -TimeoutSec 2 | Out-Null; exit 0 } catch { Start-Sleep -Seconds 1 } }; exit 1"
if errorlevel 1 goto failed
echo Server is up. Opening the browser.
start "" http://localhost:%MIND_READY_PORT%
exit /b 0

:failed
echo The server did not answer after 90 s. Read the error in the "Mind Ready server" window.
pause
