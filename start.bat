@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo.
echo ============================================
echo   NEXT-GEN PRO - Structural Takeoff Server
echo ============================================
echo.

REM Create .env from template if missing
if not exist ".env" (
  if exist ".env.example" (
    copy /Y ".env.example" ".env" >nul
    echo [CREATED] .env from .env.example
    echo.
    echo  IMPORTANT: Edit .env and set your real Gemini API key:
    echo    GEMINI_API_KEY=your_real_key
    echo.
    echo  Get a free key: https://aistudio.google.com/apikey
    echo.
    notepad ".env"
    echo.
    echo  After saving .env, press any key to start the server...
    pause >nul
  ) else (
    echo [ERROR] .env.example not found. Create .env manually with:
    echo   GEMINI_API_KEY=your_real_key
    pause
    exit /b 1
  )
)

REM Load GEMINI_API_KEY from .env into this process (enables the API)
set "GEMINI_API_KEY="
for /f "usebackq tokens=1,* delims== eol=#" %%A in (".env") do (
  if /I "%%A"=="GEMINI_API_KEY" (
    set "GEMINI_API_KEY=%%B"
  )
)

REM Trim spaces / quotes that Notepad users sometimes add
if defined GEMINI_API_KEY (
  set "GEMINI_API_KEY=%GEMINI_API_KEY:"=%"
)

if not defined GEMINI_API_KEY (
  echo [ERROR] GEMINI_API_KEY is empty in .env
  echo Edit .env and set: GEMINI_API_KEY=your_real_key
  notepad ".env"
  pause
  exit /b 1
)

echo %GEMINI_API_KEY% | findstr /I /C:"your_google_gemini_api_key_here" /C:"your_key" /C:"your_real_key" /C:"your_api_key_here" >nul
if %ERRORLEVEL%==0 (
  echo [ERROR] .env still has a placeholder API key.
  echo Replace it with your real key from https://aistudio.google.com/apikey
  notepad ".env"
  pause
  exit /b 1
)

echo [OK] Loaded GEMINI_API_KEY from .env
echo [OK] Starting Uvicorn on http://127.0.0.1:8000
echo.

REM Prefer venv python if present
if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" -m uvicorn backend:app --host 0.0.0.0 --port 8000 --reload
) else (
  python -m uvicorn backend:app --host 0.0.0.0 --port 8000 --reload
)

endlocal
