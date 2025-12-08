@echo off

echo.
echo Restoring backend python packages
echo.
call python -m pip install -r requirements.txt
if "%errorlevel%" neq "0" (
    echo Failed to restore backend python packages
    exit /B %errorlevel%
)

echo.
echo Restoring frontend npm packages
echo.
cd frontend
REM Clean up node_modules and lock file to ensure a fresh install
if exist node_modules rmdir /s /q node_modules
if exist package-lock.json del package-lock.json

REM Install npm-force-resolutions if not present
call npm install npm-force-resolutions --save-dev
if "%errorlevel%" neq "0" (
    echo Failed to install npm-force-resolutions
    exit /B %errorlevel%
)

REM Run npm-force-resolutions to enforce resolutions in package.json
call npx npm-force-resolutions
if "%errorlevel%" neq "0" (
    echo Failed to run npm-force-resolutions
    exit /B %errorlevel%
)

REM Install all frontend dependencies
call npm install
if "%errorlevel%" neq "0" (
    echo Failed to restore frontend npm packages
    exit /B %errorlevel%
)

echo.
echo Building frontend
echo.
call npm run build
if "%errorlevel%" neq "0" (
    echo Failed to build frontend
    exit /B %errorlevel%
)

echo.    
echo Starting backend    
echo.    
cd ..  
start http://127.0.0.1:50505
call python -m uvicorn app:app  --port 50505 --reload
if "%errorlevel%" neq "0" (    
    echo Failed to start backend    
    exit /B %errorlevel%    
) 
