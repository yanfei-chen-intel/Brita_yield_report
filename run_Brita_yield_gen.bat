@echo off
echo ========================================
echo   Brita Yield Report Generation
echo ========================================
echo.

REM Record start time
set START_DATE=%DATE%
set START_TIME=%TIME%
echo Start Time: %START_TIME% %START_DATE%
echo.

REM Clear output folder before running
echo Clearing output folder...
if exist "%~dp0output" (
    rd /s /q "%~dp0output"
)
mkdir "%~dp0output"
echo Output folder cleared.
echo.

REM Run main script (venv creation and package installation are handled inside)
echo [1/1] Running Brita_yield_gen.py...
echo.
python "%~dp0Brita_yield_gen.py"
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Brita_yield_gen.py execution failed
    goto :error
)

REM Record end time
set END_DATE=%DATE%
set END_TIME=%TIME%
echo.
echo ========================================
echo   Process Completed Successfully
echo ========================================
echo Start Time: %START_TIME% %START_DATE%
echo End Time:   %END_TIME% %END_DATE%
call :GetDuration %START_TIME% %END_TIME%
echo Duration:   %DURATION%
echo ========================================
goto :end

:error
set END_DATE=%DATE%
set END_TIME=%TIME%
echo.
echo ========================================
echo   Process Failed
echo ========================================
echo Start Time: %START_TIME% %START_DATE%
echo End Time:   %END_TIME% %END_DATE%
echo ========================================
goto :end

:GetDuration
set START=%~1
set END=%~2
set START_H=0
set START_M=0
set START_S=0
set END_H=0
set END_M=0
set END_S=0
for /F "tokens=1-3 delims=:., " %%a in ("%START%") do (
    set START_H=%%a
    set START_M=%%b
    set START_S=%%c
)
for /F "tokens=1-3 delims=:., " %%a in ("%END%") do (
    set END_H=%%a
    set END_M=%%b
    set END_S=%%c
)
set /A START_H=1%START_H%-100
set /A START_M=1%START_M%-100
set /A START_S=1%START_S%-100
set /A END_H=1%END_H%-100
set /A END_M=1%END_M%-100
set /A END_S=1%END_S%-100
set /A START_SEC=%START_H%*3600 + %START_M%*60 + %START_S%
set /A END_SEC=%END_H%*3600 + %END_M%*60 + %END_S%
if %END_SEC% LSS %START_SEC% set /A END_SEC=%END_SEC% + 86400
set /A DURATION_SEC=%END_SEC% - %START_SEC%
set /A DUR_H=%DURATION_SEC% / 3600
set /A DUR_M=(%DURATION_SEC% %% 3600) / 60
set /A DUR_S=%DURATION_SEC% %% 60
set DURATION=%DUR_H%h %DUR_M%m %DUR_S%s
goto :eof

:end
echo.
pause
