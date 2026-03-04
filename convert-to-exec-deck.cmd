@echo off
setlocal

set PYEXE=
where python >nul 2>nul && set PYEXE=python
if "%PYEXE%"=="" (
  where py >nul 2>nul && set PYEXE=py
)

if "%PYEXE%"=="" (
  echo Python not found. Install Python 3.8+ and ensure it is on PATH.
  exit /b 1
)

%PYEXE% run_pipeline.py %*
exit /b %ERRORLEVEL%
