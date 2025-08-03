@echo off
echo Starting Unified INDIRECT Resolver
echo ===================================

cd /d "%~dp0"

echo Current Directory: %CD%
echo.

echo Starting Unified Resolver...
C:\Users\user\anaconda3\python.exe unified_indirect_resolver.py

echo.
echo Unified Resolver closed
pause