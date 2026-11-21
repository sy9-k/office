@echo off
call "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvars64.bat" >nul
cd /d "%~dp0"
cl /nologo /utf-8 /W3 /I include lok_server.c /Fe:lok_server.exe /link user32.lib
