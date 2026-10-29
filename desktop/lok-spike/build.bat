@echo off
rem 試作のビルド（Visual Studio 2022 BuildTools の x64 環境で cl を使う）
call "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvars64.bat" >nul
cd /d "%~dp0"
cl /nologo /utf-8 /W3 /I include lok_spike.c /Fe:lok_spike.exe /link user32.lib
