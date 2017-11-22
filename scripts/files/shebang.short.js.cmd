@set @dummy=0/*
@echo off
if exist %WinDir%\SysNative\cmd.exe (
  ::echo "32bit-cmd.exe on 64bit-Windows"
  set CSCRIPT32=cscript
) else (
  if "%PROCESSOR_ARCHITECTURE%" == "x86" (
    ::echo "32bit-cmd.exe on 32bit-Windows"
    set CSCRIPT32=cscript
  ) else (
    ::echo "64bit-cmd.exe on 64bit-Windows"
    set CSCRIPT32=%windir%\SysWOW64\cscript.exe
  )
)
%CSCRIPT32% //nologo //E:JScript "%~f0" %*
set JS_ERR_LEVEL=%ERRORLEVEL%
if NOT "%JS_ERR_LEVEL%" == "0" (
  pause
)
exit /b %JS_ERR_LEVEL%
*/