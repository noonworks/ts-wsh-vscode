@echo off
cd /d %~dp0
cd ..
@cd
echo * tsc ...
echo   - tsc -p "tsconfig.json"
call tsc -p "tsconfig.json"
IF ERRORLEVEL 1 (
  echo ERRORLEVEL %ERRORLEVEL%
  exit /b 1
)

echo * encoding file ...
echo   - nkf -sx output\built.js^>output\built.sjis.js
nkf -sx output\built.js>output\built.sjis.js
IF ERRORLEVEL 1 (
  echo ERRORLEVEL %ERRORLEVEL%
  exit /b 1
)

call :GET_FOLDER_NAME %CD%
set OUT_FILE=%PARENT_NAME%.js.cmd

echo * build %OUT_FILE% ...
echo   - add shebang
type scripts\files\shebang.short.js.cmd>%OUT_FILE%
echo   - add built js
echo.>>%OUT_FILE%
echo /* generated from typescript source. */>>%OUT_FILE%
type output\built.sjis.js>>%OUT_FILE%
echo.>>%OUT_FILE%
echo /* generated from typescript source. */>>%OUT_FILE%
echo   - add main.js
echo.>>%OUT_FILE%
type scripts\files\main.js>>%OUT_FILE%

echo * finished.
exit /b 0

:GET_FOLDER_NAME
  set PARENT_NAME=%~n1
  exit /b
