@ECHO OFF
REM #################################################################################
REM # 処理名　｜mysql-to-csv（起動用バッチ）
REM # 機能　　｜PowerShell起動用のバッチ
REM #--------------------------------------------------------------------------------
REM # 　　　　｜-
REM #################################################################################
ECHO *---------------------------------------------------------
ECHO *
ECHO *  MySQL-to-csv
ECHO *
ECHO *---------------------------------------------------------
ECHO.
ECHO.

powershell -NoProfile -ExecutionPolicy Unrestricted -File .\source\Main.ps1
SET RETURNCODE=%ERRORLEVEL%

ECHO.
ECHO 処理が終了しました。
ECHO いずれかのキーを押すとウィンドウが閉じます。
PAUSE > NUL
EXIT %RETURNCODE%
