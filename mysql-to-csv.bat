@ECHO OFF
REM #################################################################################
REM # �������@�bmysql-to-csv�i�N���p�o�b�`�j
REM # �@�\�@�@�bPowerShell�N���p�̃o�b�`
REM #--------------------------------------------------------------------------------
REM # �@�@�@�@�b-
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
ECHO �������I�����܂����B
ECHO �����ꂩ�̃L�[�������ƃE�B���h�E�����܂��B
PAUSE > NUL
EXIT %RETURNCODE%
