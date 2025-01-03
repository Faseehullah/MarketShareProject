@echo off
call conda activate marketsurvey
set PYTHONPATH=%PYTHONPATH%;%~dp0
cd /d %~dp0
code .