@echo off
cd /d "C:\Users\ryanc\OneDrive\repos\Trading_Algo"

set "PY=C:\Users\ryanc\AppData\Local\Python\pythoncore-3.14-64\python.exe"

"%PY%" "scripts\Open_Trades_GW.py"
if errorlevel 1 exit /b %ERRORLEVEL%

"%PY%" "scripts\Open_Trades_ToS.py"
if errorlevel 1 exit /b %ERRORLEVEL%

"%PY%" "scripts\Open_Trade_ToS2.py"
exit /b %ERRORLEVEL%
