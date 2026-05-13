@echo off
cd /d "C:\Users\ryanc\OneDrive\repos\Trading_Algo"

set "PY=C:\Users\ryanc\AppData\Local\Python\pythoncore-3.14-64\python.exe"

REM Remvoe these comments to run the Open_Trades_GW.py script
REM Uncomment this line to run the Open_Trades_GW.py script
REM "%PY%" "scripts\Open_Trades_GW.py"
REM if errorlevel 1 exit /b %ERRORLEVEL%


"%PY%" "scripts\Open_Trades_ToS.py"
if errorlevel 1 exit /b %ERRORLEVEL%

"%PY%" "scripts\Open_Trade_ToS2.py"
exit /b %ERRORLEVEL%
