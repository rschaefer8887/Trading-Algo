@echo off
cd /d "C:\Users\ryanc\OneDrive\repos\Trading_Algo"

set PY="C:\Users\ryanc\AppData\Local\Python\pythoncore-3.14-64\python.exe"

echo y | %PY% "scripts\Exit_ToS.py"
if errorlevel 1 exit /b %ERRORLEVEL%

(echo n & echo y) | %PY% "scripts\Exit_IB_via_GW.py"
exit /b %ERRORLEVEL%