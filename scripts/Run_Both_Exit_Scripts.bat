@echo off
cd /d "C:\Users\ryanc\OneDrive\repos\Trading_Algo"

echo y | python "scripts\Exit_ToS.py"

(echo n & echo y) | python "scripts\Exit_IB_via_GW.py"
