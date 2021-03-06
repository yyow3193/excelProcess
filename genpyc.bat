@echo off


path = C:\Users\Administrator\AppData\Local\Programs\Python\Python38-32\Scripts;
pyinstaller.exe -F D:\excelProccess\excelProcess\main.py

copy /y .\__pycache__\main.cpython-38.pyc .\

pause