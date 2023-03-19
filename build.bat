@echo off
pip freeze > requirements.txt
echo refresh requirements info

pyinstaller -w -F main.py -n workhelper --icon image\window\workhelper.ico
copy dist\workhelper.exe .
del /S /Q workhelper.spec
rd /S /Q dist
rd /S /Q build
echo build finished