@echo off
pyinstaller -w -F main.py -n work_helper
copy dist\work_helper.exe .
del /S /Q work_helper.spec
rd /S /Q dist
rd /S /Q build
echo build finished