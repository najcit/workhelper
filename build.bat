@echo off
pyinstaller -w -F main.py -n work_helper --clean
copy dist\*.exe .\
rd /S /Q .\*.spec
rd /S /Q .\dist
rd /S /Q .\build
echo build finished