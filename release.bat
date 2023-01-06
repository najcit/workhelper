@echo off
pyinstaller -w -F main.py -n work_helper
python zip.py --app dist\work_helper.exe --release v1.0.0
del /S /Q work_helper.spec
rd /S /Q dist
rd /S /Q build
echo release finished