@echo on
python_installer\python-3.8.5-amd64.exe /passive InstallAllUsers=1 PrependPath=1 Include_test=0
python.exe --version >NUL 2>&1
if errorlevel 1 goto notpath
echo Python.exe found in the path.
pip install PyQt5
pip install openpyxl
pip install docx-mailmerge
if errorlevel 1 goto error
goto end
:NOTPATH