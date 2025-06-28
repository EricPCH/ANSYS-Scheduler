@echo off
REM Configure the path to ansysedt here
set "ANSYSEDT_PATH=C:\Program Files\ANSYS Inc\v251\AnsysEM\ansysedt"

echo 執行 IronPython 腳本...
"C:\Program Files\ANSYS Inc\v251\AnsysEM\common\IronPython\ipy64.exe" scheduler.py
