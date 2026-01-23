@echo off
echo Installing dependencies...
echo Please ensure you have enough disk space (approx 500MB+) before running this.
py -3.9 -m pip install PySide6 cefpython3 pywin32 -i http://mirrors.aliyun.com/pypi/simple/ --trusted-host mirrors.aliyun.com
pause
