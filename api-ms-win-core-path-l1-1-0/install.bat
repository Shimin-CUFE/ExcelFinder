@echo off
cd System32
copy /-Y api-ms-win-core-path-l1-1-0.dll C:\Windows\System32
cd ..
cd SysWOW64
copy /-Y api-ms-win-core-path-l1-1-0.dll C:\Windows\SysWOW64
pause