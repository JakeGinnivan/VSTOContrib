@echo off
call "%VS100COMNTOOLS%vsvars32.bat"
msbuild.exe /ToolsVersion:4.0 /p:Configuration=Release /p:Platform=AnyCPU "VSTO Contrib.msbuild" 
pause
