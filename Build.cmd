@echo off
call "%VS90COMNTOOLS%vsvars32.bat"
msbuild.exe /ToolsVersion:3.5 /p:Configuration=Release /p:Platform=AnyCPU "VSTO Contrib.msbuild" 
pause
