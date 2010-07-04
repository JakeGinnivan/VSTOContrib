@echo off
call "%VS100COMNTOOLS%vsvars32.bat"
msbuild.exe /target:Package /fileLogger /ToolsVersion:4.0 /p:Configuration=Release /p:Platform=AnyCPU "FacebookToOutlook.msbuild"
pause
