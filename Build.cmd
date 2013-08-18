@echo off
call "%VS110COMNTOOLS%vsvars32.bat"

mkdir build\log
msbuild.exe /ToolsVersion:4.0 /p:Configuration=Release /p:Platform=AnyCPU /l:FileLogger,Microsoft.Build.Engine;logfile=.\build\log\build.log;verbosity=diagnostic "VSTO Contrib.proj" 

pause