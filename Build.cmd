@echo off
call "%VS110COMNTOOLS%vsvars32.bat"
mkdir .\build\log\

rmdir .\build\Artifacts\ /s /q

msbuild.exe /ToolsVersion:4.0 /p:Configuration=Release /p:Platform=AnyCPU /p:TargetFrameworkVersion=v4.5 /l:FileLogger,Microsoft.Build.Engine;logfile=.\build\log\NET45.log;verbosity=diagnostic "VSTO Contrib.msbuild" 
if %errorlevel% neq 0 goto end

msbuild.exe /ToolsVersion:4.0 /p:Configuration=Release /p:Platform=AnyCPU /p:TargetFrameworkVersion=v4.0 /l:FileLogger,Microsoft.Build.Engine;logfile=.\build\log\NET40.log;verbosity=diagnostic "VSTO Contrib.msbuild" 
if %errorlevel% neq 0 goto end

:end
pause
