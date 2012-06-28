@echo off
call "%VS100COMNTOOLS%vsvars32.bat"
mkdir .\build\log\

rmdir .\build\Artifacts\

msbuild.exe /ToolsVersion:4.0 /p:Configuration=Release /p:Platform=AnyCPU /p:TargetFrameworkVersion=v4.0 /p:TargetOfficeVersion=2010 /l:FileLogger,Microsoft.Build.Engine;logfile=.\build\log\NET40_2010.log;verbosity=diagnostic "VSTO Contrib.msbuild" 
if %errorlevel% neq 0 goto end

:end
pause
