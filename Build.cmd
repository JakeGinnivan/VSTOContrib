@echo off
call "%VS100COMNTOOLS%vsvars32.bat"
mkdir .\build\log\

msbuild.exe /ToolsVersion:4.0 /target:Version "VSTO Contrib.msbuild" 

msbuild.exe /ToolsVersion:4.0 /p:Configuration=Release /p:Platform=AnyCPU /p:TargetFrameworkVersion=v3.5 /p:TargetOfficeVersion=2007 /p:IncludeCastle=True /l:FileLogger,Microsoft.Build.Engine;logfile=.\build\log\NET35_2007_WithCastle.log;verbosity=diagnostic "VSTO Contrib.msbuild" 
if %errorlevel% neq 0 goto end

msbuild.exe /ToolsVersion:4.0 /p:Configuration=Release /p:Platform=AnyCPU /p:TargetFrameworkVersion=v3.5 /p:TargetOfficeVersion=2007 /p:IncludeCastle=False /l:FileLogger,Microsoft.Build.Engine;logfile=.\build\log\NET35_2007_WithoutCastle.log;verbosity=diagnostic "VSTO Contrib.msbuild" 
if %errorlevel% neq 0 goto end

msbuild.exe /ToolsVersion:4.0 /p:Configuration=Release /p:Platform=AnyCPU /p:TargetFrameworkVersion=v3.5 /p:TargetOfficeVersion=2010 /p:IncludeCastle=True /l:FileLogger,Microsoft.Build.Engine;logfile=.\build\log\NET35_2010_WithCastle.log;verbosity=diagnostic "VSTO Contrib.msbuild" 
if %errorlevel% neq 0 goto end

msbuild.exe /ToolsVersion:4.0 /p:Configuration=Release /p:Platform=AnyCPU /p:TargetFrameworkVersion=v3.5 /p:TargetOfficeVersion=2010 /p:IncludeCastle=False /l:FileLogger,Microsoft.Build.Engine;logfile=.\build\log\NET35_2010_WithoutCastle.log;verbosity=diagnostic "VSTO Contrib.msbuild" 
if %errorlevel% neq 0 goto end

msbuild.exe /ToolsVersion:4.0 /p:Configuration=Release /p:Platform=AnyCPU /p:TargetFrameworkVersion=v4.0 /p:TargetOfficeVersion=2007 /p:IncludeCastle=True /l:FileLogger,Microsoft.Build.Engine;logfile=.\build\log\NET40_2007_WithCastle.log;verbosity=diagnostic "VSTO Contrib.msbuild" 
if %errorlevel% neq 0 goto end

msbuild.exe /ToolsVersion:4.0 /p:Configuration=Release /p:Platform=AnyCPU /p:TargetFrameworkVersion=v4.0 /p:TargetOfficeVersion=2007 /p:IncludeCastle=False /l:FileLogger,Microsoft.Build.Engine;logfile=.\build\log\NET40_2007_WithoutCastle.log;verbosity=diagnostic "VSTO Contrib.msbuild" 
if %errorlevel% neq 0 goto end

msbuild.exe /ToolsVersion:4.0 /p:Configuration=Release /p:Platform=AnyCPU /p:TargetFrameworkVersion=v4.0 /p:TargetOfficeVersion=2010 /p:IncludeCastle=True /l:FileLogger,Microsoft.Build.Engine;logfile=.\build\log\NET40_2010_WithCastle.log;verbosity=diagnostic "VSTO Contrib.msbuild" 
if %errorlevel% neq 0 goto end

msbuild.exe /ToolsVersion:4.0 /p:Configuration=Release /p:Platform=AnyCPU /p:TargetFrameworkVersion=v4.0 /p:TargetOfficeVersion=2010 /p:IncludeCastle=False /l:FileLogger,Microsoft.Build.Engine;logfile=.\build\log\NET40_2010_WithoutCastle.log;verbosity=diagnostic "VSTO Contrib.msbuild" 

:end
pause
