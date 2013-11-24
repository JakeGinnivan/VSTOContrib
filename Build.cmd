@echo off
set framework=v4.0.30319

mkdir build\log
"%SystemDrive%\Windows\Microsoft.NET\Framework\%framework%\MSBuild.exe" /ToolsVersion:4.0 /p:Configuration=Release /p:Platform=AnyCPU /p:Version=%1 /l:FileLogger,Microsoft.Build.Engine;logfile=.\build\log\build.log;verbosity=diagnostic "VSTO Contrib.proj" 

pause