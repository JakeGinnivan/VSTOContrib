@echo off

%~dp0\tools\GitVersion.exe /updateAssemblyInfo /proj %~dp0\VSTOContrib.proj

pause