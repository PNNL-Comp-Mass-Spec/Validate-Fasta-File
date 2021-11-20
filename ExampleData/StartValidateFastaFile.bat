@echo off

set ExePath=ValidateFastaFile.exe

if exist %ExePath% goto DoWork
if exist ..\%ExePath% set ExePath=..\%ExePath% && goto DoWork
if exist ..\Bin\%ExePath% set ExePath=..\Bin\%ExePath% && goto DoWork

echo Executable not found: %ExePath%
goto Done

:DoWork
echo.
echo Processing with %ExePath%
echo.

%ExePath% JunkTest.fasta /P:ValidateFastaFile_ModelSettings.xml
@echo off
echo.
echo.
@echo on

%ExePath% JunkTest_UTF8.fasta /P:ValidateFastaFile_ModelSettings.xml
@echo off
echo.
echo.
@echo on

%ExePath% JunkTest.fasta /P:ValidateFastaFile_CreateStatsFile.xml
%ExePath% JunkTest_UTF8.fasta /P:ValidateFastaFile_CreateStatsFile.xml
@echo off
echo.
echo.
@echo on

%ExePath% QC_Standards_2004-01-21_Dup.fasta /P:ValidateFastaFileOptions_GenerateFixedFasta.xml

@echo off
rem Other examples:

rem %ExePath% H_Sapiens_IPI_2009-01-20.fasta /F /D
rem ValidateFastaFile.exe /I:\\gigasax\dms_organism_Files\*.fasta /p:ValidateFastaFileOptions.xml /c /s:2


:Done

pause

