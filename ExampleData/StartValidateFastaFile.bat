..\bin\ValidateFastaFile.exe JunkTest.fasta /P:ValidateFastaFile_ModelSettings.xml
@echo off
echo.
echo.
@echo on

..\bin\ValidateFastaFile.exe JunkTest_UTF8.fasta /P:ValidateFastaFile_ModelSettings.xml
@echo off
echo.
echo.
@echo on

..\bin\ValidateFastaFile.exe JunkTest.fasta /P:ValidateFastaFile_CreateStatsFile.xml
..\bin\ValidateFastaFile.exe JunkTest_UTF8.fasta /P:ValidateFastaFile_CreateStatsFile.xml
@echo off
echo.
echo.
@echo on

..\bin\ValidateFastaFile.exe QC_Standards_2004-01-21_Dup.fasta /P:ValidateFastaFileOptions_GenerateFixedFasta.xml

@echo off
rem Other examples:

rem ..\bin\ValidateFastaFile.exe H_Sapiens_IPI_2009-01-20.fasta /F /D
rem ValidateFastaFile.exe /I:\\gigasax\dms_organism_Files\*.fasta /p:ValidateFastaFileOptions.xml /c /s:2
