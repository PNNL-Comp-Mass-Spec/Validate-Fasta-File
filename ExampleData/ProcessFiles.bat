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

rem Create a fixed fasta file, renaming proteins with duplicate names
%ExePath% JunkTest.fasta /F /R

rem Create a _ProteinHashes.txt file
%ExePath% JunkTest.fasta /B

rem Create a FastaFileStats .txt file
%ExePath% JunkTest.fasta /C
%ExePath% JunkTest_UTF8.fasta /C

rem Process a FASTA file with UTF-8 encoding
%ExePath% JunkTest_UTF8.fasta /B /F /R

rem Process a gzipped FASTA file
%ExePath% JunkTest2.fasta.gz /F /R

%ExePath% Shewanella_2003-12-19.fasta /F

rem Use a parameter file to specify options
%ExePath% QC_Standards_2004-01-21_Dup.fasta /P:ValidateFastaFileOptions_GenerateFixedFasta.xml

rem Process a file with DNA sequences
%ExePath% H_sapiens_DNA_Excerpts.fasta

:Done

pause
