@echo off

echo Copying to %1
xcopy ValidateFastaFile.dll %1 /d /Y

if exist %1\ValidateFastaFile.pdb (xcopy ValidateFastaFile.pdb %1 /d /Y)
if exist %1\ValidateFastaFile.xml (xcopy ValidateFastaFile.xml %1 /d /Y)
