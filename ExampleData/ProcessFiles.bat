rem Create a fixed fasta file, renaming proteins with duplicate names
..\bin\ValidateFastaFile.exe JunkTest.fasta /F /R

rem Create a _ProteinHashes.txt file
..\bin\ValidateFastaFile.exe JunkTest.fasta /B

rem Process a FASTA file with UTF-8 encoding
..\bin\ValidateFastaFile.exe JunkTest_UTF8.fasta /B /F /R

rem Process a gzipped FASTA file
..\bin\ValidateFastaFile.exe JunkTest2.fasta.gz /F /R

..\bin\ValidateFastaFile.exe Shewanella_2003-12-19.fasta /F

rem Use a parameter file to specify options
..\bin\ValidateFastaFile.exe QC_Standards_2004-01-21_Dup.fasta /P:ValidateFastaFileOptions_GenerateFixedFasta.xml

pause
