@echo off

call Distribute_DLL_Work.bat F:\Documents\Projects\DataMining\Protein_Digestion_Simulator\ProteinDigestionSimulator\bin\
call Distribute_DLL_Work.bat F:\Documents\Projects\DataMining\Protein_Digestion_Simulator\ProteinDigestionSimulator\bin\DLL\
call Distribute_DLL_Work.bat F:\Documents\Projects\DataMining\Protein_Digestion_Simulator\Lib\

call Distribute_DLL_Work.bat F:\Documents\Projects\KenAuberry\Organism_Database_Handler\Executables\Debug\
call Distribute_DLL_Work.bat F:\Documents\Projects\KenAuberry\Organism_Database_Handler\Executables\Release\
call Distribute_DLL_Work.bat F:\Documents\Projects\KenAuberry\Organism_Database_Handler\Protein_Uploader\bin\
call Distribute_DLL_Work.bat F:\Documents\Projects\KenAuberry\Organism_Database_Handler\Lib\

call Distribute_DLL_Work.bat F:\Documents\Projects\DataMining\Validate_Fasta_File\TestValidateFastaFileDLL\bin\

if not "%1"=="NoPause" pause
