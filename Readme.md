
# Validate Fasta File

This program reads a [FASTA file](https://en.wikipedia.org/wiki/FASTA_format) 
and displays statistics on the number of proteins and number of residues.
It also verifies that the protein names, descriptions, and sequences are in the correct format.

This is a command line program.  For a graphical (GUI) program, see the
[Protein Digestion Simulator](https://github.com/PNNL-Comp-Mass-Spec/Protein-Digestion-Simulator)
which includes a FASTA file validation window.  The executable for 
Protein Digestion Simulator is
[available on AppVeyor](https://ci.appveyor.com/project/PNNLCompMassSpec/protein-digestion-simulator/build/artifacts)
or [on GitHub](https://github.com/PNNL-Comp-Mass-Spec/Protein-Digestion-Simulator/releases)

### Continuous Integration

The latest version of the application may be available on the [AppVeyor CI server](https://ci.appveyor.com/project/PNNLCompMassSpec/validate-fasta-file/build/artifacts),
but builds are deleted after 6 months. \
Otherwise, see the [GitHub release page](https://github.com/PNNL-Comp-Mass-Spec/Validate-Fasta-File/releases)

[![Build status](https://ci.appveyor.com/api/projects/status/gg6u438llyhiupv3?svg=true)](https://ci.appveyor.com/project/PNNLCompMassSpec/validate-fasta-file)

## Additional Features

The program can optionally create a new, fixed version of the FASTA file
where proteins with duplicate sequences have been consolidated, and
proteins with duplicate names have been renamed.

To remove duplicates from huge fasta files (over 1 GB in size), first 
create the ProteinHashes.txt file by calling this program with:
```
ValidateFastaFile.exe Proteins.fasta /B /SkipDupeSeqCheck /SkipDupeNameCheck
```

Next call the program again, providing the name of the ProteinHashes file:
```
ValidateFastaFile.exe Proteins.fasta /HashFile:Proteins_ProteinHashes.txt
```

## Installation

* Download Validate-Fasta-File.zip from [AppVeyor](https://ci.appveyor.com/project/PNNLCompMassSpec/validate-fasta-file/build/build/artifacts)
  * Alternatively, download from the [GitHub release page](https://github.com/PNNL-Comp-Mass-Spec/Validate-Fasta-File/releases)
* Extract the files
* Run ValidateFastaFile.exe from the the command prompt

## Console Switches

```
ValidateFastaFile.exe
 /I:InputFilePath.fasta [/O:OutputDirectoryPath]
 [/P:ParameterFilePath] [/C]
 [/F] [/R] [/D] [/L] [/V] [/KeepSameName]
 [/AllowDash] [/AllowAsterisk]
 [/SkipDupeNameCheck] [/SkipDupeSeqCheck]
 [/B] [/HashFile]
 [/X] [/S:[MaxLevel]] [/Q]
```

The input file path can contain the wildcard character * and should point 
to a fasta file.

The output directory path is optional, and is only used if /C is used.
If omitted, the output stats file will be created in the directory 
containing the .Exe file.

The parameter file path is optional.  If included, it should point to a 
valid XML parameter file.

Use /C to specify that an output file should be created, rather than 
displaying the results on the screen.

Use /F to generate a new, fixed .Fasta file (long protein names will be auto-shortened).
At the same time, a file with protein names and hash values
for each unique protein sequences will be generated (_UniqueProteinSeqs.txt). 
This file will also list the other proteins that have duplicate sequences as the first protein mapped to each sequence.
If duplicate sequences are found, then an easily parseable mapping file will
also be created (_UniqueProteinSeqDuplicates.txt).

Use /R to rename proteins with duplicate names when using /F to generate a 
fixed fasta file.

Use /D to consolidate proteins with duplicate protein sequences when using 
/F to generate a fixed fasta file.

Use /L to ignore I/L (isoleucine vs. leucine) differences when consolidating 
proteins with duplicate protein sequences while generating a fixed fasta file.

Use /V to remove invalid residues (non-letter characters, including an 
asterisk) when using /F to generate a fixed fasta file.

Use /KeepSameName to keep proteins with the same name but differing sequences 
when using /F to generate a fixed fasta file (if they have the same name and 
same sequence, then will only retain one entry); ignored if /R or /D is used

Use /AllowDash to allow a - in residues
use /AllowAsterisk to allow * in residues

When parsing large fasta files, you can reduce the memory used by disabling 
the checking for duplicates
 /SkipDupeSeqCheck disables duplicate sequence checking (which can 
  require a large amount of memory)
 /SkipDupeNameCheck disables duplicate name checking (which typically 
  requires less memory)

Use /B to save a hash info file (even if not consolidating duplicates).
This is useful for parsing a large fasta file to obtain the sequence hash 
for each protein (hash values are not cached in memory, thus very little
memory is required).

Use /HashFile to specify a pre-computed hash file to use for determining 
which proteins to keep when generating a fixed fasta file. Use of 
/HashFile automatically enables /F and automatically disables /D, /R, and /B

Use /X to specify that a model XML parameter file should be created.
Use /S to process all valid files in the input directory and subdirectories.
Include a number after /S (like /S:2) to limit the level of subdirectories 
to examine.

The optional /Q switch will suppress all error messages.

## Contacts

Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA) \
E-mail: matthew.monroe@pnnl.gov or proteomics@pnnl.gov\
Website: https://github.com/PNNL-Comp-Mass-Spec/ or https://panomics.pnnl.gov/ or https://www.pnnl.gov/integrative-omics

## License

Licensed under the Apache License, Version 2.0; you may not use this program except 
in compliance with the License.  You may obtain a copy of the License at 
http://www.apache.org/licenses/LICENSE-2.0
