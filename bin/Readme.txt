ValidateFastaFile.exe is a VB.NET program that will parse a Fasta file and 
return the number of proteins and number of residues in the file.  
Additionally, it will check the validity of the fasta file looking for common, 
known problems.

You can use the /F switch to generate a new, Fixed Fasta file, where various
protein naming issues are fixed.  By default, long protein names will be 
shortened and invalid residues will be removed.  The processing will also
look for proteins with duplicate sequences or duplicate names.   If you use
the /R switch, then proteins with duplicate names (but differing sequences)
will be renamed to assign a unique name to each protein.  Additionally,
if you use the /D switch, then proteins with duplicate sequences will be
consolidated to keep just one copy of each protein (and the protein
description for this master copy will include the protein names of all of 
the duplicate proteins).  When looking for duplicate proteins, you can use
the /L switch to ignore I/L differences. in protein sequences

You can provide an XML parameter file to the program using the /P switch.
Related to this, run the program with the /X switch to have it generate
a model (default) XML parameter file.

-------------------------------------------------------------------------------
Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
Copyright 2005, Battelle Memorial Institute.  All Rights Reserved.

E-mail: matthew.monroe@pnl.gov or matt@alchemistmatt.com
Website: http://ncrr.pnl.gov/ or http://www.sysbio.org/resources/staff/
-------------------------------------------------------------------------------

Licensed under the Apache License, Version 2.0; you may not use this file except 
in compliance with the License.  You may obtain a copy of the License at 
http://www.apache.org/licenses/LICENSE-2.0

Notice: This computer software was prepared by Battelle Memorial Institute, 
hereinafter the Contractor, under Contract No. DE-AC05-76RL0 1830 with the 
Department of Energy (DOE).  All rights in the computer software are reserved 
by DOE on behalf of the United States Government and the Contractor as 
provided in the Contract.  NEITHER THE GOVERNMENT NOR THE CONTRACTOR MAKES ANY 
WARRANTY, EXPRESS OR IMPLIED, OR ASSUMES ANY LIABILITY FOR THE USE OF THIS 
SOFTWARE.  This notice including this sentence must appear on any copies of 
this computer software.
