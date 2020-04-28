using System;
using PRISM;
using ValidateFastaFile;

namespace TestValidateFastaFileDLL
{
    // Program written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
    // Copyright 2005, Battelle Memorial Institute.  All Rights Reserved.

    // This program can be used to test the use of the clsValidateFastaFiles in the ValidateFastaFiles.Dll file
    // Last modified May 8, 2007

    static class Program
    {
        public static int Main()
        {
            // Returns 0 if no error, error code if an error

            var testFilePath = "JunkTest.fasta";

            var returnCode = 0;

            try
            {
                // See if the user provided a custom filepath at the command line
                try
                {
                    // This command will fail if the program is called from a network share
                    var parameters = Environment.GetCommandLineArgs();

                    if (parameters != null && parameters.Length > 1)
                    {
                        // Note that parameters(0) is the path to the Executable for the calling program
                        testFilePath = parameters[1];
                    }
                }
                catch (Exception ex)
                {
                    // Ignore errors here
                }

                Console.WriteLine("Examining file: " + testFilePath);

                var fastaFileValidator = new clsValidateFastaFile();

                fastaFileValidator.SetOptionSwitch(clsValidateFastaFile.SwitchOptions.OutputToStatsFile, true);

                // Note: the following settings will be overridden if parameter file with these settings defined is provided to .ProcessFile()
                fastaFileValidator.SetOptionSwitch(clsValidateFastaFile.SwitchOptions.AddMissingLineFeedAtEOF, false);
                fastaFileValidator.SetOptionSwitch(clsValidateFastaFile.SwitchOptions.AllowAsteriskInResidues, true);

                fastaFileValidator.MaximumFileErrorsToTrack = 5;               // The maximum number of errors for each type of error; the total error count is always available, but detailed information is only saved for this many errors or warnings of each type
                fastaFileValidator.MinimumProteinNameLength = 3;
                fastaFileValidator.MaximumProteinNameLength = 34;

                fastaFileValidator.SetOptionSwitch(clsValidateFastaFile.SwitchOptions.WarnBlankLinesBetweenProteins, false);

                // Analyze the fasta file; returns true if the analysis was successful (even if the file contains errors or warnings)
                var success = fastaFileValidator.ProcessFile(testFilePath, string.Empty);

                if (success)
                {
                    var count = fastaFileValidator.GetErrorWarningCounts(clsValidateFastaFile.eMsgTypeConstants.ErrorMsg, clsValidateFastaFile.ErrorWarningCountTypes.Total);
                    if (count == 0)
                    {
                        Console.WriteLine(" No errors were found");
                    }
                    else
                    {
                        Console.WriteLine(" " + count.ToString() + " errors were found");
                    }

                    count = fastaFileValidator.GetErrorWarningCounts(clsValidateFastaFile.eMsgTypeConstants.WarningMsg, clsValidateFastaFile.ErrorWarningCountTypes.Total);
                    if (count == 0)
                    {
                        Console.WriteLine(" No warnings were found");
                    }
                    else
                    {
                        Console.WriteLine(" " + count.ToString() + " warnings were found");
                    }

                    // ' Could enumerate the errors using the following
                    for (var index = 0; index <= count - 1; index++)
                        Console.WriteLine(fastaFileValidator.GetErrorMessageTextByIndex(index, "\t"));
                }
                else
                {
                    ConsoleMsgUtils.ShowError("Error calling validateFastaFile.ProcessFile: " + fastaFileValidator.GetErrorMessage());
                }
            }
            catch (Exception ex)
            {
                ConsoleMsgUtils.ShowError("Error occurred: " + ex.Message);
                returnCode = -1;
            }

            return returnCode;
        }
    }
}
