using System;
using System.Collections.Generic;
using System.IO;
using PRISM;
using ValidateFastaFile;

namespace TestValidateFastaFileDLL
{
    /// <summary>
    /// This program can be used to test the use of the ValidateFastaFiles in the ValidateFastaFiles.Dll file
    /// </summary>
    /// <remarks>
    /// Program written by Matthew Monroe in 2005 for the Department of Energy (PNNL, Richland, WA)
    /// </remarks>
    internal static class Program
    {
        /// <summary>
        /// Entry method
        /// </summary>
        /// <returns>0 if no error, error code if an error</returns>
        public static int Main()
        {
            var testFilePaths = new List<string> {"JunkTest.fasta", "JunkTest_UTF8.fasta"};

            try
            {
                // See if the user provided a custom file path at the command line
                try
                {
                    // This command will fail if the program is called from a network share
                    var parameters = Environment.GetCommandLineArgs();

                    if (parameters.Length > 1)
                    {
                        // Note that parameters(0) is the path to the Executable for the calling program
                        testFilePaths.Clear();
                        testFilePaths.Add(parameters[1]);
                    }
                }
                catch (Exception)
                {
                    // Ignore errors here
                }

                foreach (var testFilePath in testFilePaths)
                {
                    TestReader(testFilePath);
                }

                return 0;
            }
            catch (Exception ex)
            {
                ConsoleMsgUtils.ShowError("Error occurred: " + ex.Message);
                return -1;
            }
        }

        private static bool FindInputFile(string testFilePath, out FileInfo fileInfo)
        {
            fileInfo = new FileInfo(testFilePath);
            if (fileInfo.Exists)
                return true;

            var alternateDirectories = new List<string>
            {
                Path.Combine("..", "Docs"), Path.Combine("..", "..", "Docs")
            };

            foreach (var alternateDirectory in alternateDirectories)
            {
                var alternateFile = new FileInfo(Path.Combine(alternateDirectory, fileInfo.Name));
                if (!alternateFile.Exists)
                {
                    continue;
                }

                fileInfo = alternateFile;
                return true;
            }

            return false;
        }

        private static void TestReader(string testFilePath)
        {
            if (!FindInputFile(testFilePath, out var testFile))
                return;

            Console.WriteLine("Examining file: " + testFile.FullName);

            var fastaFileValidator = new FastaValidator();

            fastaFileValidator.SetOptionSwitch(FastaValidator.SwitchOptions.OutputToStatsFile, true);

            // Note: the following settings will be overridden if a parameter file with these settings defined is passed to .ProcessFile()
            fastaFileValidator.SetOptionSwitch(FastaValidator.SwitchOptions.AddMissingLineFeedAtEOF, false);
            fastaFileValidator.SetOptionSwitch(FastaValidator.SwitchOptions.AllowAsteriskInResidues, true);

            fastaFileValidator.MaximumFileErrorsToTrack = 5;               // The maximum number of errors for each type of error; the total error count is always available, but detailed information is only saved for this many errors or warnings of each type
            fastaFileValidator.MinimumProteinNameLength = 3;
            fastaFileValidator.MaximumProteinNameLength = 34;

            fastaFileValidator.SetOptionSwitch(FastaValidator.SwitchOptions.WarnBlankLinesBetweenProteins, false);

            fastaFileValidator.SetOptionSwitch(FastaValidator.SwitchOptions.SaveProteinSequenceHashInfoFiles, true);

            // Analyze the FASTA file; returns true if the analysis was successful (even if the file contains errors or warnings)
            var success = fastaFileValidator.ProcessFile(testFile.FullName, string.Empty);

            if (success)
            {
                var errorCount = fastaFileValidator.GetErrorWarningCounts(FastaValidator.MsgTypeConstants.ErrorMsg, FastaValidator.ErrorWarningCountTypes.Total);
                if (errorCount == 0)
                {
                    Console.WriteLine(" No errors were found");
                }
                else
                {
                    Console.WriteLine(" {0} errors were found", errorCount);
                }

                var warningCount = fastaFileValidator.GetErrorWarningCounts(FastaValidator.MsgTypeConstants.WarningMsg, FastaValidator.ErrorWarningCountTypes.Total);
                if (warningCount == 0)
                {
                    Console.WriteLine(" No warnings were found");
                }
                else
                {
                    Console.WriteLine(" {0} warnings were found", warningCount);
                }

                Console.WriteLine();

                // Enumerate the errors
                for (var index = 0; index < errorCount; index++)
                {
                    Console.WriteLine(fastaFileValidator.GetErrorMessageTextByIndex(index, "\t"));
                }

                Console.WriteLine();

                // Enumerate the warnings
                for (var index = 0; index < warningCount; index++)
                {
                    Console.WriteLine(fastaFileValidator.GetWarningMessageTextByIndex(index, "\t"));
                }
            }
            else
            {
                ConsoleMsgUtils.ShowError("Error calling validateFastaFile.ProcessFile: " + fastaFileValidator.GetErrorMessage());
            }
        }
    }
}
