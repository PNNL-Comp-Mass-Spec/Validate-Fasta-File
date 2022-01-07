using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using PRISM;

namespace ValidateFastaFile
{
    /// <summary>
    /// This program will read in a FASTA file and write out stats on the number of proteins and number of residues
    /// It will also validate the protein name, descriptions, and sequences in the file
    /// </summary>
    /// <remarks>
    /// <para>
    /// Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
    /// Program started March 21, 2005
    /// </para>
    /// <para>
    /// E-mail: matthew.monroe@pnnl.gov or proteomics@pnnl.gov
    /// Website: https://github.com/PNNL-Comp-Mass-Spec/ or https://panomics.pnnl.gov/ or https://www.pnnl.gov/integrative-omics
    /// </para>
    /// <para>
    /// Licensed under the Apache License, Version 2.0; you may not use this file except
    /// in compliance with the License.  You may obtain a copy of the License at
    /// http://www.apache.org/licenses/LICENSE-2.0
    /// </para>
    /// </remarks>
    internal static class Program
    {
        // Ignore Spelling: isoleucine, leucine, parseable, pre

        public const string PROGRAM_DATE = "January 7, 2022";

        private static string mInputFilePath;
        private static string mOutputDirectoryPath;
        private static string mParameterFilePath;

        private static bool mUseStatsFile;
        private static bool mGenerateFixedFastaFile;

        private static bool mCheckForDuplicateProteinNames;
        private static bool mCheckForDuplicateProteinSequences;

        private static bool mFixedFastaRenameDuplicateNameProteins;
        private static bool mFixedFastaKeepDuplicateNamedProteins;

        private static bool mFixedFastaConsolidateDuplicateProteinSeqs;
        private static bool mFixedFastaConsolidateDupsIgnoreILDiff;
        private static bool mFixedFastaRemoveInvalidResidues;

        private static bool mAllowAsterisk;
        private static bool mAllowDash;

        private static bool mSaveBasicProteinHashInfoFile;
        private static string mProteinHashFilePath;

        private static bool mCreateModelXMLParameterFile;

        private static bool mRecurseDirectories;
        private static int mMaxLevelsToRecurse;

        private static FastaValidator mValidateFastaFile;

        private static DateTime mLastProgressReportPctTime;
        private static DateTime mLastProgressReportTime;
        private static int mLastProgressReportValue;

        /// <summary>
        /// Entry method
        /// </summary>
        /// <returns>0 if no error, error code if an error</returns>
        public static int Main()
        {
            var commandLineParser = new clsParseCommandLine();

            mInputFilePath = string.Empty;
            mOutputDirectoryPath = string.Empty;
            mParameterFilePath = string.Empty;

            mUseStatsFile = false;
            mGenerateFixedFastaFile = false;

            mCheckForDuplicateProteinNames = true;
            mCheckForDuplicateProteinSequences = true;

            mFixedFastaRenameDuplicateNameProteins = false;
            mFixedFastaKeepDuplicateNamedProteins = false;

            mFixedFastaConsolidateDuplicateProteinSeqs = false;
            mFixedFastaConsolidateDupsIgnoreILDiff = false;

            mAllowAsterisk = false;
            mAllowDash = false;

            mSaveBasicProteinHashInfoFile = false;
            mProteinHashFilePath = string.Empty;

            mRecurseDirectories = false;
            mMaxLevelsToRecurse = 0;

            mLastProgressReportPctTime = DateTime.UtcNow;
            mLastProgressReportTime = DateTime.UtcNow;
            try
            {
                var proceed = commandLineParser.ParseCommandLine() && SetOptionsUsingCommandLineParameters(commandLineParser);

                if (proceed && !commandLineParser.NeedToShowHelp && mCreateModelXMLParameterFile)
                {
                    if (string.IsNullOrEmpty(mParameterFilePath))
                    {
                        mParameterFilePath = Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().Location) + "_ModelSettings.xml";
                    }

                    mValidateFastaFile = new FastaValidator();
                    mValidateFastaFile.SaveSettingsToParameterFile(mParameterFilePath);
                    Console.WriteLine();
                    Console.WriteLine("Created example XML parameter file: ");
                    Console.WriteLine("  " + mParameterFilePath);
                    Console.WriteLine();
                    return 0;
                }

                if (!proceed || commandLineParser.NeedToShowHelp || mInputFilePath.Length == 0)
                {
                    ShowProgramHelp();
                    return -1;
                }

                mValidateFastaFile = new FastaValidator();
                mValidateFastaFile.ProgressUpdate += ValidateFastaFile_ProgressChanged;
                mValidateFastaFile.ProgressReset += ValidateFastaFile_ProgressReset;

                mValidateFastaFile.SetOptionSwitch(FastaValidator.SwitchOptions.OutputToStatsFile, mUseStatsFile);
                mValidateFastaFile.SetOptionSwitch(FastaValidator.SwitchOptions.GenerateFixedFASTAFile, mGenerateFixedFastaFile);

                // Also use mGenerateFixedFastaFile to set SaveProteinSequenceHashInfoFiles
                mValidateFastaFile.SetOptionSwitch(FastaValidator.SwitchOptions.SaveProteinSequenceHashInfoFiles, mGenerateFixedFastaFile);

                mValidateFastaFile.SetOptionSwitch(FastaValidator.SwitchOptions.FixedFastaRenameDuplicateNameProteins, mFixedFastaRenameDuplicateNameProteins);
                mValidateFastaFile.SetOptionSwitch(FastaValidator.SwitchOptions.FixedFastaKeepDuplicateNamedProteins, mFixedFastaKeepDuplicateNamedProteins);

                mValidateFastaFile.SetOptionSwitch(FastaValidator.SwitchOptions.FixedFastaConsolidateDuplicateProteinSeqs, mFixedFastaConsolidateDuplicateProteinSeqs);
                mValidateFastaFile.SetOptionSwitch(FastaValidator.SwitchOptions.FixedFastaConsolidateDupsIgnoreILDiff, mFixedFastaConsolidateDupsIgnoreILDiff);

                mValidateFastaFile.SetOptionSwitch(FastaValidator.SwitchOptions.FixedFastaRemoveInvalidResidues, mFixedFastaRemoveInvalidResidues);

                mValidateFastaFile.SetOptionSwitch(FastaValidator.SwitchOptions.AllowAsteriskInResidues, mAllowAsterisk);
                mValidateFastaFile.SetOptionSwitch(FastaValidator.SwitchOptions.AllowDashInResidues, mAllowDash);

                mValidateFastaFile.SetOptionSwitch(FastaValidator.SwitchOptions.SaveBasicProteinHashInfoFile, mSaveBasicProteinHashInfoFile);

                mValidateFastaFile.SetOptionSwitch(FastaValidator.SwitchOptions.CheckForDuplicateProteinNames, mCheckForDuplicateProteinNames);
                mValidateFastaFile.SetOptionSwitch(FastaValidator.SwitchOptions.CheckForDuplicateProteinSequences, mCheckForDuplicateProteinSequences);

                // Update the rules based on the options that were set above
                mValidateFastaFile.SetDefaultRules();

                mValidateFastaFile.ExistingProteinHashFile = mProteinHashFilePath;

                mValidateFastaFile.SkipConsoleWriteIfNoProgressListener = true;

                // Note: the following settings will be overridden if mParameterFilePath points to a valid parameter file that has these settings defined
                // mValidateFastaFile.SetOptionSwitch(FastaValidator.SwitchOptions.AddMissingLineFeedAtEOF, );
                // mValidateFastaFile.SetOptionSwitch(FastaValidator.SwitchOptions.AllowAsteriskInResidues, );
                // mValidateFastaFile.MaximumFileErrorsToTrack();
                // mValidateFastaFile.MinimumProteinNameLength();
                // mValidateFastaFile.MaximumProteinNameLength();
                // mValidateFastaFile.SetOptionSwitch(FastaValidator.SwitchOptions.WarnBlankLinesBetweenProteins, );
                // mValidateFastaFile.SetOptionSwitch(FastaValidator.SwitchOptions.CheckForDuplicateProteinSequences, );
                // mValidateFastaFile.SetOptionSwitch(FastaValidator.SwitchOptions.SaveProteinSequenceHashInfoFiles, )

                int returnCode;
                if (mRecurseDirectories)
                {
                    if (mValidateFastaFile.ProcessFilesAndRecurseDirectories(mInputFilePath, mOutputDirectoryPath, mOutputDirectoryPath, false, mParameterFilePath, mMaxLevelsToRecurse))
                    {
                        returnCode = 0;
                    }
                    else
                    {
                        returnCode = (int)mValidateFastaFile.ErrorCode;
                    }
                }
                else if (mValidateFastaFile.ProcessFilesWildcard(mInputFilePath, mOutputDirectoryPath, mParameterFilePath))
                {
                    returnCode = 0;
                }
                else
                {
                    returnCode = (int)mValidateFastaFile.ErrorCode;
                    if (returnCode != 0)
                    {
                        ShowErrorMessage("Error while processing: " + mValidateFastaFile.GetErrorMessage());
                    }
                }

                DisplayProgressPercent(mLastProgressReportValue, true);

                return returnCode;
            }
            catch (Exception ex)
            {
                ShowErrorMessage("Error occurred in Program->Main: " + ex.Message, ex);
                return -1;
            }
        }

        private static void DisplayProgressPercent(int percentComplete, bool addNewline)
        {
            if (addNewline)
            {
                Console.WriteLine();
            }

            if (percentComplete > 100)
                percentComplete = 100;

            Console.Write("Processing: " + percentComplete.ToString() + "% ");
            if (addNewline)
            {
                Console.WriteLine();
            }
        }

        private static string GetAppVersion()
        {
            return PRISM.FileProcessor.ProcessFilesOrDirectoriesBase.GetAppVersion(PROGRAM_DATE);
        }

        private static bool SetOptionsUsingCommandLineParameters(clsParseCommandLine commandLineParser)
        {
            // Returns True if no problems; otherwise, returns false

            var validParameters = new List<string>()
            {
                "I", "O", "P", "C",
                "SkipDupeNameCheck", "SkipDupeSeqCheck",
                "F", "R", "D", "L", "V",
                "KeepSameName", "AllowDash", "AllowAsterisk",
                "B", "HashFile",
                "X", "S"
            };

            try
            {
                // Make sure no invalid parameters are present
                if (commandLineParser.InvalidParametersPresent(validParameters))
                {
                    ShowErrorMessage("Invalid command line parameters",
                        (from item in commandLineParser.InvalidParameters(validParameters) select ("/" + item)).ToList());
                    return false;
                }

                // Query commandLineParser to see if various parameters are present
                if (commandLineParser.RetrieveValueForParameter("I", out var inputFile))
                {
                    mInputFilePath = inputFile;
                }
                else if (commandLineParser.NonSwitchParameterCount > 0)
                {
                    mInputFilePath = commandLineParser.RetrieveNonSwitchParameter(0);
                }

                if (commandLineParser.RetrieveValueForParameter("O", out var outputDirectory))
                    mOutputDirectoryPath = outputDirectory;

                if (commandLineParser.RetrieveValueForParameter("P", out var parameterFile))
                    mParameterFilePath = parameterFile;

                if (commandLineParser.IsParameterPresent("C"))
                    mUseStatsFile = true;

                if (commandLineParser.IsParameterPresent("SkipDupeNameCheck"))
                    mCheckForDuplicateProteinNames = false;
                if (commandLineParser.IsParameterPresent("SkipDupeSeqCheck"))
                    mCheckForDuplicateProteinSequences = false;

                if (commandLineParser.IsParameterPresent("F"))
                    mGenerateFixedFastaFile = true;

                if (commandLineParser.IsParameterPresent("R"))
                    mFixedFastaRenameDuplicateNameProteins = true;

                if (commandLineParser.IsParameterPresent("D"))
                    mFixedFastaConsolidateDuplicateProteinSeqs = true;

                if (commandLineParser.IsParameterPresent("L"))
                    mFixedFastaConsolidateDupsIgnoreILDiff = true;

                if (commandLineParser.IsParameterPresent("V"))
                    mFixedFastaRemoveInvalidResidues = true;

                if (commandLineParser.IsParameterPresent("KeepSameName"))
                    mFixedFastaKeepDuplicateNamedProteins = true;

                if (commandLineParser.IsParameterPresent("AllowAsterisk"))
                    mAllowAsterisk = true;

                if (commandLineParser.IsParameterPresent("AllowDash"))
                    mAllowDash = true;

                if (commandLineParser.IsParameterPresent("B"))
                    mSaveBasicProteinHashInfoFile = true;

                if (commandLineParser.RetrieveValueForParameter("HashFile", out var createHashFile))
                    mProteinHashFilePath = createHashFile;

                if (commandLineParser.IsParameterPresent("X"))
                    mCreateModelXMLParameterFile = true;

                if (commandLineParser.RetrieveValueForParameter("S", out var recurse))
                {
                    mRecurseDirectories = true;
                    if (int.TryParse(recurse, out var recurseDepth))
                    {
                        mMaxLevelsToRecurse = recurseDepth;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                ShowErrorMessage("Error parsing the command line parameters: " + ex.Message, ex);
            }

            return false;
        }

        private static void ShowErrorMessage(string message, Exception ex = null)
        {
            ConsoleMsgUtils.ShowError(message, ex);
        }

        private static void ShowErrorMessage(string title, IEnumerable<string> items)
        {
            ConsoleMsgUtils.ShowErrors(title, items);
        }

        private static void ShowProgramHelp()
        {
            try
            {
                var exeName = Path.GetFileName(Assembly.GetExecutingAssembly().Location);

                Console.WriteLine("== Overview ==");
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                    "This program will read a FASTA File and display statistics on the number of proteins and number of residues. " +
                    "It will also check that the protein names, descriptions, and sequences are in the correct format."));
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                    "The program can optionally create a new, fixed version of a FASTA file where proteins with duplicate sequences " +
                    "have been consolidated, and proteins with duplicate names have been renamed."));
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                    "To remove duplicates from huge FASTA files (over 1 GB in size), " +
                    "first create the ProteinHashes.txt file by calling this program with:"));
                Console.WriteLine("  {0} Proteins.fasta /B /SkipDupeSeqCheck /SkipDupeNameCheck", exeName);
                Console.WriteLine();
                Console.WriteLine("Next call the program again, providing the name of the ProteinHashes file:");
                Console.WriteLine("  {0} Proteins.fasta /HashFile:Proteins_ProteinHashes.txt", exeName);
                Console.WriteLine();
                Console.WriteLine("== Program syntax ==");
                Console.WriteLine();
                Console.WriteLine(exeName);
                Console.WriteLine(" /I:InputFilePath.fasta [/O:OutputDirectoryPath]");
                Console.WriteLine(" [/P:ParameterFilePath] [/C] ");
                Console.WriteLine(" [/F] [/R] [/D] [/L] [/V] [/KeepSameName]");
                Console.WriteLine(" [/AllowDash] [/AllowAsterisk]");
                Console.WriteLine(" [/SkipDupeNameCheck] [/SkipDupeSeqCheck]");
                Console.WriteLine(" [/B] [/HashFile]");
                Console.WriteLine(" [/X] [/S:[MaxLevel]]");
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                    "The input file path can contain the wildcard character * and should point to a FASTA file."));
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                    "The output directory path is optional, and is only used if /C is used. If omitted, the output stats file " +
                    "will be created in the directory containing the .Exe file."));
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                    "The parameter file path is optional. If included, it should point to a valid XML parameter file."));

                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                    "Use /C to specify that an output file should be created, rather than displaying the results on the screen."));

                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                    "Use /F to generate a new, fixed .Fasta file (long protein names will be auto-shortened). " +
                    "At the same time, a file with protein names and hash values for each unique protein sequences " +
                    "will be generated (_UniqueProteinSeqs.txt). This file will also list the other proteins " +
                    "that have duplicate sequences as the first protein mapped to each sequence. If duplicate sequences " +
                    "are found, then an easily parseable mapping file will also be created (_UniqueProteinSeqDuplicates.txt)."));
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                    "Use /R to rename proteins with duplicate names when using /F to generate a fixed FASTA file."));
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                    "Use /D to consolidate proteins with duplicate protein sequences when using /F to generate a fixed FASTA file."));
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                    "Use /L to ignore I/L (isoleucine vs. leucine) differences when consolidating proteins " +
                    "with duplicate protein sequences while generating a fixed FASTA file."));
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                    "Use /V to remove invalid residues (non-letter characters, including an asterisk) when using /F to generate a fixed FASTA file."));
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                    "Use /KeepSameName to keep proteins with the same name but differing sequences when using /F to generate a fixed FASTA file " +
                    "(if they have the same name and same sequence, then will only retain one entry); ignored if /R or /D is used"));
                Console.WriteLine();
                Console.WriteLine("Use /AllowDash to allow a - in residues");
                Console.WriteLine("Use /AllowAsterisk to allow * in residues");
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                    "When parsing large FASTA files, you can reduce the memory used by disabling the checking for duplicates"));

                Console.WriteLine(" /SkipDupeSeqCheck disables duplicate sequence checking (large memory footprint)");
                Console.WriteLine(" /SkipDupeNameCheck disables duplicate name checking (small memory footprint)");
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                    "Use /B to save a hash info file (even if not consolidating duplicates). " +
                    "This is useful for parsing a large FASTA file to obtain the sequence hash for each protein " +
                    "(hash values are not cached in memory, thus small memory footprint)."));
                Console.WriteLine();
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                    "Use /HashFile to specify a pre-computed hash file to use for determining which proteins to keep when generating a fixed FASTA file"));
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph("Use of /HashFile automatically enables /F and automatically disables /D, /R, and /B"));
                Console.WriteLine();
                Console.WriteLine("Use /X to specify that a model XML parameter file should be created.");
                Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                    "Use /S to process all valid files in the input directory and subdirectories. " +
                    "Include a number after /S (like /S:2) to limit the level of subdirectories to examine."));
                Console.WriteLine();

                Console.WriteLine("Program written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)");
                Console.WriteLine("Version: " + GetAppVersion());
                Console.WriteLine();

                Console.WriteLine("E-mail: matthew.monroe@pnnl.gov or proteomics@pnnl.gov");
                Console.WriteLine("Website: https://github.com/PNNL-Comp-Mass-Spec/ or https://panomics.pnnl.gov/ or https://www.pnnl.gov/integrative-omics");
                Console.WriteLine();

                // Delay for 750 msec in case the user double clicked this file from within Windows Explorer (or started the program via a shortcut)
                Thread.Sleep(750);
            }
            catch (Exception ex)
            {
                ShowErrorMessage("Error displaying the program syntax: " + ex.Message, ex);
            }
        }

        private static void ValidateFastaFile_ProgressChanged(string taskDescription, float percentComplete)
        {
            const int PERCENT_REPORT_INTERVAL = 25;
            const int PROGRESS_DOT_INTERVAL_MSEC = 500;

            if (percentComplete >= mLastProgressReportValue ||
                DateTime.UtcNow.Subtract(mLastProgressReportPctTime).TotalSeconds >= 30)
            {
                mLastProgressReportPctTime = DateTime.UtcNow;

                if (mLastProgressReportValue > 0)
                {
                    Console.WriteLine();
                }

                if (percentComplete < mLastProgressReportValue)
                {
                    DisplayProgressPercent((int)Math.Round(percentComplete, 0), false);
                }
                else
                {
                    DisplayProgressPercent(mLastProgressReportValue, false);
                }

                while (percentComplete >= mLastProgressReportValue)
                {
                    mLastProgressReportValue += PERCENT_REPORT_INTERVAL;
                }

                mLastProgressReportTime = DateTime.UtcNow;
            }
            else if (DateTime.UtcNow.Subtract(mLastProgressReportTime).TotalMilliseconds >= PROGRESS_DOT_INTERVAL_MSEC)
            {
                mLastProgressReportTime = DateTime.UtcNow;
                Console.Write(".");
            }
        }

        private static void ValidateFastaFile_ProgressReset()
        {
            mLastProgressReportTime = DateTime.UtcNow;
            mLastProgressReportPctTime = DateTime.UtcNow;
            mLastProgressReportValue = 0;
        }
    }
}