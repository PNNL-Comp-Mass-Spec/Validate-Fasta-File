using System;
using System.Collections.Generic;
using System.IO;

// ReSharper disable UnusedMember.Global

namespace ValidateFastaFile
{
    // Ignore Spelling: Validator

    /// <summary>
    /// Old class name
    /// </summary>
    [Obsolete("Renamed to 'CustomFastaValidator'", true)]
    // ReSharper disable once InconsistentNaming
    public class clsCustomValidateFastaFiles : CustomFastaValidator
    {
        // Intentionally empty
    }

    /// <summary>
    /// Custom FASTA validator
    /// </summary>
    /// <remarks>
    /// Inherits from FastaValidator and has an extended error info class
    /// Supports tracking errors for multiple FASTA files
    /// </remarks>
    public class CustomFastaValidator : FastaValidator
    {
        /// <summary>
        /// Extended error info
        /// </summary>
        public class ErrorInfoExtended
        {
            /// <summary>
            /// Constructor
            /// </summary>
            /// <param name="lineNumber"></param>
            /// <param name="proteinName"></param>
            /// <param name="messageText"></param>
            /// <param name="extraInfo"></param>
            /// <param name="type"></param>
            public ErrorInfoExtended(
                int lineNumber,
                string proteinName,
                string messageText,
                string extraInfo,
                string type)
            {
                LineNumber = lineNumber;
                ProteinName = proteinName;
                MessageText = messageText;
                ExtraInfo = extraInfo;
                Type = type;
            }

            /// <summary>
            /// Line number of the error
            /// </summary>
            public int LineNumber;

            /// <summary>
            /// Protein name
            /// </summary>
            public string ProteinName;

            /// <summary>
            /// Message text
            /// </summary>
            public string MessageText;

            /// <summary>
            /// Extra info
            /// </summary>
            public string ExtraInfo;

            /// <summary>
            /// Error type (Error or Warning)
            /// </summary>
            public string Type;
        }

        /// <summary>
        /// Validation options
        /// </summary>
        public enum ValidationOptionConstants
        {
            /// <summary>
            /// Allow asterisks in residues
            /// </summary>
            AllowAsterisksInResidues = 0,

            /// <summary>
            /// Allow dashes in residues
            /// </summary>
            AllowDashInResidues = 1,

            /// <summary>
            /// Allow symbols in protein names
            /// </summary>
            AllowAllSymbolsInProteinNames = 2
        }

        /// <summary>
        /// Validation message types
        /// </summary>
        public enum ValidationMessageTypes
        {
            /// <summary>
            /// Error message
            /// </summary>
            ErrorMsg = 0,

            /// <summary>
            /// Warning message
            /// </summary>
            WarningMsg = 1
        }

        private readonly List<ErrorInfoExtended> m_CurrentFileErrors;
        private readonly List<ErrorInfoExtended> m_CurrentFileWarnings;

        private string m_CachedFastaFilePath;

        // Note: this array gets initialized with space for 10 items
        // If ValidationOptionConstants gets more than 10 entries, then this array will need to be expanded
        private readonly bool[] mValidationOptions;

        /// <summary>
        /// Constructor
        /// </summary>
        public CustomFastaValidator() : base()
        {
            FullErrorCollection = new Dictionary<string, List<ErrorInfoExtended>>();
            FullWarningCollection = new Dictionary<string, List<ErrorInfoExtended>>();

            m_CurrentFileErrors = new List<ErrorInfoExtended>();
            m_CurrentFileWarnings = new List<ErrorInfoExtended>();

            // Reserve space for tracking up to 10 validation updates (expand later if needed)
            mValidationOptions = new bool[11];
        }

        /// <summary>
        /// Clear cached errors
        /// </summary>
        public void ClearErrorList()
        {
            FullErrorCollection?.Clear();
            FullWarningCollection?.Clear();
        }

        /// <summary>
        /// Keys are fasta filename
        /// Values are the list of fasta file errors
        /// </summary>
        public Dictionary<string, List<ErrorInfoExtended>> FullErrorCollection { get; }

        /// <summary>
        /// Keys are fasta filename
        /// Values are the list of fasta file warnings
        /// </summary>
        public Dictionary<string, List<ErrorInfoExtended>> FullWarningCollection { get; }

        /// <summary>
        /// Returns true if dictionary FullErrorCollection contains any errors for the given FASTA file
        /// </summary>
        /// <param name="fastaFileName"></param>
        public bool FASTAFileValid(string fastaFileName)
        {
            if (FullErrorCollection == null)
            {
                return true;
            }
            else
            {
                return !FullErrorCollection.ContainsKey(FASTAFileName);
            }
        }

        /// <summary>
        /// Returns true if dictionary FullWarningCollection contains any warnings for the given FASTA file
        /// </summary>
        /// <param name="fastaFileName"></param>
        public bool FASTAFileHasWarnings(string fastaFileName)
        {
            if (FullWarningCollection == null)
            {
                return false;
            }
            else
            {
                return FullWarningCollection.ContainsKey(fastaFileName);
            }
        }

        /// <summary>
        /// Returns the errors found for the given FASTA file
        /// </summary>
        /// <param name="fastaFileName"></param>
        public List<ErrorInfoExtended> RecordedFASTAFileErrors(string fastaFileName)
        {
            if (FullErrorCollection.TryGetValue(fastaFileName, out var errorList))
            {
                return errorList;
            }

            return new List<ErrorInfoExtended>();
        }

        /// <summary>
        /// Returns the warnings found for the given FASTA file
        /// </summary>
        /// <param name="fastaFileName"></param>
        public List<ErrorInfoExtended> RecordedFASTAFileWarnings(string fastaFileName)
        {
            if (FullWarningCollection.TryGetValue(fastaFileName, out var warningList))
            {
                return warningList;
            }

            return new List<ErrorInfoExtended>();
        }

        /// <summary>
        /// Return a count of the number of files with an error
        /// </summary>
        public int NumFilesWithErrors
        {
            get
            {
                if (FullWarningCollection == null)
                {
                    return 0;
                }
                else
                {
                    return FullErrorCollection.Count;
                }
            }
        }

        private void RecordFastaFileProblem(
            int lineNumber,
            string proteinName,
            int errorMessageCode,
            string extraInfo,
            ValidationMessageTypes messageType)
        {
            var msgString = LookupMessageDescription(errorMessageCode, extraInfo);

            RecordFastaFileProblemToHash(lineNumber, proteinName, msgString, extraInfo, messageType);
        }

        private void RecordFastaFileProblem(
            int lineNumber,
            string proteinName,
            string errorMessage,
            string extraInfo,
            ValidationMessageTypes messageType)
        {
            RecordFastaFileProblemToHash(lineNumber, proteinName, errorMessage, extraInfo, messageType);
        }

        private void RecordFastaFileProblemToHash(
            int lineNumber,
            string proteinName,
            string messageString,
            string extraInfo,
            ValidationMessageTypes messageType)
        {
            if (!mFastaFilePath.Equals(m_CachedFastaFilePath))
            {
                // New File being analyzed
                m_CurrentFileErrors.Clear();
                m_CurrentFileWarnings.Clear();

                m_CachedFastaFilePath = string.Copy(mFastaFilePath);
            }

            if (messageType == ValidationMessageTypes.WarningMsg)
            {
                // Treat as warning
                m_CurrentFileWarnings.Add(new ErrorInfoExtended(
                    lineNumber, proteinName, messageString, extraInfo, "Warning"));

                FullWarningCollection[Path.GetFileName(m_CachedFastaFilePath)] = m_CurrentFileWarnings;
            }
            else
            {
                // Treat as error
                m_CurrentFileErrors.Add(new ErrorInfoExtended(
                    lineNumber, proteinName, messageString, extraInfo, "Error"));

                FullErrorCollection[Path.GetFileName(m_CachedFastaFilePath)] = m_CurrentFileErrors;
            }
        }

        /// <summary>
        /// Set validation options
        /// </summary>
        /// <param name="eValidationOptionName"></param>
        /// <param name="enabled"></param>
        public void SetValidationOptions(ValidationOptionConstants eValidationOptionName, bool enabled)
        {
            mValidationOptions[(int)eValidationOptionName] = enabled;
        }

        /// <summary>
        /// Calls SimpleProcessFile(), which calls ValidateFastaFile.ProcessFile to validate filePath
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns>True if the file was successfully processed (even if it contains errors)</returns>
        public bool StartValidateFASTAFile(string filePath)
        {
            var success = SimpleProcessFile(filePath);

            if (success)
            {
                if (GetErrorWarningCounts(MsgTypeConstants.WarningMsg, ErrorWarningCountTypes.Total) > 0)
                {
                    // The file has warnings; we need to record them using RecordFastaFileProblem

                    var warnings = GetFileWarnings();

                    foreach (var item in warnings)
                        RecordFastaFileProblem(item.LineNumber, item.ProteinName, item.MessageCode, string.Empty, ValidationMessageTypes.WarningMsg);
                }

                if (GetErrorWarningCounts(MsgTypeConstants.ErrorMsg, ErrorWarningCountTypes.Total) > 0)
                {
                    // The file has errors; we need to record them using RecordFastaFileProblem
                    // However, we might ignore some of the errors

                    var errors = GetFileErrors();

                    foreach (var item in errors)
                    {
                        var errorMessage = LookupMessageDescription(item.MessageCode, item.ExtraInfo);

                        var ignoreError = false;
                        switch (errorMessage)
                        {
                            case MESSAGE_TEXT_ASTERISK_IN_RESIDUES:
                                if (mValidationOptions[(int)ValidationOptionConstants.AllowAsterisksInResidues])
                                {
                                    ignoreError = true;
                                }

                                break;

                            case MESSAGE_TEXT_DASH_IN_RESIDUES:
                                if (mValidationOptions[(int)ValidationOptionConstants.AllowDashInResidues])
                                {
                                    ignoreError = true;
                                }

                                break;
                        }

                        if (!ignoreError)
                        {
                            RecordFastaFileProblem(item.LineNumber, item.ProteinName, errorMessage, item.ExtraInfo, ValidationMessageTypes.ErrorMsg);
                        }
                    }
                }

                return true;
            }
            else
            {
                // SimpleProcessFile returned False
                return false;
            }
        }
    }
}