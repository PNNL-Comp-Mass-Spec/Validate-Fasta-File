using System.Collections.Generic;
using System.IO;

namespace ValidateFastaFile
{
    public class clsCustomValidateFastaFiles : clsValidateFastaFile
    {
        #region "Structures and enums"
        public class ErrorInfoExtended
        {
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

            public int LineNumber;
            public string ProteinName;
            public string MessageText;
            public string ExtraInfo;
            public string Type;
        }

        public enum eValidationOptionConstants : int
        {
            AllowAsterisksInResidues = 0,
            AllowDashInResidues = 1,
            AllowAllSymbolsInProteinNames = 2
        }

        public enum eValidationMessageTypes : int
        {
            ErrorMsg = 0,
            WarningMsg = 1
        }

        #endregion

        private readonly List<ErrorInfoExtended> m_CurrentFileErrors;
        private readonly List<ErrorInfoExtended> m_CurrentFileWarnings;

        private string m_CachedFastaFilePath;

        // Note: this array gets initialized with space for 10 items
        // If eValidationOptionConstants gets more than 10 entries, then this array will need to be expanded
        private readonly bool[] mValidationOptions;

        /// <summary>
        /// Constructor
        /// </summary>
        public clsCustomValidateFastaFiles() : base()
        {
            FullErrorCollection = new Dictionary<string, List<ErrorInfoExtended>>();
            FullWarningCollection = new Dictionary<string, List<ErrorInfoExtended>>();

            m_CurrentFileErrors = new List<ErrorInfoExtended>();
            m_CurrentFileWarnings = new List<ErrorInfoExtended>();

            // Reserve space for tracking up to 10 validation updates (expand later if needed)
            mValidationOptions = new bool[11];
        }

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

        public bool FASTAFileValid(string FASTAFileName)
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

        public List<ErrorInfoExtended> RecordedFASTAFileErrors(string fastaFileName)
        {
            if (FullErrorCollection.TryGetValue(fastaFileName, out var errorList))
            {
                return errorList;
            }

            return new List<ErrorInfoExtended>();
        }

        public List<ErrorInfoExtended> RecordedFASTAFileWarnings(string fastaFileName)
        {
            if (FullWarningCollection.TryGetValue(fastaFileName, out var warningList))
            {
                return warningList;
            }

            return new List<ErrorInfoExtended>();
        }

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
            eValidationMessageTypes messageType)
        {
            var msgString = LookupMessageDescription(errorMessageCode, extraInfo);

            RecordFastaFileProblemToHash(lineNumber, proteinName, msgString, extraInfo, messageType);
        }

        private void RecordFastaFileProblem(
            int lineNumber,
            string proteinName,
            string errorMessage,
            string extraInfo,
            eValidationMessageTypes messageType)
        {
            RecordFastaFileProblemToHash(lineNumber, proteinName, errorMessage, extraInfo, messageType);
        }

        private void RecordFastaFileProblemToHash(
            int lineNumber,
            string proteinName,
            string messageString,
            string extraInfo,
            eValidationMessageTypes messageType)
        {
            if (!mFastaFilePath.Equals(m_CachedFastaFilePath))
            {
                // New File being analyzed
                m_CurrentFileErrors.Clear();
                m_CurrentFileWarnings.Clear();

                m_CachedFastaFilePath = string.Copy(mFastaFilePath);
            }

            if (messageType == eValidationMessageTypes.WarningMsg)
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

        public void SetValidationOptions(eValidationOptionConstants eValidationOptionName, bool enabled)
        {
            mValidationOptions[(int)eValidationOptionName] = enabled;
        }

        /// <summary>
        /// Calls SimpleProcessFile(), which calls clsValidateFastaFile.ProcessFile to validate filePath
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns>True if the file was successfully processed (even if it contains errors)</returns>
        public bool StartValidateFASTAFile(string filePath)
        {
            var success = SimpleProcessFile(filePath);

            if (success)
            {
                if (GetErrorWarningCounts(eMsgTypeConstants.WarningMsg, ErrorWarningCountTypes.Total) > 0)
                {
                    // The file has warnings; we need to record them using RecordFastaFileProblem

                    var warnings = GetFileWarnings();

                    foreach (var item in warnings)
                        RecordFastaFileProblem(item.LineNumber, item.ProteinName, item.MessageCode, string.Empty, eValidationMessageTypes.WarningMsg);
                }

                if (GetErrorWarningCounts(eMsgTypeConstants.ErrorMsg, ErrorWarningCountTypes.Total) > 0)
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
                                if (mValidationOptions[(int)eValidationOptionConstants.AllowAsterisksInResidues])
                                {
                                    ignoreError = true;
                                }

                                break;

                            case MESSAGE_TEXT_DASH_IN_RESIDUES:
                                if (mValidationOptions[(int)eValidationOptionConstants.AllowDashInResidues])
                                {
                                    ignoreError = true;
                                }

                                break;
                        }

                        if (!ignoreError)
                        {
                            RecordFastaFileProblem(item.LineNumber, item.ProteinName, errorMessage, item.ExtraInfo, eValidationMessageTypes.ErrorMsg);
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