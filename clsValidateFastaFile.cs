﻿// This class will read a protein fasta file and validate its contents
//
// -------------------------------------------------------------------------------
// Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
// Program started March 21, 2005
//
// E-mail: matthew.monroe@pnnl.gov or proteomics@pnnl.gov
// Website: https://omics.pnl.gov/ or https://panomics.pnnl.gov/
// -------------------------------------------------------------------------------
//
// Licensed under the Apache License, Version 2.0; you may not use this file except
// in compliance with the License.  You may obtain a copy of the License at
// http://www.apache.org/licenses/LICENSE-2.0

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using PRISM;

namespace ValidateFastaFile
{
    public class clsValidateFastaFile : PRISM.FileProcessor.ProcessFilesBase
    {
        /// <summary>
        /// Constructor
        /// </summary>
        public clsValidateFastaFile()
        {
            mFileDate = "April 15, 2020";
            InitializeLocalVariables();
        }

        /// <summary>
        /// Constructor that takes a parameter file
        /// </summary>
        /// <param name="parameterFilePath"></param>
        public clsValidateFastaFile(string parameterFilePath) : this()
        {
            LoadParameterFileSettings(parameterFilePath);
        }

        #region "Constants and Enums"
        private const int DEFAULT_MINIMUM_PROTEIN_NAME_LENGTH = 3;

        /// <summary>
        /// The maximum suggested value when using SEQUEST is 34 characters
        /// In contrast, MS-GF+ supports long protein names
        /// </summary>
        /// <remarks></remarks>
        public const int DEFAULT_MAXIMUM_PROTEIN_NAME_LENGTH = 60;
        private const int DEFAULT_MAXIMUM_RESIDUES_PER_LINE = 120;

        public const char DEFAULT_PROTEIN_LINE_START_CHAR = '>';
        public const char DEFAULT_LONG_PROTEIN_NAME_SPLIT_CHAR = '|';
        public const string DEFAULT_PROTEIN_NAME_FIRST_REF_SEP_CHARS = ":|";
        public const string DEFAULT_PROTEIN_NAME_SUBSEQUENT_REF_SEP_CHARS = ":|;";

        private const char INVALID_PROTEIN_NAME_CHAR_REPLACEMENT = '_';

        private const int CUSTOM_RULE_ID_START = 1000;
        private const int DEFAULT_CONTEXT_LENGTH = 13;

        public const string MESSAGE_TEXT_PROTEIN_DESCRIPTION_MISSING = "Line contains a protein name, but not a description";
        public const string MESSAGE_TEXT_PROTEIN_DESCRIPTION_TOO_LONG = "Protein description is over 900 characters long";
        public const string MESSAGE_TEXT_ASTERISK_IN_RESIDUES = "An asterisk was found in the residues";
        public const string MESSAGE_TEXT_DASH_IN_RESIDUES = "A dash was found in the residues";

        public const string XML_SECTION_OPTIONS = "ValidateFastaFileOptions";
        public const string XML_SECTION_FIXED_FASTA_FILE_OPTIONS = "ValidateFastaFixedFASTAFileOptions";

        public const string XML_SECTION_FASTA_HEADER_LINE_RULES = "ValidateFastaHeaderLineRules";
        public const string XML_SECTION_FASTA_PROTEIN_NAME_RULES = "ValidateFastaProteinNameRules";
        public const string XML_SECTION_FASTA_PROTEIN_DESCRIPTION_RULES = "ValidateFastaProteinDescriptionRules";
        public const string XML_SECTION_FASTA_PROTEIN_SEQUENCE_RULES = "ValidateFastaProteinSequenceRules";

        public const string XML_OPTION_ENTRY_RULE_COUNT = "RuleCount";

        // The value of 7995 is chosen because the maximum varchar() value in Sql Server is varchar(8000)
        // and we want to prevent truncation errors when importing protein names and descriptions into Sql Server
        public const int MAX_PROTEIN_DESCRIPTION_LENGTH = 7995;

        private const string MEM_USAGE_PREFIX = "MemUsage: ";
        private const bool REPORT_DETAILED_MEMORY_USAGE = false;

        private const string PROTEIN_NAME_COLUMN = "Protein_Name";
        private const string SEQUENCE_LENGTH_COLUMN = "Sequence_Length";
        private const string SEQUENCE_HASH_COLUMN = "Sequence_Hash";
        private const string PROTEIN_HASHES_FILENAME_SUFFIX = "_ProteinHashes.txt";

        private const int DEFAULT_WARNING_SEVERITY = 3;
        private const int DEFAULT_ERROR_SEVERITY = 7;

        // Note: Custom rules start with message code CUSTOM_RULE_ID_START=1000, and therefore
        // the values in enum eMessageCodeConstants should all be less than CUSTOM_RULE_ID_START
        public enum eMessageCodeConstants
        {
            UnspecifiedError = 0,

            // Error messages
            ProteinNameIsTooLong = 1,
            LineStartsWithSpace = 2,
            // RightArrowFollowedBySpace = 3
            // RightArrowFollowedByTab = 4
            // RightArrowButNoProteinName = 5
            BlankLineBetweenProteinNameAndResidues = 6,
            BlankLineInMiddleOfResidues = 7,
            ResiduesFoundWithoutProteinHeader = 8,
            ProteinEntriesNotFound = 9,
            FinalProteinEntryMissingResidues = 10,
            FileDoesNotEndWithLinefeed = 11,
            DuplicateProteinName = 12,

            // Warning messages
            ProteinNameIsTooShort = 13,
            // ProteinNameContainsVerticalBars = 14
            // ProteinNameContainsWarningCharacters = 21
            // ProteinNameWithoutDescription = 14
            BlankLineBeforeProteinName = 15,
            // ProteinNameAndDescriptionSeparatedByTab = 16
            // ProteinDescriptionWithTab = 25
            // ProteinDescriptionWithQuotationMark = 26
            // ProteinDescriptionWithEscapedSlash = 27
            // ProteinDescriptionWithUndesirableCharacter = 28
            ResiduesLineTooLong = 17,
            // ResiduesLineContainsU = 30
            DuplicateProteinSequence = 18,
            RenamedProtein = 19,
            ProteinRemovedSinceDuplicateSequence = 20,
            DuplicateProteinNameRetained = 21
        }

        public struct udtMsgInfoType
        {
            public int LineNumber;
            public int ColNumber;
            public string ProteinName;
            public int MessageCode;
            public string ExtraInfo;
            public string Context;

            public override string ToString()
            {
                return string.Format("Line {0}, protein {1}, code {2}: {3}", LineNumber, ProteinName, MessageCode, ExtraInfo);
            }
        }

        public struct udtOutputOptionsType
        {
            public string SourceFile;
            public bool OutputToStatsFile;
            public StreamWriter OutFile;
            public string SepChar;

            public override string ToString()
            {
                return SourceFile;
            }
        }

        public enum RuleTypes
        {
            HeaderLine,
            ProteinName,
            ProteinDescription,
            ProteinSequence
        }

        public enum SwitchOptions
        {
            AddMissingLineFeedAtEOF,
            AllowAsteriskInResidues,
            CheckForDuplicateProteinNames,
            GenerateFixedFASTAFile,
            SplitOutMultipleRefsInProteinName,
            OutputToStatsFile,
            WarnBlankLinesBetweenProteins,
            WarnLineStartsWithSpace,
            NormalizeFileLineEndCharacters,
            CheckForDuplicateProteinSequences,
            FixedFastaRenameDuplicateNameProteins,
            SaveProteinSequenceHashInfoFiles,
            FixedFastaConsolidateDuplicateProteinSeqs,
            FixedFastaConsolidateDupsIgnoreILDiff,
            FixedFastaTruncateLongProteinNames,
            FixedFastaSplitOutMultipleRefsForKnownAccession,
            FixedFastaWrapLongResidueLines,
            FixedFastaRemoveInvalidResidues,
            SaveBasicProteinHashInfoFile,
            AllowDashInResidues,
            FixedFastaKeepDuplicateNamedProteins,        // Keep duplicate named proteins, unless the name and sequence match exactly, then they're removed
            AllowAllSymbolsInProteinNames
        }

        public enum FixedFASTAFileValues
        {
            DuplicateProteinNamesSkippedCount,
            ProteinNamesInvalidCharsReplaced,
            ProteinNamesMultipleRefsRemoved,
            TruncatedProteinNameCount,
            UpdatedResidueLines,
            DuplicateProteinNamesRenamedCount,
            DuplicateProteinSeqsSkippedCount
        }

        public enum ErrorWarningCountTypes
        {
            Specified,
            Unspecified,
            Total
        }

        public enum eMsgTypeConstants
        {
            ErrorMsg = 0,
            WarningMsg = 1,
            StatusMsg = 2
        }

        public enum eValidateFastaFileErrorCodes
        {
            NoError = 0,
            OptionsSectionNotFound = 1,
            ErrorReadingInputFile = 2,
            ErrorCreatingStatsFile = 4,
            ErrorVerifyingLinefeedAtEOF = 8,
            UnspecifiedError = -1
        }

        public enum eLineEndingCharacters
        {
            CRLF,  // Windows
            CR,    // Old Style Mac
            LF,    // Unix/Linux/OS X
            LFCR,  // Oddball (Just for completeness!)
        }
        #endregion

        #region "Structures"

        private struct udtErrorStatsType
        {
            public int MessageCode;               // Note: Custom rules start with message code CUSTOM_RULE_ID_START
            public int CountSpecified;
            public int CountUnspecified;

            public override string ToString()
            {
                return MessageCode + ": " + CountSpecified + " specified, " + CountUnspecified + " unspecified";
            }
        }

        private struct udtItemSummaryIndexedType
        {
            public int ErrorStatsCount;
            public udtErrorStatsType[] ErrorStats;        // Note: This array ranges from 0 to .ErrorStatsCount since it is Dimmed with extra space
            public Dictionary<int, int> MessageCodeToArrayIndex;
        }

        public struct udtRuleDefinitionType
        {
            public string MatchRegEx;
            public bool MatchIndicatesProblem;        // True means text matching the RegEx means a problem; false means if text doesn't match the RegEx, then that means a problem
            public string MessageWhenProblem;         // Message to display if a problem is present
            public short Severity;                    // 0 is lowest severity, 9 is highest severity; value >= 5 means error
            public bool DisplayMatchAsExtraInfo;      // If true, then the matching text is stored as the context info
            public int CustomRuleID;                  // This value is auto-assigned

            public override string ToString()
            {
                return CustomRuleID + ": " + MessageWhenProblem;
            }
        }

        private struct udtRuleDefinitionExtendedType
        {
            public udtRuleDefinitionType RuleDefinition;
            public Regex reRule;

            // ReSharper disable once NotAccessedField.Local
            public bool Valid;

            public override string ToString()
            {
                return RuleDefinition.CustomRuleID + ": " + RuleDefinition.MessageWhenProblem;
            }
        }

        private struct udtFixedFastaOptionsType
        {
            public bool SplitOutMultipleRefsInProteinName;
            public bool SplitOutMultipleRefsForKnownAccession;
            public char[] LongProteinNameSplitChars;
            public char[] ProteinNameInvalidCharsToRemove;
            public bool RenameProteinsWithDuplicateNames;
            public bool KeepDuplicateNamedProteinsUnlessMatchingSequence;      // Ignored if RenameProteinsWithDuplicateNames=true or ConsolidateProteinsWithDuplicateSeqs=true
            public bool ConsolidateProteinsWithDuplicateSeqs;
            public bool ConsolidateDupsIgnoreILDiff;
            public bool TruncateLongProteinNames;
            public bool WrapLongResidueLines;
            public bool RemoveInvalidResidues;
        }

        private struct udtFixedFastaStatsType
        {
            public int TruncatedProteinNameCount;
            public int UpdatedResidueLines;
            public int ProteinNamesInvalidCharsReplaced;
            public int ProteinNamesMultipleRefsRemoved;
            public int DuplicateNameProteinsSkipped;
            public int DuplicateNameProteinsRenamed;
            public int DuplicateSequenceProteinsSkipped;
        }

        private struct udtProteinNameTruncationRegex
        {
            public Regex reMatchIPI;
            public Regex reMatchGI;
            public Regex reMatchJGI;
            public Regex reMatchJGIBaseAndID;
            public Regex reMatchGeneric;
            public Regex reMatchDoubleBarOrColonAndBar;
        }

        #endregion

        #region "Classwide Variables"

        /// <summary>
        /// Fasta file path being examined
        /// </summary>
        /// <remarks>Used by clsCustomValidateFastaFiles</remarks>
        protected string mFastaFilePath;

        private int mLineCount;
        private int mProteinCount;
        private long mResidueCount;

        private udtFixedFastaStatsType mFixedFastaStats;

        private int mFileErrorCount;
        private udtMsgInfoType[] mFileErrors;
        private udtItemSummaryIndexedType mFileErrorStats;

        private int mFileWarningCount;

        private udtMsgInfoType[] mFileWarnings;
        private udtItemSummaryIndexedType mFileWarningStats;

        private udtRuleDefinitionType[] mHeaderLineRules;
        private udtRuleDefinitionType[] mProteinNameRules;
        private udtRuleDefinitionType[] mProteinDescriptionRules;
        private udtRuleDefinitionType[] mProteinSequenceRules;
        private int mMasterCustomRuleID = CUSTOM_RULE_ID_START;

        private char[] mProteinNameFirstRefSepChars;
        private char[] mProteinNameSubsequentRefSepChars;

        private bool mAddMissingLinefeedAtEOF;
        private bool mCheckForDuplicateProteinNames;

        // This will be set to True if
        // mSaveProteinSequenceHashInfoFiles = True or mFixedFastaOptions.ConsolidateProteinsWithDuplicateSeqs = True
        private bool mCheckForDuplicateProteinSequences;

        private int mMaximumFileErrorsToTrack;        // This is the maximum # of errors per type to track
        private int mMinimumProteinNameLength;
        private int mMaximumProteinNameLength;
        private int mMaximumResiduesPerLine;

        private udtFixedFastaOptionsType mFixedFastaOptions;     // Used if mGenerateFixedFastaFile = True

        private bool mOutputToStatsFile;
        private string mStatsFilePath;

        private bool mGenerateFixedFastaFile;
        private bool mSaveProteinSequenceHashInfoFiles;

        // When true, creates a text file that will contain the protein name and sequence hash for each protein;
        // this option will not store protein names and/or hashes in memory, and is thus useful for processing
        // huge .Fasta files to determine duplicate proteins
        private bool mSaveBasicProteinHashInfoFile;

        private char mProteinLineStartChar;

        private bool mAllowAsteriskInResidues;
        private bool mAllowDashInResidues;
        private bool mAllowAllSymbolsInProteinNames;

        private bool mWarnBlankLinesBetweenProteins;
        private bool mWarnLineStartsWithSpace;
        private bool mNormalizeFileLineEndCharacters;

        /// <summary>
        /// The number of characters at the start of key strings to use when adding items to clsNestedStringDictionary instances
        /// </summary>
        /// <remarks>
        /// If this value is too short, all of the items added to the clsNestedStringDictionary instance
        /// will be tracked by the same dictionary, which could result in a dictionary surpassing the 2 GB boundary
        /// </remarks>
        private byte mProteinNameSpannerCharLength = 1;

        private eValidateFastaFileErrorCodes mLocalErrorCode;

        private clsMemoryUsageLogger mMemoryUsageLogger;

        private float mProcessMemoryUsageMBAtStart;

        private string mSortUtilityErrorMessage;

        private List<string> mTempFilesToDelete;

        #endregion

        #region "Properties"

        /// <summary>
        /// Gets or sets a processing option
        /// </summary>
        /// <param name="SwitchName"></param>
        /// <value></value>
        /// <returns></returns>
        /// <remarks>Be sure to call SetDefaultRules() after setting all of the options</remarks>
        [Obsolete("Use GetOptionSwitchValue instead", true)]
        public bool get_OptionSwitch(SwitchOptions switchName)
        {
            return GetOptionSwitchValue(switchName);
        }

        [Obsolete("Use SetOptionSwitch instead", true)]
        public void set_OptionSwitch(SwitchOptions switchName, bool value)
        {
            SetOptionSwitch(switchName, value);
        }

        /// <summary>
        /// Set a processing option
        /// </summary>
        /// <param name="SwitchName"></param>
        /// <param name="State"></param>
        /// <remarks>Be sure to call SetDefaultRules() after setting all of the options</remarks>
        public void SetOptionSwitch(SwitchOptions switchName, bool state)
        {
            switch (switchName)
            {
                case SwitchOptions.AddMissingLineFeedAtEOF:
                    mAddMissingLinefeedAtEOF = state;
                    break;
                case SwitchOptions.AllowAsteriskInResidues:
                    mAllowAsteriskInResidues = state;
                    break;
                case SwitchOptions.CheckForDuplicateProteinNames:
                    mCheckForDuplicateProteinNames = state;
                    break;
                case SwitchOptions.GenerateFixedFASTAFile:
                    mGenerateFixedFastaFile = state;
                    break;
                case SwitchOptions.OutputToStatsFile:
                    mOutputToStatsFile = state;
                    break;
                case SwitchOptions.SplitOutMultipleRefsInProteinName:
                    mFixedFastaOptions.SplitOutMultipleRefsInProteinName = state;
                    break;
                case SwitchOptions.WarnBlankLinesBetweenProteins:
                    mWarnBlankLinesBetweenProteins = state;
                    break;
                case SwitchOptions.WarnLineStartsWithSpace:
                    mWarnLineStartsWithSpace = state;
                    break;
                case SwitchOptions.NormalizeFileLineEndCharacters:
                    mNormalizeFileLineEndCharacters = state;
                    break;
                case SwitchOptions.CheckForDuplicateProteinSequences:
                    mCheckForDuplicateProteinSequences = state;
                    break;
                case SwitchOptions.FixedFastaRenameDuplicateNameProteins:
                    mFixedFastaOptions.RenameProteinsWithDuplicateNames = state;
                    break;
                case SwitchOptions.FixedFastaKeepDuplicateNamedProteins:
                    mFixedFastaOptions.KeepDuplicateNamedProteinsUnlessMatchingSequence = state;
                    break;
                case SwitchOptions.SaveProteinSequenceHashInfoFiles:
                    mSaveProteinSequenceHashInfoFiles = state;
                    break;
                case SwitchOptions.FixedFastaConsolidateDuplicateProteinSeqs:
                    mFixedFastaOptions.ConsolidateProteinsWithDuplicateSeqs = state;
                    break;
                case SwitchOptions.FixedFastaConsolidateDupsIgnoreILDiff:
                    mFixedFastaOptions.ConsolidateDupsIgnoreILDiff = state;
                    break;
                case SwitchOptions.FixedFastaTruncateLongProteinNames:
                    mFixedFastaOptions.TruncateLongProteinNames = state;
                    break;
                case SwitchOptions.FixedFastaSplitOutMultipleRefsForKnownAccession:
                    mFixedFastaOptions.SplitOutMultipleRefsForKnownAccession = state;
                    break;
                case SwitchOptions.FixedFastaWrapLongResidueLines:
                    mFixedFastaOptions.WrapLongResidueLines = state;
                    break;
                case SwitchOptions.FixedFastaRemoveInvalidResidues:
                    mFixedFastaOptions.RemoveInvalidResidues = state;
                    break;
                case SwitchOptions.SaveBasicProteinHashInfoFile:
                    mSaveBasicProteinHashInfoFile = state;
                    break;
                case SwitchOptions.AllowDashInResidues:
                    mAllowDashInResidues = state;
                    break;
                case SwitchOptions.AllowAllSymbolsInProteinNames:
                    mAllowAllSymbolsInProteinNames = state;
                    break;
            }
        }

        public bool GetOptionSwitchValue(SwitchOptions SwitchName)
        {
            switch (SwitchName)
            {
                case SwitchOptions.AddMissingLineFeedAtEOF:
                    return mAddMissingLinefeedAtEOF;
                case SwitchOptions.AllowAsteriskInResidues:
                    return mAllowAsteriskInResidues;
                case SwitchOptions.CheckForDuplicateProteinNames:
                    return mCheckForDuplicateProteinNames;
                case SwitchOptions.GenerateFixedFASTAFile:
                    return mGenerateFixedFastaFile;
                case SwitchOptions.OutputToStatsFile:
                    return mOutputToStatsFile;
                case SwitchOptions.SplitOutMultipleRefsInProteinName:
                    return mFixedFastaOptions.SplitOutMultipleRefsInProteinName;
                case SwitchOptions.WarnBlankLinesBetweenProteins:
                    return mWarnBlankLinesBetweenProteins;
                case SwitchOptions.WarnLineStartsWithSpace:
                    return mWarnLineStartsWithSpace;
                case SwitchOptions.NormalizeFileLineEndCharacters:
                    return mNormalizeFileLineEndCharacters;
                case SwitchOptions.CheckForDuplicateProteinSequences:
                    return mCheckForDuplicateProteinSequences;
                case SwitchOptions.FixedFastaRenameDuplicateNameProteins:
                    return mFixedFastaOptions.RenameProteinsWithDuplicateNames;
                case SwitchOptions.FixedFastaKeepDuplicateNamedProteins:
                    return mFixedFastaOptions.KeepDuplicateNamedProteinsUnlessMatchingSequence;
                case SwitchOptions.SaveProteinSequenceHashInfoFiles:
                    return mSaveProteinSequenceHashInfoFiles;
                case SwitchOptions.FixedFastaConsolidateDuplicateProteinSeqs:
                    return mFixedFastaOptions.ConsolidateProteinsWithDuplicateSeqs;
                case SwitchOptions.FixedFastaConsolidateDupsIgnoreILDiff:
                    return mFixedFastaOptions.ConsolidateDupsIgnoreILDiff;
                case SwitchOptions.FixedFastaTruncateLongProteinNames:
                    return mFixedFastaOptions.TruncateLongProteinNames;
                case SwitchOptions.FixedFastaSplitOutMultipleRefsForKnownAccession:
                    return mFixedFastaOptions.SplitOutMultipleRefsForKnownAccession;
                case SwitchOptions.FixedFastaWrapLongResidueLines:
                    return mFixedFastaOptions.WrapLongResidueLines;
                case SwitchOptions.FixedFastaRemoveInvalidResidues:
                    return mFixedFastaOptions.RemoveInvalidResidues;
                case SwitchOptions.SaveBasicProteinHashInfoFile:
                    return mSaveBasicProteinHashInfoFile;
                case SwitchOptions.AllowDashInResidues:
                    return mAllowDashInResidues;
                case SwitchOptions.AllowAllSymbolsInProteinNames:
                    return mAllowAllSymbolsInProteinNames;
            }

            return false;
        }

        [Obsolete("Use GetErrorWarningCounts instead", true)]
        public int get_ErrorWarningCounts(
            eMsgTypeConstants messageType,
            ErrorWarningCountTypes CountType)
        {
            return GetErrorWarningCounts(messageType, CountType);
        }

        public int GetErrorWarningCounts(
            eMsgTypeConstants messageType,
            ErrorWarningCountTypes CountType)
        {
            var tmpValue = 0;
            switch (CountType)
            {
                case ErrorWarningCountTypes.Total:
                    switch (messageType)
                    {
                        case eMsgTypeConstants.ErrorMsg:
                            tmpValue = mFileErrorCount + ComputeTotalUnspecifiedCount(mFileErrorStats);
                            break;
                        case eMsgTypeConstants.WarningMsg:
                            tmpValue = mFileWarningCount + ComputeTotalUnspecifiedCount(mFileWarningStats);
                            break;
                        case eMsgTypeConstants.StatusMsg:
                            tmpValue = 0;
                            break;
                    }

                    break;

                case ErrorWarningCountTypes.Unspecified:
                    switch (messageType)
                    {
                        case eMsgTypeConstants.ErrorMsg:
                            tmpValue = ComputeTotalUnspecifiedCount(mFileErrorStats);
                            break;
                        case eMsgTypeConstants.WarningMsg:
                            tmpValue = ComputeTotalSpecifiedCount(mFileWarningStats);
                            break;
                        case eMsgTypeConstants.StatusMsg:
                            tmpValue = 0;
                            break;
                    }

                    break;

                case ErrorWarningCountTypes.Specified:
                    switch (messageType)
                    {
                        case eMsgTypeConstants.ErrorMsg:
                            tmpValue = mFileErrorCount;
                            break;
                        case eMsgTypeConstants.WarningMsg:
                            tmpValue = mFileWarningCount;
                            break;
                        case eMsgTypeConstants.StatusMsg:
                            tmpValue = 0;
                            break;
                    }

                    break;
            }

            return tmpValue;
        }

        [Obsolete("Use GetFixedFASTAFileStats instead", true)]
        public int get_FixedFASTAFileStats(FixedFASTAFileValues valueType)
        {
            return GetFixedFASTAFileStats(valueType);
        }

        public int GetFixedFASTAFileStats(FixedFASTAFileValues valueType)
        {
            var tmpValue = 0;
            switch (valueType)
            {
                case FixedFASTAFileValues.DuplicateProteinNamesSkippedCount:
                    tmpValue = mFixedFastaStats.DuplicateNameProteinsSkipped;
                    break;
                case FixedFASTAFileValues.ProteinNamesInvalidCharsReplaced:
                    tmpValue = mFixedFastaStats.ProteinNamesInvalidCharsReplaced;
                    break;
                case FixedFASTAFileValues.ProteinNamesMultipleRefsRemoved:
                    tmpValue = mFixedFastaStats.ProteinNamesMultipleRefsRemoved;
                    break;
                case FixedFASTAFileValues.TruncatedProteinNameCount:
                    tmpValue = mFixedFastaStats.TruncatedProteinNameCount;
                    break;
                case FixedFASTAFileValues.UpdatedResidueLines:
                    tmpValue = mFixedFastaStats.UpdatedResidueLines;
                    break;
                case FixedFASTAFileValues.DuplicateProteinNamesRenamedCount:
                    tmpValue = mFixedFastaStats.DuplicateNameProteinsRenamed;
                    break;
                case FixedFASTAFileValues.DuplicateProteinSeqsSkippedCount:
                    tmpValue = mFixedFastaStats.DuplicateSequenceProteinsSkipped;
                    break;
            }

            return tmpValue;
        }

        public int ProteinCount => mProteinCount;

        public int LineCount => mLineCount;

        public eValidateFastaFileErrorCodes LocalErrorCode => mLocalErrorCode;

        public long ResidueCount => mResidueCount;

        public string FastaFilePath => mFastaFilePath;

        [Obsolete("Use GetErrorMessageTextByIndex", true)]
        public string get_ErrorMessageTextByIndex(int index, string valueSeparator)
        {
            return GetErrorMessageTextByIndex(index, valueSeparator);
        }

        public string GetErrorMessageTextByIndex(int index, string valueSeparator)
        {
            return GetFileErrorTextByIndex(index, valueSeparator);
        }

        [Obsolete("Use GetWarningMessageTextByIndex", true)]
        public string get_WarningMessageTextByIndex(int index, string valueSeparator)
        {
            return GetWarningMessageTextByIndex(index, valueSeparator);
        }

        public string GetWarningMessageTextByIndex(int index, string valueSeparator)
        {
            return GetFileWarningTextByIndex(index, valueSeparator);
        }

        [Obsolete("Use GetErrorsByIndex", true)]
        public udtMsgInfoType get_ErrorsByIndex(int errorIndex)
        {
            return GetErrorsByIndex(errorIndex);
        }

        public udtMsgInfoType GetErrorsByIndex(int errorIndex)
        {
            return GetFileErrorByIndex(errorIndex);
        }

        [Obsolete("Use GetWarningsByIndex", true)]
        public udtMsgInfoType get_WarningsByIndex(int warningIndex)
        {
            return GetWarningsByIndex(warningIndex);
        }

        public udtMsgInfoType GetWarningsByIndex(int warningIndex)
        {
            return GetFileWarningByIndex(warningIndex);
        }

        /// <summary>
        /// Existing protein hash file to load into memory instead of computing new hash values while reading the fasta file
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public string ExistingProteinHashFile { get; set; }

        public int MaximumFileErrorsToTrack
        {
            get => mMaximumFileErrorsToTrack;
            set
            {
                if (value < 1)
                    value = 1;

                mMaximumFileErrorsToTrack = value;
            }
        }

        public int MaximumProteinNameLength
        {
            get => mMaximumProteinNameLength;
            set
            {
                if (value < 8)
                {
                    // Do not allow maximum lengths less than 8; use the default
                    value = DEFAULT_MAXIMUM_PROTEIN_NAME_LENGTH;
                }

                mMaximumProteinNameLength = value;
            }
        }

        public int MinimumProteinNameLength
        {
            get => mMinimumProteinNameLength;
            set
            {
                if (value < 1)
                    value = DEFAULT_MINIMUM_PROTEIN_NAME_LENGTH;

                mMinimumProteinNameLength = value;
            }
        }

        public int MaximumResiduesPerLine
        {
            get => mMaximumResiduesPerLine;
            set
            {
                if (value == 0)
                {
                    value = DEFAULT_MAXIMUM_RESIDUES_PER_LINE;
                }
                else if (value < 40)
                {
                    value = 40;
                }

                mMaximumResiduesPerLine = value;
            }
        }

        public char ProteinLineStartChar
        {
            get => mProteinLineStartChar;
            set => mProteinLineStartChar = value;
        }

        public string StatsFilePath
        {
            get
            {
                if (mStatsFilePath == null)
                {
                    return string.Empty;
                }
                else
                {
                    return mStatsFilePath;
                }
            }
        }

        public string ProteinNameInvalidCharsToRemove
        {
            get => CharArrayToString(mFixedFastaOptions.ProteinNameInvalidCharsToRemove);
            set
            {
                if (value == null)
                {
                    value = string.Empty;
                }

                // Check for and remove any spaces from Value, since
                // a space does not make sense for an invalid protein name character
                value = value.Replace(" ", string.Empty);
                if (value.Length > 0)
                {
                    mFixedFastaOptions.ProteinNameInvalidCharsToRemove = value.ToCharArray();
                }
                else
                {
                    mFixedFastaOptions.ProteinNameInvalidCharsToRemove = new char[] { };      // Default to an empty character array if Value is empty
                }
            }
        }

        public string ProteinNameFirstRefSepChars
        {
            get => CharArrayToString(mProteinNameFirstRefSepChars);
            set
            {
                if (value == null)
                {
                    value = string.Empty;
                }

                // Check for and remove any spaces from Value, since
                // a space does not make sense for a separation character
                value = value.Replace(" ", string.Empty);
                if (value.Length > 0)
                {
                    mProteinNameFirstRefSepChars = value.ToCharArray();
                }
                else
                {
                    mProteinNameFirstRefSepChars = DEFAULT_PROTEIN_NAME_FIRST_REF_SEP_CHARS.ToCharArray();     // Use the default if Value is empty
                }
            }
        }

        public string ProteinNameSubsequentRefSepChars
        {
            get => CharArrayToString(mProteinNameSubsequentRefSepChars);
            set
            {
                if (value == null)
                {
                    value = string.Empty;
                }

                // Check for and remove any spaces from Value, since
                // a space does not make sense for a separation character
                value = value.Replace(" ", string.Empty);
                if (value.Length > 0)
                {
                    mProteinNameSubsequentRefSepChars = value.ToCharArray();
                }
                else
                {
                    mProteinNameSubsequentRefSepChars = DEFAULT_PROTEIN_NAME_SUBSEQUENT_REF_SEP_CHARS.ToCharArray();     // Use the default if Value is empty
                }
            }
        }

        public string LongProteinNameSplitChars
        {
            get => CharArrayToString(mFixedFastaOptions.LongProteinNameSplitChars);
            set
            {
                if (value != null)
                {
                    // Check for and remove any spaces from Value, since
                    // a space does not make sense for a protein name split char
                    value = value.Replace(" ", string.Empty);
                    if (value.Length > 0)
                    {
                        mFixedFastaOptions.LongProteinNameSplitChars = value.ToCharArray();
                    }
                }
            }
        }

        public List<udtMsgInfoType> FileWarningList => GetFileWarnings();

        public List<udtMsgInfoType> FileErrorList => GetFileErrors();

        #endregion
        public event ProgressCompletedEventHandler ProgressCompleted;

        public delegate void ProgressCompletedEventHandler();

        public event WroteLineEndNormalizedFASTAEventHandler WroteLineEndNormalizedFASTA;

        public delegate void WroteLineEndNormalizedFASTAEventHandler(string newFilePath);

        private void OnProgressComplete()
        {
            ProgressCompleted?.Invoke();
            OperationComplete();
        }

        private void OnWroteLineEndNormalizedFASTA(string newFilePath)
        {
            WroteLineEndNormalizedFASTA?.Invoke(newFilePath);
        }

        /// <summary>
        /// Examine the given fasta file to look for problems.
        /// Optionally create a new, fixed fasta file
        /// Optionally also consolidate proteins with duplicate sequences
        /// </summary>
        /// <param name="fastaFilePathToCheck"></param>
        /// <param name="preloadedProteinNamesToKeep">
        /// Preloaded list of protein names to include in the fixed fasta file
        /// Keys are protein names, values are the number of entries written to the fixed fasta file for the given protein name
        /// </param>
        /// <returns>True if the file was successfully analyzed (even if errors were found)</returns>
        /// <remarks>Assumes fastaFilePathToCheck exists</remarks>
        private bool AnalyzeFastaFile(string fastaFilePathToCheck, clsNestedStringIntList preloadedProteinNamesToKeep)
        {
            StreamWriter fixedFastaWriter = null;
            StreamWriter sequenceHashWriter = null;

            string fastaFilePathOut = "UndefinedFilePath.xyz";

            bool success;
            bool exceptionCaught = false;

            bool consolidateDuplicateProteinSeqsInFasta = false;
            bool keepDuplicateNamedProteinsUnlessMatchingSequence = false;
            bool consolidateDupsIgnoreILDiff = false;

            // This array tracks protein hash details
            var proteinSequenceHashCount = 0;
            clsProteinHashInfo[] proteinSeqHashInfo;

            udtRuleDefinitionExtendedType[] headerLineRuleDetails;
            udtRuleDefinitionExtendedType[] proteinNameRuleDetails;
            udtRuleDefinitionExtendedType[] proteinDescriptionRuleDetails;
            udtRuleDefinitionExtendedType[] proteinSequenceRuleDetails;

            try
            {
                // Reset the data structures and variables
                ResetStructures();
                proteinSeqHashInfo = new clsProteinHashInfo[1];

                headerLineRuleDetails = new udtRuleDefinitionExtendedType[2];
                proteinNameRuleDetails = new udtRuleDefinitionExtendedType[2];
                proteinDescriptionRuleDetails = new udtRuleDefinitionExtendedType[2];
                proteinSequenceRuleDetails = new udtRuleDefinitionExtendedType[2];

                // This is a dictionary of dictionaries, with one dictionary for each letter or number that a SHA-1 hash could start with
                // This dictionary of dictionaries provides a quick lookup for existing protein hashes
                // This dictionary is not used if preloadedProteinNamesToKeep contains data
                const int SPANNER_CHAR_LENGTH = 1;
                var proteinSequenceHashes = new clsNestedStringDictionary<int>(false, SPANNER_CHAR_LENGTH);
                bool usingPreloadedProteinNames = false;

                if (preloadedProteinNamesToKeep != null && preloadedProteinNamesToKeep.Count > 0)
                {
                    // Auto enable/disable some options
                    mSaveBasicProteinHashInfoFile = false;
                    mCheckForDuplicateProteinSequences = false;

                    // Auto-enable creating a fixed fasta file
                    mGenerateFixedFastaFile = true;
                    mFixedFastaOptions.ConsolidateProteinsWithDuplicateSeqs = false;
                    mFixedFastaOptions.RenameProteinsWithDuplicateNames = false;

                    // Note: do not change .ConsolidateDupsIgnoreILDiff
                    // If .ConsolidateDupsIgnoreILDiff was enabled when the hash file was made with /B
                    // it should also be enabled when using /HashFile

                    usingPreloadedProteinNames = true;
                }

                if (mNormalizeFileLineEndCharacters)
                {
                    mFastaFilePath = NormalizeFileLineEndings(
                        fastaFilePathToCheck,
                        "CRLF_" + Path.GetFileName(fastaFilePathToCheck),
                        eLineEndingCharacters.CRLF);

                    if ((mFastaFilePath ?? "") != (fastaFilePathToCheck ?? ""))
                    {
                        fastaFilePathToCheck = string.Copy(mFastaFilePath);
                        OnWroteLineEndNormalizedFASTA(fastaFilePathToCheck);
                    }
                }
                else
                {
                    mFastaFilePath = string.Copy(fastaFilePathToCheck);
                }

                OnProgressUpdate("Parsing " + Path.GetFileName(mFastaFilePath), 0);

                bool proteinHeaderFound = false;
                bool processingResidueBlock = false;
                bool blankLineProcessed = false;

                string proteinName = string.Empty;
                var sbCurrentResidues = new StringBuilder();

                // Initialize the RegEx objects

                var reProteinNameTruncation = new udtProteinNameTruncationRegex();

                // Note that each of these RegEx tests contain two groups with captured text:

                // The following will extract IPI:IPI00048500.11 from IPI:IPI00048500.11|ref|23848934
                reProteinNameTruncation.reMatchIPI =
                    new Regex(@"^(IPI:IPI[\w.]{2,})\|(.+)",
                    RegexOptions.Singleline | RegexOptions.Compiled);

                // The following will extract gi|169602219 from gi|169602219|ref|XP_001794531.1|
                reProteinNameTruncation.reMatchGI =
                    new Regex(@"^(gi\|\d+)\|(.+)",
                    RegexOptions.Singleline | RegexOptions.Compiled);

                // The following will extract jgi|Batde5|906240 from jgi|Batde5|90624|GP3.061830
                reProteinNameTruncation.reMatchJGI =
                    new Regex(@"^(jgi\|[^|]+\|[^|]+)\|(.+)",
                    RegexOptions.Singleline | RegexOptions.Compiled);

                // The following will extract bob|234384 from  bob|234384|ref|483293
                // or bob|845832 from  bob|845832;ref|384923
                reProteinNameTruncation.reMatchGeneric =
                    new Regex(@"^(\w{2,}[" +
                        CharArrayToString(mProteinNameFirstRefSepChars) + @"][\w\d._]{2,})[" +
                        CharArrayToString(mProteinNameSubsequentRefSepChars) + "](.+)",
                        RegexOptions.Singleline | RegexOptions.Compiled);
                // The following matches jgi|Batde5|23435 ; it requires that there be a number after the second bar
                reProteinNameTruncation.reMatchJGIBaseAndID =
                new Regex(@"^jgi\|[^|]+\|\d+",
                RegexOptions.Singleline | RegexOptions.Compiled);

                // Note that this RegEx contains a group with captured text:
                reProteinNameTruncation.reMatchDoubleBarOrColonAndBar =
                    new Regex("[" +
                        CharArrayToString(mProteinNameFirstRefSepChars) + "][^" +
                        CharArrayToString(mProteinNameSubsequentRefSepChars) + "]*([" +
                        CharArrayToString(mProteinNameSubsequentRefSepChars) + "])",
                        RegexOptions.Singleline | RegexOptions.Compiled);

                // Non-letter characters in residues
                string allowedResidueChars = "A-Z";
                if (mAllowAsteriskInResidues)
                    allowedResidueChars += "*";
                if (mAllowDashInResidues)
                    allowedResidueChars += "-";

                var reNonLetterResidues =
                    new Regex("[^" + allowedResidueChars + "]",
                    RegexOptions.Singleline | RegexOptions.Compiled);

                // Make sure mFixedFastaOptions.LongProteinNameSplitChars contains at least one character
                if (mFixedFastaOptions.LongProteinNameSplitChars == null || mFixedFastaOptions.LongProteinNameSplitChars.Length == 0)
                {
                    mFixedFastaOptions.LongProteinNameSplitChars = new char[] { DEFAULT_LONG_PROTEIN_NAME_SPLIT_CHAR };
                }

                // Initialize the rule details UDTs, which contain a RegEx object for each rule
                InitializeRuleDetails(ref mHeaderLineRules, ref headerLineRuleDetails);
                InitializeRuleDetails(ref mProteinNameRules, ref proteinNameRuleDetails);
                InitializeRuleDetails(ref mProteinDescriptionRules, ref proteinDescriptionRuleDetails);
                InitializeRuleDetails(ref mProteinSequenceRules, ref proteinSequenceRuleDetails);

                // Open the file and read, at most, the first 100,000 characters to see if it contains CrLf or just Lf
                int terminatorSize = DetermineLineTerminatorSize(fastaFilePathToCheck);

                // Pre-scan a portion of the fasta file to determine the appropriate value for mProteinNameSpannerCharLength
                AutoDetermineFastaProteinNameSpannerCharLength(mFastaFilePath, terminatorSize);

                // Open the input file
                var startTime = DateTime.UtcNow;

                using (var fastaReader = new StreamReader(new FileStream(fastaFilePathToCheck, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)))
                {
                    // Optionally, open the output fasta file
                    if (mGenerateFixedFastaFile)
                    {
                        try
                        {
                            fastaFilePathOut =
                                Path.Combine(Path.GetDirectoryName(fastaFilePathToCheck),
                                Path.GetFileNameWithoutExtension(fastaFilePathToCheck) + "_new.fasta");
                            fixedFastaWriter = new StreamWriter(new FileStream(fastaFilePathOut, FileMode.Create, FileAccess.Write, FileShare.ReadWrite));
                        }
                        catch (Exception ex)
                        {
                            // Error opening output file
                            RecordFastaFileError(0, 0, string.Empty, (int)eMessageCodeConstants.UnspecifiedError,
                                "Error creating output file " + fastaFilePathOut + ": " + ex.Message, string.Empty);
                            OnErrorEvent("Error creating output file (Create _new.fasta)", ex);
                            return false;
                        }
                    }

                    // Optionally, open the Sequence Hash file
                    if (mSaveBasicProteinHashInfoFile)
                    {
                        string basicProteinHashInfoFilePath = "<undefined>";

                        try
                        {
                            basicProteinHashInfoFilePath =
                                Path.Combine(Path.GetDirectoryName(fastaFilePathToCheck),
                                Path.GetFileNameWithoutExtension(fastaFilePathToCheck) + PROTEIN_HASHES_FILENAME_SUFFIX);
                            sequenceHashWriter = new StreamWriter(new FileStream(basicProteinHashInfoFilePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite));

                            var headerNames = new List<string>()
                            {
                                "Protein_ID",
                                PROTEIN_NAME_COLUMN,
                                SEQUENCE_LENGTH_COLUMN,
                                SEQUENCE_HASH_COLUMN
                            };

                            sequenceHashWriter.WriteLine(FlattenList(headerNames));
                        }
                        catch (Exception ex)
                        {
                            // Error opening output file
                            RecordFastaFileError(0, 0, string.Empty, (int)eMessageCodeConstants.UnspecifiedError,
                                "Error creating output file " + basicProteinHashInfoFilePath + ": " + ex.Message, string.Empty);
                            OnErrorEvent("Error creating output file (Create " + PROTEIN_HASHES_FILENAME_SUFFIX + ")", ex);
                        }
                    }

                    if (mGenerateFixedFastaFile && (mFixedFastaOptions.RenameProteinsWithDuplicateNames || mFixedFastaOptions.KeepDuplicateNamedProteinsUnlessMatchingSequence))
                    {
                        // Make sure mCheckForDuplicateProteinNames is enabled
                        mCheckForDuplicateProteinNames = true;
                    }

                    // Initialize proteinNames
                    var proteinNames = new SortedSet<string>(StringComparer.CurrentCultureIgnoreCase);

                    // Optionally, initialize the protein sequence hash objects
                    if (mSaveProteinSequenceHashInfoFiles)
                    {
                        mCheckForDuplicateProteinSequences = true;
                    }

                    if (mGenerateFixedFastaFile && mFixedFastaOptions.ConsolidateProteinsWithDuplicateSeqs)
                    {
                        mCheckForDuplicateProteinSequences = true;
                        mSaveProteinSequenceHashInfoFiles = !usingPreloadedProteinNames;
                        consolidateDuplicateProteinSeqsInFasta = true;
                        keepDuplicateNamedProteinsUnlessMatchingSequence = mFixedFastaOptions.KeepDuplicateNamedProteinsUnlessMatchingSequence;
                        consolidateDupsIgnoreILDiff = mFixedFastaOptions.ConsolidateDupsIgnoreILDiff;
                    }
                    else if (mGenerateFixedFastaFile && mFixedFastaOptions.KeepDuplicateNamedProteinsUnlessMatchingSequence)
                    {
                        mCheckForDuplicateProteinSequences = true;
                        mSaveProteinSequenceHashInfoFiles = !usingPreloadedProteinNames;
                        consolidateDuplicateProteinSeqsInFasta = false;
                        keepDuplicateNamedProteinsUnlessMatchingSequence = mFixedFastaOptions.KeepDuplicateNamedProteinsUnlessMatchingSequence;
                        consolidateDupsIgnoreILDiff = mFixedFastaOptions.ConsolidateDupsIgnoreILDiff;
                    }

                    if (mCheckForDuplicateProteinSequences)
                    {
                        proteinSequenceHashes.Clear();
                        proteinSequenceHashCount = 0;
                        proteinSeqHashInfo = new clsProteinHashInfo[100];
                    }

                    // Parse each line in the file
                    long bytesRead = 0;
                    var lastMemoryUsageReport = DateTime.UtcNow;

                    // Note: This value is updated only if the line length is < mMaximumResiduesPerLine
                    int currentValidResidueLineLengthMax = 0;
                    bool processingDuplicateOrInvalidProtein = false;

                    var lastProgressReport = DateTime.UtcNow;

                    while (!fastaReader.EndOfStream)
                    {
                        string lineIn;
                        try
                        {
                            lineIn = fastaReader.ReadLine();
                        }
                        catch (OutOfMemoryException ex)
                        {
                            OnErrorEvent(string.Format(
                                "Error in AnalyzeFastaFile reading line {0}; " +
                                "it is most likely millions of characters long, " +
                                "indicating a corrupt fasta file", mLineCount + 1), ex);
                            exceptionCaught = true;
                            break;
                        }
                        catch (Exception ex)
                        {
                            OnErrorEvent(string.Format("Error in AnalyzeFastaFile reading line {0}", mLineCount + 1), ex);
                            exceptionCaught = true;
                            break;
                        }

                        bytesRead += lineIn.Length + terminatorSize;

                        if (mLineCount % 250 == 0)
                        {
                            if (AbortProcessing)
                                break;

                            if (DateTime.UtcNow.Subtract(lastProgressReport).TotalSeconds >= 0.5)
                            {
                                lastProgressReport = DateTime.UtcNow;
                                float percentComplete = (float)(bytesRead / (double)fastaReader.BaseStream.Length * 100.0);
                                if (consolidateDuplicateProteinSeqsInFasta || keepDuplicateNamedProteinsUnlessMatchingSequence)
                                {
                                    // Bump the % complete down so that 100% complete in this routine will equate to 75% complete
                                    // The remaining 25% will occur in ConsolidateDuplicateProteinSeqsInFasta
                                    percentComplete = percentComplete * 3 / 4;
                                }

                                UpdateProgress("Validating FASTA File (" + Math.Round(percentComplete, 0) + "% Done)", percentComplete);

                                if (DateTime.UtcNow.Subtract(lastMemoryUsageReport).TotalMinutes >= 1)
                                {
                                    lastMemoryUsageReport = DateTime.UtcNow;
                                    ReportMemoryUsage(preloadedProteinNamesToKeep, proteinSequenceHashes, proteinNames, proteinSeqHashInfo);
                                }
                            }
                        }

                        mLineCount += 1;

                        if (lineIn.Length > 10000000)
                        {
                            RecordFastaFileError(mLineCount, 0, proteinName, (int)eMessageCodeConstants.ResiduesLineTooLong, "Line is over 10 million residues long; skipping", string.Empty);
                            continue;
                        }
                        else if (lineIn.Length > 1000000)
                        {
                            RecordFastaFileWarning(mLineCount, 0, proteinName, (int)eMessageCodeConstants.ResiduesLineTooLong, "Line is over 1 million residues long; this is very suspicious", string.Empty);
                        }
                        else if (lineIn.Length > 100000)
                        {
                            RecordFastaFileWarning(mLineCount, 0, proteinName, (int)eMessageCodeConstants.ResiduesLineTooLong, "Line is over 1 million residues long; this could indicate a problem", string.Empty);
                        }

                        if (lineIn == null)
                            continue;

                        if (lineIn.Trim().Length == 0)
                        {
                            // We typically only want blank lines at the end of the fasta file or between two protein entries
                            blankLineProcessed = true;
                            continue;
                        }

                        if (lineIn[0] == ' ')
                        {
                            if (mWarnLineStartsWithSpace)
                            {
                                RecordFastaFileError(mLineCount, 0, string.Empty,
                                                     (int)eMessageCodeConstants.LineStartsWithSpace, string.Empty, ExtractContext(lineIn, 0));
                            }
                        }

                        // Note: Only trim the start of the line; do not trim the end of the line since Sequest incorrectly notates the peptide terminal state if a residue has a space after it
                        lineIn = lineIn.TrimStart();

                        if (lineIn[0] == mProteinLineStartChar)
                        {
                            // Protein entry

                            if (sbCurrentResidues.Length > 0)
                            {
                                ProcessResiduesForPreviousProtein(
                                    proteinName, sbCurrentResidues,
                                    proteinSequenceHashes,
                                    ref proteinSequenceHashCount, ref proteinSeqHashInfo,
                                    consolidateDupsIgnoreILDiff,
                                    fixedFastaWriter,
                                    currentValidResidueLineLengthMax,
                                    sequenceHashWriter);

                                currentValidResidueLineLengthMax = 0;
                            }

                            // Now process this protein entry
                            mProteinCount += 1;
                            proteinHeaderFound = true;
                            processingResidueBlock = false;
                            processingDuplicateOrInvalidProtein = false;

                            proteinName = string.Empty;

                            AnalyzeFastaProcessProteinHeader(
                                fixedFastaWriter,
                                lineIn,
                                out proteinName,
                                out processingDuplicateOrInvalidProtein,
                                preloadedProteinNamesToKeep,
                                proteinNames,
                                headerLineRuleDetails,
                                proteinNameRuleDetails,
                                proteinDescriptionRuleDetails,
                                reProteinNameTruncation);

                            if (blankLineProcessed)
                            {
                                // The previous line was blank; raise a warning
                                if (mWarnBlankLinesBetweenProteins)
                                {
                                    RecordFastaFileWarning(mLineCount, 0, proteinName, (int)eMessageCodeConstants.BlankLineBeforeProteinName);
                                }
                            }
                        }
                        else
                        {
                            // Protein residues

                            if (!processingResidueBlock)
                            {
                                if (proteinHeaderFound)
                                {
                                    proteinHeaderFound = false;

                                    if (blankLineProcessed)
                                    {
                                        RecordFastaFileError(mLineCount, 0, proteinName, (int)eMessageCodeConstants.BlankLineBetweenProteinNameAndResidues);
                                    }
                                }
                                else
                                {
                                    RecordFastaFileError(mLineCount, 0, string.Empty, (int)eMessageCodeConstants.ResiduesFoundWithoutProteinHeader);
                                }

                                processingResidueBlock = true;
                            }
                            else if (blankLineProcessed)
                            {
                                RecordFastaFileError(mLineCount, 0, proteinName, (int)eMessageCodeConstants.BlankLineInMiddleOfResidues);
                            }

                            int newResidueCount = lineIn.Length;
                            mResidueCount += newResidueCount;

                            // Check the line length; raise a warning if longer than suggested
                            if (newResidueCount > mMaximumResiduesPerLine)
                            {
                                RecordFastaFileWarning(mLineCount, 0, proteinName, (int)eMessageCodeConstants.ResiduesLineTooLong, newResidueCount.ToString(), string.Empty);
                            }

                            // Test the protein sequence rules
                            EvaluateRules(proteinSequenceRuleDetails, proteinName, lineIn, 0, lineIn, 5);

                            if (mGenerateFixedFastaFile || mCheckForDuplicateProteinSequences || mSaveBasicProteinHashInfoFile)
                            {
                                string residuesClean;

                                if (mFixedFastaOptions.RemoveInvalidResidues)
                                {
                                    // Auto-fix residues to remove any non-letter characters (spaces, asterisks, etc.)
                                    residuesClean = reNonLetterResidues.Replace(lineIn, string.Empty);
                                }
                                else
                                {
                                    // Do not remove non-letter characters, but do remove leading or trailing whitespace
                                    residuesClean = string.Copy(lineIn.Trim());
                                }

                                if (fixedFastaWriter != null && !processingDuplicateOrInvalidProtein)
                                {
                                    if ((residuesClean ?? "") != (lineIn ?? ""))
                                    {
                                        mFixedFastaStats.UpdatedResidueLines += 1;
                                    }

                                    if (!mFixedFastaOptions.WrapLongResidueLines)
                                    {
                                        // Only write out this line if not auto-wrapping long residue lines
                                        // If we are auto-wrapping, then the residues will be written out by the call to ProcessResiduesForPreviousProtein
                                        fixedFastaWriter.WriteLine(residuesClean);
                                    }
                                }

                                if (mCheckForDuplicateProteinSequences || mFixedFastaOptions.WrapLongResidueLines)
                                {
                                    // Only add the residues if this is not a duplicate/invalid protein
                                    if (!processingDuplicateOrInvalidProtein)
                                    {
                                        sbCurrentResidues.Append(residuesClean);
                                        if (residuesClean.Length > currentValidResidueLineLengthMax)
                                        {
                                            currentValidResidueLineLengthMax = residuesClean.Length;
                                        }
                                    }
                                }
                            }

                            // Reset the blank line tracking variable
                            blankLineProcessed = false;
                        }
                    }

                    if (sbCurrentResidues.Length > 0)
                    {
                        ProcessResiduesForPreviousProtein(
                            proteinName, sbCurrentResidues,
                            proteinSequenceHashes,
                            ref proteinSequenceHashCount, ref proteinSeqHashInfo,
                            consolidateDupsIgnoreILDiff,
                            fixedFastaWriter, currentValidResidueLineLengthMax,
                            sequenceHashWriter);
                    }

                    if (mCheckForDuplicateProteinSequences)
                    {
                        // Step through proteinSeqHashInfo and look for duplicate sequences
                        for (int index = 0; index <= proteinSequenceHashCount - 1; index++)
                        {
                            if (proteinSeqHashInfo[index].AdditionalProteins.Count() > 0)
                            {
                                var proteinHashInfo = proteinSeqHashInfo[index];
                                RecordFastaFileWarning(mLineCount, 0, proteinHashInfo.ProteinNameFirst, (int)eMessageCodeConstants.DuplicateProteinSequence,
                                    proteinHashInfo.ProteinNameFirst + ", " + FlattenArray(proteinHashInfo.AdditionalProteins, ','), proteinHashInfo.SequenceStart);
                            }
                        }
                    }

                    float memoryUsageMB = clsMemoryUsageLogger.GetProcessMemoryUsageMB();
                    if (memoryUsageMB > mProcessMemoryUsageMBAtStart * 4 ||
                        memoryUsageMB - mProcessMemoryUsageMBAtStart > 50)
                    {
                        ReportMemoryUsage(preloadedProteinNamesToKeep, proteinSequenceHashes, proteinNames, proteinSeqHashInfo);
                    }
                }

                double totalSeconds = DateTime.UtcNow.Subtract(startTime).TotalSeconds;
                if (totalSeconds > 5)
                {
                    int linesPerSecond = (int)Math.Round(mLineCount / totalSeconds);
                    Console.WriteLine();
                    ShowMessage(string.Format(
                        "Processing complete after {0:N0} seconds; read {1:N0} lines/second",
                        totalSeconds, linesPerSecond));
                }

                // Close the output files
                if (fixedFastaWriter != null)
                {
                    fixedFastaWriter.Close();
                }

                if (sequenceHashWriter != null)
                {
                    sequenceHashWriter.Close();
                }

                if (mProteinCount == 0)
                {
                    RecordFastaFileError(mLineCount, 0, string.Empty, (int)eMessageCodeConstants.ProteinEntriesNotFound);
                }
                else if (proteinHeaderFound)
                {
                    RecordFastaFileError(mLineCount, 0, proteinName, (int)eMessageCodeConstants.FinalProteinEntryMissingResidues);
                }
                else if (!blankLineProcessed)
                {
                    // File does not end in multiple blank lines; need to re-open it using a binary reader and check the last two characters to make sure they're valid
                    System.Threading.Thread.Sleep(100);

                    if (!VerifyLinefeedAtEOF(fastaFilePathToCheck, mAddMissingLinefeedAtEOF))
                    {
                        RecordFastaFileError(mLineCount, 0, string.Empty, (int)eMessageCodeConstants.FileDoesNotEndWithLinefeed);
                    }
                }

                if (usingPreloadedProteinNames)
                {
                    // Report stats on the number of proteins read, the number written, and any that had duplicate protein names in the original fasta file
                    int nameCountNotFound = 0;
                    int duplicateProteinNameCount = 0;
                    int proteinCountWritten = 0;
                    int preloadedProteinNameCount = 0;

                    foreach (var spanningKey in preloadedProteinNamesToKeep.GetSpanningKeys())
                    {
                        var proteinsForKey = preloadedProteinNamesToKeep.GetListForSpanningKey(spanningKey);
                        preloadedProteinNameCount += proteinsForKey.Count;

                        foreach (var proteinEntry in proteinsForKey)
                        {
                            if (proteinEntry.Value == 0)
                            {
                                nameCountNotFound += 1;
                            }
                            else
                            {
                                proteinCountWritten += 1;
                                if (proteinEntry.Value > 1)
                                {
                                    duplicateProteinNameCount += 1;
                                }
                            }
                        }
                    }

                    Console.WriteLine();
                    if (proteinCountWritten == preloadedProteinNameCount)
                    {
                        ShowMessage("Fixed Fasta has all " + proteinCountWritten.ToString("#,##0") + " proteins determined from the pre-existing protein hash file");
                    }
                    else
                    {
                        ShowMessage("Fixed Fasta has " + proteinCountWritten.ToString("#,##0") + " of the " + preloadedProteinNameCount.ToString("#,##0") + " proteins determined from the pre-existing protein hash file");
                    }

                    if (nameCountNotFound > 0)
                    {
                        ShowMessage("WARNING: " + nameCountNotFound.ToString("#,##0") + " protein names were in the protein name list to keep, but were not found in the fasta file");
                    }

                    if (duplicateProteinNameCount > 0)
                    {
                        ShowMessage("WARNING: " + duplicateProteinNameCount.ToString("#,##0") + " protein names were present multiple times in the fasta file; duplicate entries were skipped");
                    }

                    success = !exceptionCaught;
                }
                else if (mSaveProteinSequenceHashInfoFiles)
                {
                    float percentComplete = 98.0F;
                    if (consolidateDuplicateProteinSeqsInFasta || keepDuplicateNamedProteinsUnlessMatchingSequence)
                    {
                        percentComplete = percentComplete * 3 / 4;
                    }

                    UpdateProgress("Validating FASTA File (" + Math.Round(percentComplete, 0) + "% Done)", percentComplete);

                    bool hashInfoSuccess = AnalyzeFastaSaveHashInfo(
                        fastaFilePathToCheck,
                        proteinSequenceHashCount,
                        proteinSeqHashInfo,
                        consolidateDuplicateProteinSeqsInFasta,
                        consolidateDupsIgnoreILDiff,
                        keepDuplicateNamedProteinsUnlessMatchingSequence,
                        fastaFilePathOut);

                    success = hashInfoSuccess && !exceptionCaught;
                }
                else
                {
                    success = !exceptionCaught;
                }

                if (AbortProcessing)
                {
                    UpdateProgress("Parsing aborted");
                }
                else
                {
                    UpdateProgress("Parsing complete", 100);
                }
            }
            catch (Exception ex)
            {
                OnErrorEvent(string.Format("Error in AnalyzeFastaFile reading line {0}", mLineCount), ex);
                success = false;
            }
            finally
            {
                // These close statements will typically be redundant,
                // However, if an exception occurs, then they will be needed to close the files

                if (fixedFastaWriter != null)
                {
                    fixedFastaWriter.Close();
                }

                if (sequenceHashWriter != null)
                {
                    sequenceHashWriter.Close();
                }
            }

            return success;
        }

        private void AnalyzeFastaProcessProteinHeader(
            TextWriter fixedFastaWriter,
            string lineIn,
            out string proteinName,
            out bool processingDuplicateOrInvalidProtein,
            clsNestedStringIntList preloadedProteinNamesToKeep,
            ISet<string> proteinNames,
            IList<udtRuleDefinitionExtendedType> headerLineRuleDetails,
            IList<udtRuleDefinitionExtendedType> proteinNameRuleDetails,
            IList<udtRuleDefinitionExtendedType> proteinDescriptionRuleDetails,
            udtProteinNameTruncationRegex reProteinNameTruncation)
        {
            int descriptionStartIndex;

            string proteinDescription = string.Empty;

            bool skipDuplicateProtein = false;

            proteinName = string.Empty;
            processingDuplicateOrInvalidProtein = true;

            try
            {
                SplitFastaProteinHeaderLine(lineIn, out proteinName, out proteinDescription, out descriptionStartIndex);

                if (proteinName.Length == 0)
                {
                    processingDuplicateOrInvalidProtein = true;
                }
                else
                {
                    processingDuplicateOrInvalidProtein = false;
                }

                // Test the header line rules
                EvaluateRules(headerLineRuleDetails, proteinName, lineIn, 0, lineIn, DEFAULT_CONTEXT_LENGTH);

                if (proteinDescription.Length > 0)
                {
                    // Test the protein description rules

                    EvaluateRules(
                        proteinDescriptionRuleDetails, proteinName, proteinDescription,
                        descriptionStartIndex, lineIn, DEFAULT_CONTEXT_LENGTH);
                }

                if (proteinName.Length > 0)
                {

                    // Check for protein names that are too long or too short
                    if (proteinName.Length < mMinimumProteinNameLength)
                    {
                        RecordFastaFileWarning(mLineCount, 1, proteinName,
                            (int)eMessageCodeConstants.ProteinNameIsTooShort, proteinName.Length.ToString(), string.Empty);
                    }
                    else if (proteinName.Length > mMaximumProteinNameLength)
                    {
                        RecordFastaFileError(mLineCount, 1, proteinName,
                            (int)eMessageCodeConstants.ProteinNameIsTooLong, proteinName.Length.ToString(), string.Empty);
                    }

                    // Test the protein name rules
                    EvaluateRules(proteinNameRuleDetails, proteinName, proteinName, 1, lineIn, DEFAULT_CONTEXT_LENGTH);

                    if (preloadedProteinNamesToKeep != null && preloadedProteinNamesToKeep.Count > 0)
                    {
                        // See if preloadedProteinNamesToKeep contains proteinName
                        int matchCount = preloadedProteinNamesToKeep.GetValueForItem(proteinName, -1);

                        if (matchCount >= 0)
                        {
                            // Name is known; increment the value for this protein

                            if (matchCount == 0)
                            {
                                skipDuplicateProtein = false;
                            }
                            else
                            {
                                // An entry with this protein name has already been written
                                // Do not include the duplicate
                                skipDuplicateProtein = true;
                            }

                            if (!preloadedProteinNamesToKeep.SetValueForItem(proteinName, matchCount + 1))
                            {
                                ShowMessage("WARNING: protein " + proteinName + " not found in preloadedProteinNamesToKeep");
                            }
                        }
                        else
                        {
                            // Unknown protein name; do not keep this protein
                            skipDuplicateProtein = true;
                            processingDuplicateOrInvalidProtein = true;
                        }

                        if (mGenerateFixedFastaFile)
                        {
                            // Make sure proteinDescription doesn't start with a | or space
                            if (proteinDescription.Length > 0)
                            {
                                proteinDescription = proteinDescription.TrimStart(new char[] { '|', ' ' });
                            }
                        }
                    }
                    else
                    {
                        if (mGenerateFixedFastaFile)
                        {
                            proteinName = AutoFixProteinNameAndDescription(ref proteinName, ref proteinDescription, reProteinNameTruncation);
                        }

                        // Optionally, check for duplicate protein names
                        if (mCheckForDuplicateProteinNames)
                        {
                            proteinName = ExamineProteinName(ref proteinName, proteinNames, out skipDuplicateProtein, ref processingDuplicateOrInvalidProtein);

                            if (skipDuplicateProtein)
                            {
                                processingDuplicateOrInvalidProtein = true;
                            }
                        }
                    }

                    if (fixedFastaWriter != null && !skipDuplicateProtein)
                    {
                        fixedFastaWriter.WriteLine(ConstructFastaHeaderLine(proteinName.Trim(), proteinDescription.Trim()));
                    }
                }
            }
            catch (Exception ex)
            {
                RecordFastaFileError(0, 0, string.Empty, (int)eMessageCodeConstants.UnspecifiedError,
                    "Error parsing protein header line '" + lineIn + "': " + ex.Message, string.Empty);
                OnErrorEvent("Error parsing protein header line", ex);
            }
        }

        private bool AnalyzeFastaSaveHashInfo(
            string fastaFilePathToCheck,
            int proteinSequenceHashCount,
            IList<clsProteinHashInfo> proteinSeqHashInfo,
            bool consolidateDuplicateProteinSeqsInFasta,
            bool consolidateDupsIgnoreILDiff,
            bool keepDuplicateNamedProteinsUnlessMatchingSequence,
            string fastaFilePathOut)
        {
            StreamWriter swUniqueProteinSeqsOut;
            StreamWriter swDuplicateProteinMapping = null;

            string uniqueProteinSeqsFileOut = string.Empty;
            string duplicateProteinMappingFileOut = string.Empty;

            int index;
            var duplicateIndex = 0;

            bool duplicateProteinSeqsFound;
            bool success;

            duplicateProteinSeqsFound = false;

            try
            {
                uniqueProteinSeqsFileOut =
                    Path.Combine(Path.GetDirectoryName(fastaFilePathToCheck),
                    Path.GetFileNameWithoutExtension(fastaFilePathToCheck) + "_UniqueProteinSeqs.txt");

                // Create swUniqueProteinSeqsOut
                swUniqueProteinSeqsOut = new StreamWriter(new FileStream(uniqueProteinSeqsFileOut, FileMode.Create, FileAccess.Write, FileShare.ReadWrite));
            }
            catch (Exception ex)
            {
                // Error opening output file
                RecordFastaFileError(0, 0, string.Empty, (int)eMessageCodeConstants.UnspecifiedError,
                    "Error creating output file " + uniqueProteinSeqsFileOut + ": " + ex.Message, string.Empty);
                OnErrorEvent("Error creating output file (SaveHashInfo to _UniqueProteinSeqs.txt)", ex);
                return false;
            }

            try
            {
                // Define the path to the protein mapping file, but don't create it yet; just delete it if it exists
                // We'll only create it if two or more proteins have the same protein sequence
                duplicateProteinMappingFileOut =
                    Path.Combine(Path.GetDirectoryName(fastaFilePathToCheck),
                    Path.GetFileNameWithoutExtension(fastaFilePathToCheck) + "_UniqueProteinSeqDuplicates.txt");                       // Look for duplicateProteinMappingFileOut and erase it if it exists

                if (File.Exists(duplicateProteinMappingFileOut))
                {
                    File.Delete(duplicateProteinMappingFileOut);
                }
            }
            catch (Exception ex)
            {
                // Error deleting output file
                RecordFastaFileError(0, 0, string.Empty, (int)eMessageCodeConstants.UnspecifiedError,
                    "Error deleting output file " + duplicateProteinMappingFileOut + ": " + ex.Message, string.Empty);
                OnErrorEvent("Error deleting output file (SaveHashInfo to _UniqueProteinSeqDuplicates.txt)", ex);
                return false;
            }

            try
            {
                var headerColumns = new List<string>()
                {
                    "Sequence_Index",
                    "Protein_Name_First",
                    SEQUENCE_LENGTH_COLUMN,
                    SEQUENCE_HASH_COLUMN,
                    "Protein_Count",
                    "Duplicate_Proteins"
                };

                swUniqueProteinSeqsOut.WriteLine(FlattenList(headerColumns));

                for (index = 0; index <= proteinSequenceHashCount - 1; index++)
                {
                    var proteinHashInfo = proteinSeqHashInfo[index];

                    var dataValues = new List<string>()
                    {
                        (index + 1).ToString(),
                        proteinHashInfo.ProteinNameFirst,
                        proteinHashInfo.SequenceLength.ToString(),
                        proteinHashInfo.SequenceHash,
                        (proteinHashInfo.AdditionalProteins.Count() + 1).ToString(),
                        FlattenArray(proteinHashInfo.AdditionalProteins, ',')
                    };

                    swUniqueProteinSeqsOut.WriteLine(FlattenList(dataValues));

                    if (proteinHashInfo.AdditionalProteins.Count() > 0)
                    {
                        duplicateProteinSeqsFound = true;

                        if (swDuplicateProteinMapping == null)
                        {
                            // Need to create swDuplicateProteinMapping
                            swDuplicateProteinMapping = new StreamWriter(new FileStream(duplicateProteinMappingFileOut, FileMode.Create, FileAccess.Write, FileShare.ReadWrite));

                            var proteinHeaderColumns = new List<string>()
                            {
                                "Sequence_Index",
                                "Protein_Name_First",
                                SEQUENCE_LENGTH_COLUMN,
                                "Duplicate_Protein"
                            };

                            swDuplicateProteinMapping.WriteLine(FlattenList(proteinHeaderColumns));
                        }

                        foreach (string additionalProtein in proteinHashInfo.AdditionalProteins)
                        {
                            if (proteinHashInfo.AdditionalProteins.ElementAtOrDefault(duplicateIndex) != null)
                            {
                                if (additionalProtein.Trim().Length > 0)
                                {
                                    var proteinDataValues = new List<string>()
                                    {
                                        (index + 1).ToString(),
                                        proteinHashInfo.ProteinNameFirst,
                                        proteinHashInfo.SequenceLength.ToString(),
                                        additionalProtein
                                    };

                                    swDuplicateProteinMapping.WriteLine(FlattenList(proteinDataValues));
                                }
                            }
                        }
                    }
                    else if (proteinHashInfo.DuplicateProteinNameCount > 0)
                    {
                        duplicateProteinSeqsFound = true;
                    }
                }

                swUniqueProteinSeqsOut.Close();
                if (swDuplicateProteinMapping != null)
                    swDuplicateProteinMapping.Close();

                success = true;
            }
            catch (Exception ex)
            {
                // Error writing results
                RecordFastaFileError(0, 0, string.Empty, (int)eMessageCodeConstants.UnspecifiedError,
                    "Error writing results to " + uniqueProteinSeqsFileOut + " or " + duplicateProteinMappingFileOut + ": " + ex.Message, string.Empty);
                OnErrorEvent("Error writing results to " + uniqueProteinSeqsFileOut + " or " + duplicateProteinMappingFileOut, ex);
                success = false;
            }

            if (success && proteinSequenceHashCount > 0 && duplicateProteinSeqsFound)
            {
                if (consolidateDuplicateProteinSeqsInFasta || keepDuplicateNamedProteinsUnlessMatchingSequence)
                {
                    success = CorrectForDuplicateProteinSeqsInFasta(consolidateDuplicateProteinSeqsInFasta, consolidateDupsIgnoreILDiff, fastaFilePathOut, proteinSequenceHashCount, proteinSeqHashInfo);
                }
            }

            return success;
        }

        /// <summary>
        /// Pre-scan a portion of the Fasta file to determine the appropriate value for mProteinNameSpannerCharLength
        /// </summary>
        /// <param name="fastaFilePathToTest">Fasta file to examine</param>
        /// <param name="terminatorSize">Linefeed length (1 for LF or 2 for CRLF)</param>
        /// <remarks>
        /// Reads 50 MB chunks from 10 sections of the Fasta file (or the entire Fasta file if under 500 MB in size)
        /// Keeps track of the portion of protein names in common between adjacent proteins
        /// Uses this information to determine an appropriate value for mProteinNameSpannerCharLength
        /// </remarks>
        private void AutoDetermineFastaProteinNameSpannerCharLength(string fastaFilePathToTest, int terminatorSize)
        {
            const int PARTS_TO_SAMPLE = 10;
            const int KILOBYTES_PER_SAMPLE = 51200;

            var proteinStartLetters = new Dictionary<string, int>();
            var startTime = DateTime.UtcNow;
            bool showStats = false;

            var fastaFile = new FileInfo(fastaFilePathToTest);
            if (!fastaFile.Exists)
                return;
            long fullScanLengthBytes = 1024L * PARTS_TO_SAMPLE * KILOBYTES_PER_SAMPLE;
            long linesReadTotal = 0;

            if (fastaFile.Length < fullScanLengthBytes)
            {
                fullScanLengthBytes = fastaFile.Length;
                linesReadTotal = AutoDetermineFastaProteinNameSpannerCharLength(fastaFile, terminatorSize, proteinStartLetters, 0, fastaFile.Length);
            }
            else
            {
                long stepSizeBytes = (long) Math.Round(fastaFile.Length / (double)PARTS_TO_SAMPLE, 0);

                for (long byteOffsetStart = 0; byteOffsetStart <= fastaFile.Length; byteOffsetStart += stepSizeBytes)
                {
                    long linesRead = AutoDetermineFastaProteinNameSpannerCharLength(fastaFile, terminatorSize, proteinStartLetters, byteOffsetStart, KILOBYTES_PER_SAMPLE * 1024);

                    if (linesRead < 0)
                    {
                        // This indicates an error, probably from a corrupt file; do not read further
                        break;
                    }

                    linesReadTotal += linesRead;

                    if (!showStats && DateTime.UtcNow.Subtract(startTime).TotalMilliseconds > 500)
                    {
                        showStats = true;
                        ShowMessage("Pre-scanning the file to look for common base protein names");
                    }
                }
            }

            if (proteinStartLetters.Count == 0)
            {
                mProteinNameSpannerCharLength = 1;
            }
            else
            {
                int preScanProteinCount = (from item in proteinStartLetters select item.Value).Sum();

                if (showStats)
                {
                    double percentFileProcessed = fullScanLengthBytes / (double)fastaFile.Length * 100;

                    ShowMessage(string.Format(
                        "  parsed {0:0}% of the file, reading {1:#,##0} lines and finding {2:#,##0} proteins",
                        percentFileProcessed, linesReadTotal, preScanProteinCount));
                }

                // Determine the appropriate spanner length given the observation counts of the base names
                mProteinNameSpannerCharLength = clsNestedStringIntList.DetermineSpannerLengthUsingStartLetterStats(proteinStartLetters);
            }

            ShowMessage("Using ProteinNameSpannerCharLength = " + mProteinNameSpannerCharLength);
            Console.WriteLine();
        }

        /// <summary>
        /// Read a portion of the Fasta file, comparing adjacent protein names and keeping track of the name portions in common
        /// </summary>
        /// <param name="fastaFile"></param>
        /// <param name="terminatorSize"></param>
        /// <param name="proteinStartLetters"></param>
        /// <param name="startOffset"></param>
        /// <param name="bytesToRead"></param>
        /// <returns>The number of lines read</returns>
        /// <remarks></remarks>
        private long AutoDetermineFastaProteinNameSpannerCharLength(
            FileInfo fastaFile,
            int terminatorSize,
            IDictionary<string, int> proteinStartLetters,
            long startOffset,
            long bytesToRead)
        {
            long linesRead = 0L;

            try
            {
                int previousProteinLength = 0;
                string previousProteinName = string.Empty;

                if (startOffset >= fastaFile.Length)
                {
                    ShowMessage("Ignoring byte offset of " + startOffset +
                        " in AutoDetermineProteinNameSpannerCharLength since past the end of the file " +
                        "(" + fastaFile.Length + " bytes");
                    return 0;
                }

                long bytesRead = 0;

                var firstLineDiscarded = false;
                if (startOffset == 0)
                {
                    firstLineDiscarded = true;
                }

                using (var inStream = new FileStream(fastaFile.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    inStream.Position = startOffset;

                    using (var reader = new StreamReader(inStream))
                    {
                        while (!reader.EndOfStream)
                        {
                            string lineIn = reader.ReadLine();
                            bytesRead += terminatorSize;
                            linesRead += 1;

                            if (string.IsNullOrEmpty(lineIn))
                            {
                                continue;
                            }

                            bytesRead += lineIn.Length;
                            if (!firstLineDiscarded)
                            {
                                // We can't trust that this was a full line of text; skip it
                                firstLineDiscarded = true;
                                continue;
                            }

                            if (bytesRead > bytesToRead)
                            {
                                break;
                            }

                            if (!(lineIn[0] == mProteinLineStartChar))
                            {
                                continue;
                            }

                            // Make sure the protein name and description are valid
                            // Find the first space and/or tab
                            int spaceIndex = GetBestSpaceIndex(lineIn);
                            string proteinName;

                            if (spaceIndex > 1)
                            {
                                proteinName = lineIn.Substring(1, spaceIndex - 1);
                            }
                            else if (spaceIndex <= 0)
                            {
                                // Line does not contain a description
                                if (lineIn.Trim().Length <= 1)
                                {
                                    continue;
                                }
                                else
                                {
                                    // The line contains a protein name, but not a description
                                    proteinName = lineIn.Substring(1);
                                }
                            }
                            else
                            {
                                // Space or tab found directly after the > symbol
                                continue;
                            }

                            if (previousProteinLength == 0)
                            {
                                previousProteinName = string.Copy(proteinName);
                                previousProteinLength = previousProteinName.Length;
                                continue;
                            }

                            int currentNameLength = proteinName.Length;
                            int charIndex = 0;

                            while (charIndex < previousProteinLength)
                            {
                                if (charIndex >= currentNameLength)
                                {
                                    break;
                                }

                                if (previousProteinName[charIndex] != proteinName[charIndex])
                                {
                                    // Difference found; add/update the dictionary
                                    break;
                                }

                                charIndex += 1;
                            }

                            int charsInCommon = charIndex;
                            if (charsInCommon > 0)
                            {
                                string baseName = previousProteinName.Substring(0, charsInCommon);
                                int matchCount = 0;

                                if (proteinStartLetters.TryGetValue(baseName, out matchCount))
                                {
                                    proteinStartLetters[baseName] = matchCount + 1;
                                }
                                else
                                {
                                    proteinStartLetters.Add(baseName, 1);
                                }
                            }

                            previousProteinName = string.Copy(proteinName);
                            previousProteinLength = previousProteinName.Length;
                        }
                    }
                }
            }
            catch (OutOfMemoryException ex)
            {
                OnErrorEvent("Out of memory exception in AutoDetermineProteinNameSpannerCharLength", ex);

                // Example message: Insufficient memory to continue the execution of the program
                // This can happen with a corrupt .fasta file with a line that has millions of characters
                return -1;
            }
            catch (Exception ex)
            {
                OnErrorEvent("Error in AutoDetermineProteinNameSpannerCharLength", ex);
            }

            return linesRead;
        }

        private string AutoFixProteinNameAndDescription(
            ref string proteinName,
            ref string proteinDescription,
            udtProteinNameTruncationRegex reProteinNameTruncation)
        {
            bool proteinNameTooLong;
            Match reMatch;
            string newProteinName;
            int charIndex;
            int minCharIndex;
            string extraProteinNameText;

            bool multipleRefsSplitOutFromKnownAccession = false;

            // Auto-fix potential errors in the protein name

            // Possibly truncate to mMaximumProteinNameLength characters
            if (proteinName.Length > mMaximumProteinNameLength)
            {
                proteinNameTooLong = true;
            }
            else
            {
                proteinNameTooLong = false;
            }

            if (mFixedFastaOptions.SplitOutMultipleRefsForKnownAccession ||
                mFixedFastaOptions.TruncateLongProteinNames && proteinNameTooLong)
            {

                // First see if the name fits the pattern IPI:IPI00048500.11|
                // Next see if the name fits the pattern gi|7110699|
                // Next see if the name fits the pattern jgi
                // Next see if the name fits the generic pattern defined by reProteinNameTruncation.reMatchGeneric
                // Otherwise, use mFixedFastaOptions.LongProteinNameSplitChars to define where to truncate

                newProteinName = string.Copy(proteinName);
                extraProteinNameText = string.Empty;

                reMatch = reProteinNameTruncation.reMatchIPI.Match(proteinName);
                if (reMatch.Success)
                {
                    multipleRefsSplitOutFromKnownAccession = true;
                }
                else
                {
                    // IPI didn't match; try gi
                    reMatch = reProteinNameTruncation.reMatchGI.Match(proteinName);
                }

                if (reMatch.Success)
                {
                    multipleRefsSplitOutFromKnownAccession = true;
                }
                else
                {
                    // GI didn't match; try jgi
                    reMatch = reProteinNameTruncation.reMatchJGI.Match(proteinName);
                }

                if (reMatch.Success)
                {
                    multipleRefsSplitOutFromKnownAccession = true;
                }
                // jgi didn't match; try generic (text separated by a series of colons or bars),
                // but only if the name is too long
                else if (mFixedFastaOptions.TruncateLongProteinNames && proteinNameTooLong)
                {
                    reMatch = reProteinNameTruncation.reMatchGeneric.Match(proteinName);
                }

                if (reMatch.Success)
                {
                    // Truncate the protein name, but move the truncated portion into the next group
                    newProteinName = reMatch.Groups[1].Value;
                    extraProteinNameText = reMatch.Groups[2].Value;
                }
                else if (mFixedFastaOptions.TruncateLongProteinNames && proteinNameTooLong)
                {
                    // Name is too long, but it didn't match the known patterns
                    // Find the last occurrence of mFixedFastaOptions.LongProteinNameSplitChars (default is vertical bar)
                    // and truncate the text following the match
                    // Repeat the process until the protein name length >= mMaximumProteinNameLength

                    // See if any of the characters in proteinNameSplitChars is present after
                    // character 6 but less than character mMaximumProteinNameLength
                    minCharIndex = 6;

                    do
                    {
                        charIndex = newProteinName.LastIndexOfAny(mFixedFastaOptions.LongProteinNameSplitChars);
                        if (charIndex >= minCharIndex)
                        {
                            if (extraProteinNameText.Length > 0)
                            {
                                extraProteinNameText = "|" + extraProteinNameText;
                            }

                            extraProteinNameText = newProteinName.Substring(charIndex + 1) + extraProteinNameText;
                            newProteinName = newProteinName.Substring(0, charIndex);
                        }
                        else
                        {
                            charIndex = -1;
                        }
                    }
                    while (charIndex > 0 && newProteinName.Length > mMaximumProteinNameLength);
                }

                if (extraProteinNameText.Length > 0)
                {
                    if (proteinNameTooLong)
                    {
                        mFixedFastaStats.TruncatedProteinNameCount += 1;
                    }
                    else
                    {
                        mFixedFastaStats.ProteinNamesMultipleRefsRemoved += 1;
                    }

                    proteinName = string.Copy(newProteinName);

                    PrependExtraTextToProteinDescription(extraProteinNameText, ref proteinDescription);
                }
            }

            if (mFixedFastaOptions.ProteinNameInvalidCharsToRemove.Length > 0)
            {
                newProteinName = string.Copy(proteinName);

                // First remove invalid characters from the beginning or end of the protein name
                newProteinName = newProteinName.Trim(mFixedFastaOptions.ProteinNameInvalidCharsToRemove);

                if (newProteinName.Length >= 1)
                {
                    foreach (var invalidChar in mFixedFastaOptions.ProteinNameInvalidCharsToRemove)
                        // Next, replace any remaining instances of the character with an underscore
                        newProteinName = newProteinName.Replace(invalidChar, INVALID_PROTEIN_NAME_CHAR_REPLACEMENT);

                    if ((proteinName ?? "") != (newProteinName ?? ""))
                    {
                        if (newProteinName.Length >= 3)
                        {
                            proteinName = string.Copy(newProteinName);
                            mFixedFastaStats.ProteinNamesInvalidCharsReplaced += 1;
                        }
                    }
                }
            }

            if (mFixedFastaOptions.SplitOutMultipleRefsInProteinName && !multipleRefsSplitOutFromKnownAccession)
            {
                // Look for multiple refs in the protein name, but only if we didn't already split out multiple refs above

                reMatch = reProteinNameTruncation.reMatchDoubleBarOrColonAndBar.Match(proteinName);
                if (reMatch.Success)
                {
                    // Protein name contains 2 or more vertical bars, or a colon and a bar
                    // Split out the multiple refs and place them in the description
                    // However, jgi names are supposed to have two vertical bars, so we need to treat that data differently

                    extraProteinNameText = string.Empty;

                    reMatch = reProteinNameTruncation.reMatchJGIBaseAndID.Match(proteinName);
                    if (reMatch.Success)
                    {
                        // ProteinName is similar to jgi|Organism|00000
                        // Check whether there is any text following the match
                        if (reMatch.Length < proteinName.Length)
                        {
                            // Extra text exists; populate extraProteinNameText
                            extraProteinNameText = proteinName.Substring(reMatch.Length + 1);
                            proteinName = reMatch.ToString();
                        }
                    }
                    else
                    {
                        // Find the first vertical bar or colon
                        charIndex = proteinName.IndexOfAny(mProteinNameFirstRefSepChars);

                        if (charIndex > 0)
                        {
                            // Find the second vertical bar, colon, or semicolon
                            charIndex = proteinName.IndexOfAny(mProteinNameSubsequentRefSepChars, charIndex + 1);

                            if (charIndex > 0)
                            {
                                // Split the protein name
                                extraProteinNameText = proteinName.Substring(charIndex + 1);
                                proteinName = proteinName.Substring(0, charIndex);
                            }
                        }
                    }

                    if (extraProteinNameText.Length > 0)
                    {
                        PrependExtraTextToProteinDescription(extraProteinNameText, ref proteinDescription);
                        mFixedFastaStats.ProteinNamesMultipleRefsRemoved += 1;
                    }
                }
            }

            // Make sure proteinDescription doesn't start with a | or space
            if (proteinDescription.Length > 0)
            {
                proteinDescription = proteinDescription.TrimStart(new char[] { '|', ' ' });
            }

            return proteinName;
        }

        private string BoolToStringInt(bool value)
        {
            if (value)
            {
                return "1";
            }
            else
            {
                return "0";
            }
        }

        private string CharArrayToString(IEnumerable<char> charArray)
        {
            return string.Join("", charArray);
        }

        private void ClearAllRules()
        {
            ClearRules(RuleTypes.HeaderLine);
            ClearRules(RuleTypes.ProteinDescription);
            ClearRules(RuleTypes.ProteinName);
            ClearRules(RuleTypes.ProteinSequence);

            mMasterCustomRuleID = CUSTOM_RULE_ID_START;
        }

        private void ClearRules(RuleTypes ruleType)
        {
            switch (ruleType)
            {
                case RuleTypes.HeaderLine:
                    ClearRulesDataStructure(ref mHeaderLineRules);
                    break;
                case RuleTypes.ProteinDescription:
                    ClearRulesDataStructure(ref mProteinDescriptionRules);
                    break;
                case RuleTypes.ProteinName:
                    ClearRulesDataStructure(ref mProteinNameRules);
                    break;
                case RuleTypes.ProteinSequence:
                    ClearRulesDataStructure(ref mProteinSequenceRules);
                    break;
            }
        }

        private void ClearRulesDataStructure(ref udtRuleDefinitionType[] rules)
        {
            rules = new udtRuleDefinitionType[0];
        }

        public string ComputeProteinHash(StringBuilder sbResidues, bool consolidateDupsIgnoreILDiff)
        {
            if (sbResidues.Length > 0)
            {
                // Compute the hash value for sbCurrentResidues
                if (consolidateDupsIgnoreILDiff)
                {
                    return HashUtilities.ComputeStringHashSha1(sbResidues.ToString().Replace('L', 'I')).ToUpper();
                }
                else
                {
                    return HashUtilities.ComputeStringHashSha1(sbResidues.ToString()).ToUpper();
                }
            }
            else
            {
                return string.Empty;
            }
        }

        private int ComputeTotalSpecifiedCount(udtItemSummaryIndexedType errorStats)
        {
            int total;
            int index;

            total = 0;
            for (index = 0; index <= errorStats.ErrorStatsCount - 1; index++)
                total += errorStats.ErrorStats[index].CountSpecified;

            return total;
        }

        private int ComputeTotalUnspecifiedCount(udtItemSummaryIndexedType errorStats)
        {
            int total;
            int index;

            total = 0;
            for (index = 0; index <= errorStats.ErrorStatsCount - 1; index++)
                total += errorStats.ErrorStats[index].CountUnspecified;

            return total;
        }

        /// <summary>
        /// Looks for duplicate proteins in the Fasta file
        /// Creates a new fasta file that has exact duplicates removed
        /// Will consolidate proteins with the same sequence if consolidateDuplicateProteinSeqsInFasta=True
        /// </summary>
        /// <param name="consolidateDuplicateProteinSeqsInFasta"></param>
        /// <param name="fixedFastaFilePath"></param>
        /// <param name="proteinSequenceHashCount"></param>
        /// <param name="proteinSeqHashInfo"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool CorrectForDuplicateProteinSeqsInFasta(
            bool consolidateDuplicateProteinSeqsInFasta,
            bool consolidateDupsIgnoreILDiff,
            string fixedFastaFilePath,
            int proteinSequenceHashCount,
            IList<clsProteinHashInfo> proteinSeqHashInfo)
        {
            Stream fsInFile;
            StreamWriter consolidatedFastaWriter = null;
            long bytesRead;

            int terminatorSize;
            float percentComplete;
            int lineCountRead;

            string fixedFastaFilePathTemp = string.Empty;
            string lineIn;

            string cachedProteinName = string.Empty;
            string cachedProteinDescription = string.Empty;
            var sbCachedProteinResidueLines = new StringBuilder(250);
            var sbCachedProteinResidues = new StringBuilder(250);

            // This list contains the protein names that we will keep; values are the index values pointing into proteinSeqHashInfo
            // If consolidateDuplicateProteinSeqsInFasta=False, this will contain all protein names
            // If consolidateDuplicateProteinSeqsInFasta=True, we only keep the first name found for a given sequence
            clsNestedStringDictionary<int> proteinNameFirst;

            // This list keeps track of the protein names that have been written out to the new fasta file
            // Keys are the protein names; values are the index of the entry in proteinSeqHashInfo()
            clsNestedStringDictionary<int> proteinsWritten;

            // This list contains the names of duplicate proteins; the hash values are the protein names of the master protein that has the same sequence
            clsNestedStringDictionary<string> duplicateProteinList;

            int descriptionStartIndex;

            bool success;

            if (proteinSequenceHashCount <= 0)
            {
                return true;
            }

            // '''''''''''''''''''''
            // Processing Steps
            // '''''''''''''''''''''
            //
            // Open fixedFastaFilePath with the fasta file reader
            // Create a new fasta file with a writer

            // For each protein, check whether it has duplicates
            // If not, just write it out to the new fasta file

            // If it does have duplicates and it is the master, then append the duplicate protein names to the end of the description for the protein
            // and write out the name, new description, and sequence to the new fasta file

            // Otherwise, check if it is a duplicate of a master protein
            // If it is, then do not write the name, description, or sequence to the new fasta file

            try
            {
                fixedFastaFilePathTemp = fixedFastaFilePath + ".TempFixed";

                if (File.Exists(fixedFastaFilePathTemp))
                {
                    File.Delete(fixedFastaFilePathTemp);
                }

                File.Move(fixedFastaFilePath, fixedFastaFilePathTemp);
            }
            catch (Exception ex)
            {
                RecordFastaFileError(0, 0, string.Empty, (int)eMessageCodeConstants.UnspecifiedError,
                    "Error renaming " + fixedFastaFilePath + " to " + fixedFastaFilePathTemp + ": " + ex.Message, string.Empty);
                OnErrorEvent("Error renaming fixed fasta to .tempfixed", ex);
                return false;
            }

            StreamReader fastaReader;

            try
            {
                // Open the file and read, at most, the first 100,000 characters to see if it contains CrLf or just Lf
                terminatorSize = DetermineLineTerminatorSize(fixedFastaFilePathTemp);

                // Open the Fixed fasta file
                fsInFile = new FileStream(
                    fixedFastaFilePathTemp,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.ReadWrite);

                fastaReader = new StreamReader(fsInFile);
            }
            catch (Exception ex)
            {
                RecordFastaFileError(0, 0, string.Empty, (int)eMessageCodeConstants.UnspecifiedError,
                    "Error opening " + fixedFastaFilePathTemp + ": " + ex.Message, string.Empty);
                OnErrorEvent("Error opening fixedFastaFilePathTemp", ex);
                return false;
            }

            try
            {
                // Create the new fasta file
                consolidatedFastaWriter = new StreamWriter(new FileStream(fixedFastaFilePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite));
            }
            catch (Exception ex)
            {
                RecordFastaFileError(0, 0, string.Empty, (int)eMessageCodeConstants.UnspecifiedError,
                    "Error creating consolidated fasta output file " + fixedFastaFilePath + ": " + ex.Message, string.Empty);
                OnErrorEvent("Error creating consolidated fasta output file", ex);
            }

            try
            {
                // Populate proteinNameFirst with the protein names in proteinSeqHashInfo().ProteinNameFirst
                proteinNameFirst = new clsNestedStringDictionary<int>(true, mProteinNameSpannerCharLength);

                // Populate htDuplicateProteinList with the protein names in proteinSeqHashInfo().AdditionalProteins
                duplicateProteinList = new clsNestedStringDictionary<string>(true, mProteinNameSpannerCharLength);

                for (int index = 0; index <= proteinSequenceHashCount - 1; index++)
                {
                    var proteinHashInfo = proteinSeqHashInfo[index];

                    if (!proteinNameFirst.ContainsKey(proteinHashInfo.ProteinNameFirst))
                    {
                        proteinNameFirst.Add(proteinHashInfo.ProteinNameFirst, index);
                    }
                    else
                    {
                        // .ProteinNameFirst is already present in proteinNameFirst
                        // The fixed fasta file will only actually contain the first occurrence of .ProteinNameFirst, so we can effectively ignore this entry
                        // but we should increment the DuplicateNameSkipCount

                    }

                    if (proteinHashInfo.AdditionalProteins.Count() > 0)
                    {
                        foreach (string additionalProtein in proteinHashInfo.AdditionalProteins)
                        {
                            if (consolidateDuplicateProteinSeqsInFasta)
                            {
                                // Update the duplicate protein name list
                                if (!duplicateProteinList.ContainsKey(additionalProtein))
                                {
                                    duplicateProteinList.Add(additionalProtein, proteinHashInfo.ProteinNameFirst);
                                }
                            }
                            // We are not consolidating proteins with the same sequence but different protein names
                            // Append this entry to proteinNameFirst

                            else if (!proteinNameFirst.ContainsKey(additionalProtein))
                            {
                                proteinNameFirst.Add(additionalProtein, index);
                            }
                            else
                            {
                                // .AdditionalProteins(dupIndex) is already present in proteinNameFirst
                                // Increment the DuplicateNameSkipCount
                            }
                        }
                    }
                }

                proteinsWritten = new clsNestedStringDictionary<int>(false, mProteinNameSpannerCharLength);

                var lastMemoryUsageReport = DateTime.UtcNow;

                // Parse each line in the file
                lineCountRead = 0;
                bytesRead = 0;
                mFixedFastaStats.DuplicateSequenceProteinsSkipped = 0;

                while (!fastaReader.EndOfStream)
                {
                    lineIn = fastaReader.ReadLine();
                    bytesRead += lineIn.Length + terminatorSize;

                    if (lineCountRead % 50 == 0)
                    {
                        if (AbortProcessing)
                            break;
                        percentComplete = 75 + (float)(bytesRead / (double)fastaReader.BaseStream.Length * 100.0) / 4;
                        UpdateProgress("Consolidating duplicate proteins to create a new FASTA File (" + Math.Round(percentComplete, 0) + "% Done)", percentComplete);

                        if (DateTime.UtcNow.Subtract(lastMemoryUsageReport).TotalMinutes >= 1)
                        {
                            lastMemoryUsageReport = DateTime.UtcNow;
                            ReportMemoryUsage(proteinNameFirst, proteinsWritten, duplicateProteinList);
                        }
                    }

                    lineCountRead += 1;

                    if (lineIn != null)
                    {
                        if (lineIn.Trim().Length > 0)
                        {
                            // Note: Trim the start of the line (however, since this is a fixed fasta file it should not start with a space)
                            lineIn = lineIn.TrimStart();

                            if (lineIn[0] == mProteinLineStartChar)
                            {
                                // Protein entry line

                                if (!string.IsNullOrEmpty(cachedProteinName))
                                {
                                    // Write out the cached protein and it's residues

                                    WriteCachedProtein(
                                        cachedProteinName, cachedProteinDescription,
                                        consolidatedFastaWriter, proteinSeqHashInfo,
                                        sbCachedProteinResidueLines, sbCachedProteinResidues,
                                        consolidateDuplicateProteinSeqsInFasta, consolidateDupsIgnoreILDiff,
                                        proteinNameFirst, duplicateProteinList,
                                        lineCountRead, proteinsWritten);

                                    cachedProteinName = string.Empty;
                                    sbCachedProteinResidueLines.Length = 0;
                                    sbCachedProteinResidues.Length = 0;
                                }

                                // Extract the protein name and description
                                SplitFastaProteinHeaderLine(lineIn, out cachedProteinName, out cachedProteinDescription, out descriptionStartIndex);
                            }
                            else
                            {
                                // Protein residues
                                sbCachedProteinResidueLines.AppendLine(lineIn);
                                sbCachedProteinResidues.Append(lineIn.Trim());
                            }
                        }
                    }
                }

                if (!string.IsNullOrEmpty(cachedProteinName))
                {
                    // Write out the cached protein and it's residues
                    WriteCachedProtein(
                        cachedProteinName, cachedProteinDescription,
                        consolidatedFastaWriter, proteinSeqHashInfo,
                        sbCachedProteinResidueLines, sbCachedProteinResidues,
                        consolidateDuplicateProteinSeqsInFasta, consolidateDupsIgnoreILDiff,
                        proteinNameFirst, duplicateProteinList,
                        lineCountRead, proteinsWritten);
                }

                ReportMemoryUsage(proteinNameFirst, proteinsWritten, duplicateProteinList);
                success = true;
            }
            catch (Exception ex)
            {
                RecordFastaFileError(0, 0, string.Empty, (int)eMessageCodeConstants.UnspecifiedError,
                    "Error writing to consolidated fasta file " + fixedFastaFilePath + ": " + ex.Message, string.Empty);
                OnErrorEvent("Error writing to consolidated fasta file", ex);
                return false;
            }
            finally
            {
                try
                {
                    if (fastaReader != null)
                        fastaReader.Close();
                    if (consolidatedFastaWriter != null)
                        consolidatedFastaWriter.Close();

                    System.Threading.Thread.Sleep(100);

                    File.Delete(fixedFastaFilePathTemp);
                }
                catch (Exception ex)
                {
                    // Ignore errors here
                    OnWarningEvent("Error closing file handles in CorrectForDuplicateProteinSeqsInFasta: " + ex.Message);
                }
            }

            return success;
        }

        private string ConstructFastaHeaderLine(string proteinName, string proteinDescription)
        {
            if (proteinName == null)
                proteinName = "????";

            if (string.IsNullOrWhiteSpace(proteinDescription))
            {
                return mProteinLineStartChar + proteinName;
            }
            else
            {
                return mProteinLineStartChar + proteinName + " " + proteinDescription;
            }
        }

        private string ConstructStatsFilePath(string outputFolderPath)
        {
            string outFilePath = string.Empty;

            try
            {
                // Record the current time in now
                outFilePath = "FastaFileStats_" + DateTime.Now.ToString("yyyy-MM-dd") + ".txt";

                if (outputFolderPath != null && outputFolderPath.Length > 0)
                {
                    outFilePath = Path.Combine(outputFolderPath, outFilePath);
                }
            }
            catch (Exception ex)
            {
                if (string.IsNullOrWhiteSpace(outFilePath))
                {
                    outFilePath = "FastaFileStats.txt";
                }
            }

            return outFilePath;
        }

        private void DeleteTempFiles()
        {
            if (mTempFilesToDelete != null && mTempFilesToDelete.Count > 0)
            {
                foreach (var filePath in mTempFilesToDelete)
                {
                    try
                    {
                        if (File.Exists(filePath))
                        {
                            File.Delete(filePath);
                        }
                    }
                    catch (Exception ex)
                    {
                        // Ignore errors
                    }
                }
            }
        }

        private int DetermineLineTerminatorSize(string inputFilePath)
        {
            var endCharType = DetermineLineTerminatorType(inputFilePath);

            switch (endCharType)
            {
                case eLineEndingCharacters.CR:
                    return 1;
                case eLineEndingCharacters.LF:
                    return 1;
                case eLineEndingCharacters.CRLF:
                    return 2;
                case eLineEndingCharacters.LFCR:
                    return 2;
            }

            return 2;
        }

        private eLineEndingCharacters DetermineLineTerminatorType(string inputFilePath)
        {
            int oneByte;

            var endCharacterType = default(eLineEndingCharacters);

            try
            {
                // Open the input file and look for the first carriage return or line feed
                using (var fsInFile = new FileStream(inputFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    while (fsInFile.Position < fsInFile.Length && fsInFile.Position < 100000)
                    {
                        oneByte = fsInFile.ReadByte();

                        if (oneByte == 10)
                        {
                            // Found linefeed
                            if (fsInFile.Position < fsInFile.Length)
                            {
                                oneByte = fsInFile.ReadByte();
                                if (oneByte == 13)
                                {
                                    // LfCr
                                    endCharacterType = eLineEndingCharacters.LFCR;
                                }
                                else
                                {
                                    // Lf only
                                    endCharacterType = eLineEndingCharacters.LF;
                                }
                            }
                            else
                            {
                                endCharacterType = eLineEndingCharacters.LF;
                            }

                            break;
                        }
                        else if (oneByte == 13)
                        {
                            // Found carriage return
                            if (fsInFile.Position < fsInFile.Length)
                            {
                                oneByte = fsInFile.ReadByte();
                                if (oneByte == 10)
                                {
                                    // CrLf
                                    endCharacterType = eLineEndingCharacters.CRLF;
                                }
                                else
                                {
                                    // Cr only
                                    endCharacterType = eLineEndingCharacters.CR;
                                }
                            }
                            else
                            {
                                endCharacterType = eLineEndingCharacters.CR;
                            }

                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                SetLocalErrorCode(eValidateFastaFileErrorCodes.ErrorVerifyingLinefeedAtEOF);
            }

            return endCharacterType;
        }

        // Unused function
        //private string ExtractListItem(string list, int item)
        //{
        //    string itemStr = string.Empty;
        //
        //    if (item >= 1 && list != null)
        //    {
        //        var items = list.Split(',');
        //        if (items.Length >= item)
        //        {
        //            itemStr = items[item - 1];
        //        }
        //    }
        //
        //    return itemStr;
        //}

        private string NormalizeFileLineEndings(
            string pathOfFileToFix,
            string newFileName,
            eLineEndingCharacters desiredLineEndCharacterType)
        {
            string newEndChar = "\r\n";

            var endCharType = DetermineLineTerminatorType(pathOfFileToFix);

            var origEndCharCount = 0;

            if (endCharType != desiredLineEndCharacterType)
            {
                switch (desiredLineEndCharacterType)
                {
                    case eLineEndingCharacters.CRLF:
                        newEndChar = "\r\n";
                        break;
                    case eLineEndingCharacters.CR:
                        newEndChar = "\r";
                        break;
                    case eLineEndingCharacters.LF:
                        newEndChar = "\n";
                        break;
                    case eLineEndingCharacters.LFCR:
                        newEndChar = "\r\n";
                        break;
                }

                switch (endCharType)
                {
                    case eLineEndingCharacters.CR:
                        origEndCharCount = 2;
                        break;
                    case eLineEndingCharacters.CRLF:
                        origEndCharCount = 1;
                        break;
                    case eLineEndingCharacters.LF:
                        origEndCharCount = 1;
                        break;
                    case eLineEndingCharacters.LFCR:
                        origEndCharCount = 2;
                        break;
                }

                if (!Path.IsPathRooted(newFileName))
                {
                    newFileName = Path.Combine(Path.GetDirectoryName(pathOfFileToFix), Path.GetFileName(newFileName));
                }

                var targetFile = new FileInfo(pathOfFileToFix);
                long fileSizeBytes = targetFile.Length;

                var reader = targetFile.OpenText();

                using (var writer = new StreamWriter(new FileStream(newFileName, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)))
                {
                    this.OnProgressUpdate("Normalizing Line Endings...", 0.0F);

                    string dataLine = reader.ReadLine();
                    long linesRead = 0;
                    while (dataLine != null)
                    {
                        writer.Write(dataLine);
                        writer.Write(newEndChar);

                        int currentFilePos = dataLine.Length + origEndCharCount;
                        linesRead += 1;

                        if (linesRead % 1000 == 0)
                        {
                            OnProgressUpdate("Normalizing Line Endings (" +
                                Math.Round(currentFilePos / (double)fileSizeBytes * 100.0, 1).ToString() +
                                " % complete", (float)(currentFilePos / (double)fileSizeBytes * 100));
                        }

                        dataLine = reader.ReadLine();
                    }

                    reader.Close();
                }

                return newFileName;
            }
            else
            {
                return pathOfFileToFix;
            }
        }

        private void EvaluateRules(
            IList<udtRuleDefinitionExtendedType> ruleDetails,
            string proteinName,
            string textToTest,
            int testTextOffsetInLine,
            string entireLine,
            int contextLength)
        {
            int index;
            Match reMatch;
            string extraInfo;
            int charIndexOfMatch;

            for (index = 0; index <= ruleDetails.Count - 1; index++)
            {
                var ruleDetail = ruleDetails[index];

                reMatch = ruleDetail.reRule.Match(textToTest);

                if (ruleDetail.RuleDefinition.MatchIndicatesProblem && reMatch.Success ||
                    !(ruleDetail.RuleDefinition.MatchIndicatesProblem && !reMatch.Success))
                {
                    if (ruleDetail.RuleDefinition.DisplayMatchAsExtraInfo)
                    {
                        extraInfo = reMatch.ToString();
                    }
                    else
                    {
                        extraInfo = string.Empty;
                    }

                    charIndexOfMatch = testTextOffsetInLine + reMatch.Index;
                    if (ruleDetail.RuleDefinition.Severity >= 5)
                    {
                        RecordFastaFileError(mLineCount, charIndexOfMatch, proteinName,
                            ruleDetail.RuleDefinition.CustomRuleID, extraInfo,
                            ExtractContext(entireLine, charIndexOfMatch, contextLength));
                    }
                    else
                    {
                        RecordFastaFileWarning(mLineCount, charIndexOfMatch, proteinName,
                            ruleDetail.RuleDefinition.CustomRuleID, extraInfo,
                            ExtractContext(entireLine, charIndexOfMatch, contextLength));
                    }
                }
            }
        }

        private string ExamineProteinName(
            ref string proteinName,
            ISet<string> proteinNames,
            out bool skipDuplicateProtein,
            ref bool processingDuplicateOrInvalidProtein)
        {
            bool duplicateName = proteinNames.Contains(proteinName);
            skipDuplicateProtein = false;

            if (duplicateName && mGenerateFixedFastaFile)
            {
                if (mFixedFastaOptions.RenameProteinsWithDuplicateNames)
                {
                    char letterToAppend = 'b';
                    int numberToAppend = 0;
                    string newProteinName;

                    do
                    {
                        newProteinName = proteinName + '-' + letterToAppend;
                        if (numberToAppend > 0)
                        {
                            newProteinName += numberToAppend.ToString();
                        }

                        duplicateName = proteinNames.Contains(newProteinName);

                        if (duplicateName)
                        {
                            // Increment letterToAppend to the next letter and then try again to rename the protein
                            if (letterToAppend == 'z')
                            {
                                // We've reached "z"
                                // Change back to "a" but increment numberToAppend
                                letterToAppend = 'a';
                                numberToAppend += 1;
                            }
                            else
                            {
                                // letterToAppend = Chr(Asc(letterToAppend) + 1)
                                letterToAppend = (char)(letterToAppend + 1);
                            }
                        }
                    }
                    while (duplicateName);

                    RecordFastaFileWarning(mLineCount, 1, proteinName, (int)eMessageCodeConstants.RenamedProtein, "--> " + newProteinName, string.Empty);

                    proteinName = string.Copy(newProteinName);
                    mFixedFastaStats.DuplicateNameProteinsRenamed += 1;
                    skipDuplicateProtein = false;
                }
                else if (mFixedFastaOptions.KeepDuplicateNamedProteinsUnlessMatchingSequence)
                {
                    skipDuplicateProtein = false;
                }
                else
                {
                    skipDuplicateProtein = true;
                }
            }

            if (duplicateName)
            {
                if (skipDuplicateProtein || !mGenerateFixedFastaFile)
                {
                    RecordFastaFileError(mLineCount, 1, proteinName, (int)eMessageCodeConstants.DuplicateProteinName);
                    if (mSaveBasicProteinHashInfoFile)
                    {
                        processingDuplicateOrInvalidProtein = false;
                    }
                    else
                    {
                        processingDuplicateOrInvalidProtein = true;
                        mFixedFastaStats.DuplicateNameProteinsSkipped += 1;
                    }
                }
                else
                {
                    RecordFastaFileWarning(mLineCount, 1, proteinName, (int)eMessageCodeConstants.DuplicateProteinName);
                    processingDuplicateOrInvalidProtein = false;
                }
            }
            else
            {
                processingDuplicateOrInvalidProtein = false;
            }

            if (!proteinNames.Contains(proteinName))
            {
                proteinNames.Add(proteinName);
            }

            return proteinName;
        }

        private string ExtractContext(string text, int startIndex)
        {
            return ExtractContext(text, startIndex, DEFAULT_CONTEXT_LENGTH);
        }

        private string ExtractContext(string text, int startIndex, int contextLength)
        {
            // Note that contextLength should be an odd number; if it isn't, we'll add 1 to it

            int contextStartIndex;
            int contextEndIndex;

            if (contextLength % 2 == 0)
            {
                contextLength += 1;
            }
            else if (contextLength < 1)
            {
                contextLength = 1;
            }

            if (text == null)
            {
                return string.Empty;
            }
            else if (text.Length <= 1)
            {
                return text;
            }
            else
            {
                // Define the start index for extracting the context from text
                contextStartIndex = startIndex - (int)Math.Round((contextLength - 1) / (double)2);
                if (contextStartIndex < 0)
                    contextStartIndex = 0;

                // Define the end index for extracting the context from text
                contextEndIndex = Math.Max(startIndex + (int)Math.Round((contextLength - 1) / (double)2), contextStartIndex + contextLength - 1);
                if (contextEndIndex >= text.Length)
                {
                    contextEndIndex = text.Length - 1;
                }

                // Return the context portion of text
                return text.Substring(contextStartIndex, contextEndIndex - contextStartIndex + 1);
            }
        }

        private string FlattenArray(IEnumerable<string> items, char sepChar)
        {
            if (items == null)
            {
                return string.Empty;
            }
            else
            {
                return FlattenArray(items, items.Count(), sepChar);
            }
        }

        private string FlattenArray(IEnumerable<string> items, int dataCount, char sepChar)
        {
            int index;
            string result;

            if (items == null)
            {
                return string.Empty;
            }
            else if (items.Count() == 0 || dataCount <= 0)
            {
                return string.Empty;
            }
            else
            {
                if (dataCount > items.Count())
                {
                    dataCount = items.Count();
                }

                result = items.ElementAtOrDefault(0);
                if (result == null)
                    result = string.Empty;

                for (index = 1; index <= dataCount - 1; index++)
                {
                    if (items.ElementAtOrDefault(index) == null)
                    {
                        result += sepChar.ToString();
                    }
                    else
                    {
                        result += sepChar + items.ElementAtOrDefault(index);
                    }
                }

                return result;
            }
        }

        /// <summary>
        /// Convert a list of strings to a tab-delimited string
        /// </summary>
        /// <param name="dataValues"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        private string FlattenList(IEnumerable<string> dataValues)
        {
            return FlattenArray(dataValues, '\t');
        }

        /// <summary>
        /// Find the first space (or first tab) in the protein header line
        /// </summary>
        /// <param name="headerLine"></param>
        /// <returns></returns>
        /// <remarks>Used for determining protein name</remarks>
        private int GetBestSpaceIndex(string headerLine)
        {
            int spaceIndex = headerLine.IndexOf(' ');
            int tabIndex = headerLine.IndexOf('\t');

            if (spaceIndex == 1)
            {
                // Space found directly after the > symbol
            }
            else if (tabIndex > 0)
            {
                if (tabIndex == 1)
                {
                    // Tab character found directly after the > symbol
                    spaceIndex = tabIndex;
                }
                else if (spaceIndex <= 0 || spaceIndex > 0 && tabIndex < spaceIndex)
                {
                    // Tab character found; does it separate the protein name and description?
                    spaceIndex = tabIndex;
                }
            }

            return spaceIndex;
        }

        public override IList<string> GetDefaultExtensionsToParse()
        {
            var extensionsToParse = new List<string>() { ".fasta" };

            return extensionsToParse;
        }

        public override string GetErrorMessage()
        {
            // Returns "" if no error

            string errorMessage;

            if (ErrorCode == ProcessFilesErrorCodes.LocalizedError ||
                ErrorCode == ProcessFilesErrorCodes.NoError)
            {
                switch (mLocalErrorCode)
                {
                    case eValidateFastaFileErrorCodes.NoError:
                        errorMessage = "";
                        break;
                    case eValidateFastaFileErrorCodes.OptionsSectionNotFound:
                        errorMessage = "The section " + XML_SECTION_OPTIONS + " was not found in the parameter file";
                        break;
                    case eValidateFastaFileErrorCodes.ErrorReadingInputFile:
                        errorMessage = "Error reading input file";
                        break;
                    case eValidateFastaFileErrorCodes.ErrorCreatingStatsFile:
                        errorMessage = "Error creating stats output file";
                        break;
                    case eValidateFastaFileErrorCodes.ErrorVerifyingLinefeedAtEOF:
                        errorMessage = "Error verifying linefeed at end of file";
                        break;
                    case eValidateFastaFileErrorCodes.UnspecifiedError:
                        errorMessage = "Unspecified localized error";
                        break;
                    default:
                        // This shouldn't happen
                        errorMessage = "Unknown error state";
                        break;
                }
            }
            else
            {
                errorMessage = GetBaseClassErrorMessage();
            }

            return errorMessage;
        }

        private string GetFileErrorTextByIndex(int fileErrorIndex, string sepChar)
        {
            string proteinName;

            if (mFileErrorCount <= 0 || fileErrorIndex < 0 || fileErrorIndex >= mFileErrorCount)
            {
                return string.Empty;
            }
            else
            {
                var fileError = mFileErrors[fileErrorIndex];
                if (fileError.ProteinName == null || fileError.ProteinName.Length == 0)
                {
                    proteinName = "N/A";
                }
                else
                {
                    proteinName = string.Copy(fileError.ProteinName);
                }

                return LookupMessageType(eMsgTypeConstants.ErrorMsg) + sepChar +
                    "Line " + fileError.LineNumber.ToString() + sepChar +
                    "Col " + fileError.ColNumber.ToString() + sepChar +
                    proteinName + sepChar +
                    LookupMessageDescription(fileError.MessageCode, fileError.ExtraInfo) + sepChar + fileError.Context;
            }
        }

        private udtMsgInfoType GetFileErrorByIndex(int fileErrorIndex)
        {
            if (mFileErrorCount <= 0 || fileErrorIndex < 0 || fileErrorIndex >= mFileErrorCount)
            {
                return new udtMsgInfoType();
            }
            else
            {
                return mFileErrors[fileErrorIndex];
            }
        }

        /// <summary>
        /// Retrieve the errors reported by the validator
        /// </summary>
        /// <returns></returns>
        /// <remarks>Used by clsCustomValidateFastaFiles</remarks>
        protected List<udtMsgInfoType> GetFileErrors()
        {
            var fileErrors = new List<udtMsgInfoType>();

            for (int i = 0; i <= mFileErrorCount - 1; i++)
                fileErrors.Add(mFileErrors[i]);
            return fileErrors;
        }

        private string GetFileWarningTextByIndex(int fileWarningIndex, string sepChar)
        {
            string proteinName;

            if (mFileWarningCount <= 0 || fileWarningIndex < 0 || fileWarningIndex >= mFileWarningCount)
            {
                return string.Empty;
            }
            else
            {
                var fileWarning = mFileWarnings[fileWarningIndex];
                if (fileWarning.ProteinName == null || fileWarning.ProteinName.Length == 0)
                {
                    proteinName = "N/A";
                }
                else
                {
                    proteinName = string.Copy(fileWarning.ProteinName);
                }

                return LookupMessageType(eMsgTypeConstants.WarningMsg) + sepChar +
                    "Line " + fileWarning.LineNumber.ToString() + sepChar +
                    "Col " + fileWarning.ColNumber.ToString() + sepChar +
                    proteinName + sepChar +
                    LookupMessageDescription(fileWarning.MessageCode, fileWarning.ExtraInfo) +
                    sepChar + fileWarning.Context;
            }
        }

        private udtMsgInfoType GetFileWarningByIndex(int fileWarningIndex)
        {
            if (mFileWarningCount <= 0 || fileWarningIndex < 0 || fileWarningIndex >= mFileWarningCount)
            {
                return new udtMsgInfoType();
            }
            else
            {
                return mFileWarnings[fileWarningIndex];
            }
        }

        /// <summary>
        /// Retrieve the warnings reported by the validator
        /// </summary>
        /// <returns></returns>
        /// <remarks>Used by clsCustomValidateFastaFiles</remarks>
        protected List<udtMsgInfoType> GetFileWarnings()
        {
            var fileWarnings = new List<udtMsgInfoType>();

            for (int i = 0; i <= mFileWarningCount - 1; i++)
                fileWarnings.Add(mFileWarnings[i]);

            return fileWarnings;
        }

        private string GetProcessMemoryUsageWithTimestamp()
        {
            return GetTimeStamp() + "\t" + clsMemoryUsageLogger.GetProcessMemoryUsageMB().ToString("0.0") + " MB in use";
        }

        private string GetTimeStamp()
        {
            // Record the current time
            return DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString();
        }

        private void InitializeLocalVariables()
        {
            mLocalErrorCode = eValidateFastaFileErrorCodes.NoError;

            SetOptionSwitch(SwitchOptions.AddMissingLineFeedAtEOF, false);

            MaximumFileErrorsToTrack = 5;

            MinimumProteinNameLength = DEFAULT_MINIMUM_PROTEIN_NAME_LENGTH;
            MaximumProteinNameLength = DEFAULT_MAXIMUM_PROTEIN_NAME_LENGTH;
            MaximumResiduesPerLine = DEFAULT_MAXIMUM_RESIDUES_PER_LINE;
            ProteinLineStartChar = DEFAULT_PROTEIN_LINE_START_CHAR;

            SetOptionSwitch(SwitchOptions.AllowAsteriskInResidues, false);
            SetOptionSwitch(SwitchOptions.AllowDashInResidues, false);
            SetOptionSwitch(SwitchOptions.WarnBlankLinesBetweenProteins, false);
            SetOptionSwitch(SwitchOptions.WarnLineStartsWithSpace, true);

            SetOptionSwitch(SwitchOptions.CheckForDuplicateProteinNames, true);
            SetOptionSwitch(SwitchOptions.CheckForDuplicateProteinSequences, true);

            SetOptionSwitch(SwitchOptions.GenerateFixedFASTAFile, false);

            SetOptionSwitch(SwitchOptions.FixedFastaSplitOutMultipleRefsForKnownAccession, true);
            SetOptionSwitch(SwitchOptions.SplitOutMultipleRefsInProteinName, false);

            SetOptionSwitch(SwitchOptions.FixedFastaRenameDuplicateNameProteins, false);
            SetOptionSwitch(SwitchOptions.FixedFastaKeepDuplicateNamedProteins, false);

            SetOptionSwitch(SwitchOptions.FixedFastaConsolidateDuplicateProteinSeqs, false);
            SetOptionSwitch(SwitchOptions.FixedFastaConsolidateDupsIgnoreILDiff, false);

            SetOptionSwitch(SwitchOptions.FixedFastaTruncateLongProteinNames, true);
            SetOptionSwitch(SwitchOptions.FixedFastaWrapLongResidueLines, true);
            SetOptionSwitch(SwitchOptions.FixedFastaRemoveInvalidResidues, false);

            SetOptionSwitch(SwitchOptions.SaveProteinSequenceHashInfoFiles, false);

            SetOptionSwitch(SwitchOptions.SaveBasicProteinHashInfoFile, false);

            mProteinNameFirstRefSepChars = DEFAULT_PROTEIN_NAME_FIRST_REF_SEP_CHARS.ToCharArray();
            mProteinNameSubsequentRefSepChars = DEFAULT_PROTEIN_NAME_SUBSEQUENT_REF_SEP_CHARS.ToCharArray();

            mFixedFastaOptions.LongProteinNameSplitChars = new char[] { DEFAULT_LONG_PROTEIN_NAME_SPLIT_CHAR };
            mFixedFastaOptions.ProteinNameInvalidCharsToRemove = new char[] { };          // Default to an empty character array

            SetDefaultRules();

            ResetStructures();
            mFastaFilePath = string.Empty;

            mMemoryUsageLogger = new clsMemoryUsageLogger(string.Empty);
            mProcessMemoryUsageMBAtStart = clsMemoryUsageLogger.GetProcessMemoryUsageMB();

            // ReSharper disable once VbUnreachableCode
            if (REPORT_DETAILED_MEMORY_USAGE)
            {
                // mMemoryUsageMBAtStart = mMemoryUsageLogger.GetFreeMemoryMB()
                Console.WriteLine(MEM_USAGE_PREFIX + mMemoryUsageLogger.GetMemoryUsageHeader());
                Console.WriteLine(MEM_USAGE_PREFIX + mMemoryUsageLogger.GetMemoryUsageSummary());
            }

            mTempFilesToDelete = new List<string>();
        }

        private void InitializeRuleDetails(
            ref udtRuleDefinitionType[] ruleDefinitions,
            ref udtRuleDefinitionExtendedType[] ruleDetails)
        {
            int index;

            if (ruleDefinitions == null || ruleDefinitions.Length == 0)
            {
                ruleDetails = new udtRuleDefinitionExtendedType[0];
            }
            else
            {
                ruleDetails = new udtRuleDefinitionExtendedType[ruleDefinitions.Length];

                for (index = 0; index <= ruleDefinitions.Length - 1; index++)
                {
                    try
                    {
                        ruleDetails[index].RuleDefinition = ruleDefinitions[index];
                        ruleDetails[index].reRule = new Regex(
                            ruleDetails[index].RuleDefinition.MatchRegEx,
                            RegexOptions.Singleline |
                            RegexOptions.Compiled);
                        ruleDetails[index].Valid = true;
                    }
                    catch (Exception ex)
                    {
                        // Ignore the error, but mark .Valid = false
                        ruleDetails[index].Valid = false;
                    }
                }
            }
        }

        private bool LoadExistingProteinHashFile(
            string proteinHashFilePath,
            out clsNestedStringIntList preloadedProteinNamesToKeep)
        {
            // List of protein names to keep
            // Keys are protein names, values are the number of entries written to the fixed fasta file for the given protein name
            preloadedProteinNamesToKeep = null;

            try
            {
                var proteinHashFile = new FileInfo(proteinHashFilePath);

                if (!proteinHashFile.Exists)
                {
                    ShowErrorMessage("Protein hash file not found: " + proteinHashFilePath);
                    return false;
                }

                // Sort the protein has file on the Sequence_Hash column
                // First cache the column names from the header line

                var headerInfo = new Dictionary<string, int>(StringComparer.InvariantCultureIgnoreCase);
                long proteinHashFileLines = 0;
                string cachedHeaderLine = string.Empty;

                ShowMessage("Examining pre-existing protein hash file to count the number of entries: " + Path.GetFileName(proteinHashFilePath));
                var lastStatus = DateTime.UtcNow;
                bool progressDotShown = false;

                using (var hashFileReader = new StreamReader(new FileStream(proteinHashFile.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)))
                {
                    if (!hashFileReader.EndOfStream)
                    {
                        cachedHeaderLine = hashFileReader.ReadLine();

                        var headerNames = cachedHeaderLine.Split('\t');

                        for (int colIndex = 0; colIndex <= headerNames.Count() - 1; colIndex++)
                            headerInfo.Add(headerNames[colIndex], colIndex);

                        proteinHashFileLines = 1;
                    }

                    while (!hashFileReader.EndOfStream)
                    {
                        hashFileReader.ReadLine();
                        proteinHashFileLines += 1;

                        if (proteinHashFileLines % 10000 == 0)
                        {
                            if (DateTime.UtcNow.Subtract(lastStatus).TotalSeconds >= 10)
                            {
                                Console.Write(".");
                                progressDotShown = true;
                                lastStatus = DateTime.UtcNow;
                            }
                        }
                    }
                }

                if (progressDotShown)
                    Console.WriteLine();

                if (headerInfo.Count == 0)
                {
                    ShowErrorMessage("Protein hash file is empty: " + proteinHashFilePath);
                    return false;
                }

                int sequenceHashColumnIndex;
                if (!headerInfo.TryGetValue(SEQUENCE_HASH_COLUMN, out sequenceHashColumnIndex))
                {
                    ShowErrorMessage("Protein hash file is missing the " + SEQUENCE_HASH_COLUMN + " column: " + proteinHashFilePath);
                    return false;
                }

                int proteinNameColumnIndex;
                if (!headerInfo.TryGetValue(PROTEIN_NAME_COLUMN, out proteinNameColumnIndex))
                {
                    ShowErrorMessage("Protein hash file is missing the " + PROTEIN_NAME_COLUMN + " column: " + proteinHashFilePath);
                    return false;
                }

                int sequenceLengthColumnIndex;
                if (!headerInfo.TryGetValue(SEQUENCE_LENGTH_COLUMN, out sequenceLengthColumnIndex))
                {
                    ShowErrorMessage("Protein hash file is missing the " + SEQUENCE_LENGTH_COLUMN + " column: " + proteinHashFilePath);
                    return false;
                }

                string baseHashFileName = Path.GetFileNameWithoutExtension(proteinHashFile.Name);
                string sortedProteinHashSuffix;
                string proteinHashFilenameSuffixNoExtension = Path.GetFileNameWithoutExtension(PROTEIN_HASHES_FILENAME_SUFFIX);

                if (baseHashFileName.EndsWith(proteinHashFilenameSuffixNoExtension))
                {
                    baseHashFileName = baseHashFileName.Substring(0, baseHashFileName.Length - proteinHashFilenameSuffixNoExtension.Length);
                    sortedProteinHashSuffix = proteinHashFilenameSuffixNoExtension;
                }
                else
                {
                    sortedProteinHashSuffix = string.Empty;
                }

                string baseDataFileName = Path.Combine(proteinHashFile.Directory.FullName, baseHashFileName);

                // Note: do not add sortedProteinHashFilePath to mTempFilesToDelete
                string sortedProteinHashFilePath = baseDataFileName + sortedProteinHashSuffix + "_Sorted.tmp";

                var sortedProteinHashFile = new FileInfo(sortedProteinHashFilePath);
                long sortedHashFileLines = 0;
                bool sortRequired = true;

                if (sortedProteinHashFile.Exists)
                {
                    lastStatus = DateTime.UtcNow;
                    progressDotShown = false;
                    ShowMessage("Validating existing sorted protein hash file: " + sortedProteinHashFile.Name);

                    // The sorted file exists; if it has the same number of lines as the sort file, assume that it is complete
                    using (var sortedHashFileReader = new StreamReader(new FileStream(sortedProteinHashFile.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)))
                    {
                        if (!sortedHashFileReader.EndOfStream)
                        {
                            string headerLine = sortedHashFileReader.ReadLine();
                            if (!string.Equals(headerLine, cachedHeaderLine))
                            {
                                sortedHashFileLines = -1;
                            }
                            else
                            {
                                sortedHashFileLines = 1;
                            }
                        }

                        if (sortedHashFileLines > 0)
                        {
                            while (!sortedHashFileReader.EndOfStream)
                            {
                                sortedHashFileReader.ReadLine();
                                sortedHashFileLines += 1;

                                if (sortedHashFileLines % 10000 == 0)
                                {
                                    if (DateTime.UtcNow.Subtract(lastStatus).TotalSeconds >= 10)
                                    {
                                        Console.Write((sortedHashFileLines / (double)proteinHashFileLines * 100.0).ToString("0") + "% ");
                                        progressDotShown = true;
                                        lastStatus = DateTime.UtcNow;
                                    }
                                }
                            }
                        }
                    }

                    if (progressDotShown)
                        Console.WriteLine();

                    if (sortedHashFileLines == proteinHashFileLines)
                    {
                        sortRequired = false;
                    }
                    else
                    {
                        if (sortedHashFileLines < 0)
                        {
                            ShowMessage("Existing sorted hash file has an incorrect header; re-creating it");
                        }
                        else
                        {
                            ShowMessage(string.Format("Existing sorted hash file has fewer lines ({0}) than the original ({1}); re-creating it",
                                                      sortedHashFileLines, proteinHashFileLines));
                        }

                        sortedProteinHashFile.Delete();
                        System.Threading.Thread.Sleep(50);
                    }
                }

                // Create the sorted protein sequence hash file if necessary
                if (sortRequired)
                {
                    Console.WriteLine();
                    ShowMessage("Sorting the existing protein hash file to create " + Path.GetFileName(sortedProteinHashFilePath));
                    bool sortHashSuccess = SortFile(proteinHashFile, sequenceHashColumnIndex, sortedProteinHashFilePath);
                    if (!sortHashSuccess)
                    {
                        return false;
                    }
                }

                ShowMessage("Determining the best spanner length for protein names");

                // Examine the protein names in the sequence hash file to determine the appropriate spanner length for the names
                byte spannerCharLength = clsNestedStringIntList.AutoDetermineSpannerCharLength(proteinHashFile, proteinNameColumnIndex, true);
                const bool RAISE_EXCEPTION_IF_ADDED_DATA_NOT_SORTED = true;

                // List of protein names to keep
                preloadedProteinNamesToKeep = new clsNestedStringIntList(spannerCharLength, RAISE_EXCEPTION_IF_ADDED_DATA_NOT_SORTED);

                lastStatus = DateTime.UtcNow;
                progressDotShown = false;
                Console.WriteLine();
                ShowMessage("Finding the name of the first protein for each protein hash");

                string currentHash = string.Empty;
                string currentHashSeqLength = string.Empty;
                var proteinNamesCurrentHash = new SortedSet<string>();

                int linesRead = 0;
                int proteinNamesUnsortedCount = 0;

                string proteinNamesUnsortedFilePath = baseDataFileName + "_ProteinNamesUnsorted.tmp";
                string proteinNamesToKeepFilePath = baseDataFileName + "_ProteinNamesToKeep.tmp";
                string uniqueProteinSeqsFilePath = baseDataFileName + "_UniqueProteinSeqs.txt";
                string uniqueProteinSeqDuplicateFilePath = baseDataFileName + "_UniqueProteinSeqDuplicates.txt";

                mTempFilesToDelete.Add(proteinNamesUnsortedFilePath);
                mTempFilesToDelete.Add(proteinNamesToKeepFilePath);

                // Write the first protein name for each sequence hash to a new file
                // Also create the _UniqueProteinSeqs.txt and _UniqueProteinSeqDuplicates.txt files

                using (var sortedHashFileReader = new StreamReader(new FileStream(sortedProteinHashFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)))
                using (var proteinNamesUnsortedWriter = new StreamWriter(new FileStream(proteinNamesUnsortedFilePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)))
                using (var proteinNamesToKeepWriter = new StreamWriter(new FileStream(proteinNamesToKeepFilePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)))
                using (var uniqueProteinSeqsWriter = new StreamWriter(new FileStream(uniqueProteinSeqsFilePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)))
                using (var uniqueProteinSeqDuplicateWriter = new StreamWriter(new FileStream(uniqueProteinSeqDuplicateFilePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)))
                {
                    proteinNamesUnsortedWriter.WriteLine(FlattenList(new List<string>() { PROTEIN_NAME_COLUMN, SEQUENCE_HASH_COLUMN }));

                    proteinNamesToKeepWriter.WriteLine(PROTEIN_NAME_COLUMN);

                    var headerColumnsProteinSeqs = new List<string>()
                    {
                        "Sequence_Index",
                        "Protein_Name_First",
                        SEQUENCE_LENGTH_COLUMN,
                        "Sequence_Hash",
                        "Protein_Count",
                        "Duplicate_Proteins"
                    };

                    uniqueProteinSeqsWriter.WriteLine(FlattenList(headerColumnsProteinSeqs));

                    var headerColumnsSeqDups = new List<string>()
                    {
                        "Sequence_Index",
                        "Protein_Name_First",
                        SEQUENCE_LENGTH_COLUMN,
                        "Duplicate_Protein"
                    };

                    uniqueProteinSeqDuplicateWriter.WriteLine(FlattenList(headerColumnsSeqDups));

                    if (!sortedHashFileReader.EndOfStream)
                    {
                        // Read the header line
                        sortedHashFileReader.ReadLine();
                    }

                    int currentSequenceIndex = 0;

                    while (!sortedHashFileReader.EndOfStream)
                    {
                        string dataLine = sortedHashFileReader.ReadLine();
                        linesRead += 1;

                        var dataValues = dataLine.Split('\t');

                        string proteinName = dataValues[proteinNameColumnIndex];
                        string proteinHash = dataValues[sequenceHashColumnIndex];

                        proteinNamesUnsortedWriter.WriteLine(FlattenList(new List<string>() { proteinName, proteinHash }));
                        proteinNamesUnsortedCount += 1;

                        if (string.Equals(currentHash, proteinHash))
                        {
                            // Existing sequence hash
                            if (!proteinNamesCurrentHash.Contains(proteinName))
                            {
                                proteinNamesCurrentHash.Add(proteinName);
                            }
                        }
                        else
                        {
                            // New sequence hash found

                            // First write out the data for the last hash
                            if (!string.IsNullOrEmpty(currentHash))
                            {
                                WriteCachedProteinHashMetadata(
                                    proteinNamesToKeepWriter,
                                    uniqueProteinSeqsWriter,
                                    uniqueProteinSeqDuplicateWriter,
                                    currentSequenceIndex,
                                    currentHash,
                                    currentHashSeqLength,
                                    proteinNamesCurrentHash);
                            }

                            // Update the currentHash values
                            currentSequenceIndex += 1;
                            currentHash = string.Copy(proteinHash);
                            currentHashSeqLength = dataValues[sequenceLengthColumnIndex];

                            proteinNamesCurrentHash.Clear();
                            proteinNamesCurrentHash.Add(proteinName);
                        }

                        if (linesRead % 10000 == 0)
                        {
                            if (DateTime.UtcNow.Subtract(lastStatus).TotalSeconds >= 10)
                            {
                                Console.Write((linesRead / (double)proteinHashFileLines * 100.0).ToString("0") + "% ");
                                progressDotShown = true;
                                lastStatus = DateTime.UtcNow;
                            }
                        }
                    }

                    // Write out the data for the last hash
                    if (!string.IsNullOrEmpty(currentHash))
                    {
                        WriteCachedProteinHashMetadata(
                            proteinNamesToKeepWriter,
                            uniqueProteinSeqsWriter,
                            uniqueProteinSeqDuplicateWriter,
                            currentSequenceIndex,
                            currentHash,
                            currentHashSeqLength,
                            proteinNamesCurrentHash);
                    }
                }

                if (progressDotShown)
                    Console.WriteLine();
                Console.WriteLine();

                // Sort the protein names to keep file
                string sortedProteinNamesToKeepFilePath = baseDataFileName + "_ProteinsToKeepSorted.tmp";
                mTempFilesToDelete.Add(sortedProteinNamesToKeepFilePath);

                ShowMessage("Sorting the protein names to keep file to create " + Path.GetFileName(sortedProteinNamesToKeepFilePath));
                bool sortProteinNamesToKeepSuccess = SortFile(new FileInfo(proteinNamesToKeepFilePath), 0, sortedProteinNamesToKeepFilePath);

                if (!sortProteinNamesToKeepSuccess)
                {
                    return false;
                }

                string lastProteinAdded = string.Empty;

                // Read the sorted protein names to keep file and cache the protein names in memory
                using (var sortedNamesFileReader = new StreamReader(new FileStream(sortedProteinNamesToKeepFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)))
                {
                    // Skip the header line
                    sortedNamesFileReader.ReadLine();

                    while (!sortedNamesFileReader.EndOfStream)
                    {
                        string proteinName = sortedNamesFileReader.ReadLine();

                        if (string.Equals(lastProteinAdded, proteinName))
                        {
                            continue;
                        }

                        // Store the protein name, plus a 0
                        // The stored value will be incremented when this protein name is encountered by the validator
                        preloadedProteinNamesToKeep.Add(proteinName, 0);

                        lastProteinAdded = string.Copy(proteinName);
                    }
                }

                // Confirm that the data is sorted
                preloadedProteinNamesToKeep.Sort();

                ShowMessage("Cached " + preloadedProteinNamesToKeep.Count.ToString("#,##0") + " protein names into memory");
                ShowMessage("The fixed FASTA file will only contain entries for these proteins");

                // Sort the protein names file so that we can check for duplicate protein names
                string sortedProteinNamesFilePath = baseDataFileName + "_ProteinNamesSorted.tmp";
                mTempFilesToDelete.Add(sortedProteinNamesFilePath);

                Console.WriteLine();
                ShowMessage("Sorting the protein names file to create " + Path.GetFileName(sortedProteinNamesFilePath));
                bool sortProteinNamesSuccess = SortFile(new FileInfo(proteinNamesUnsortedFilePath), 0, sortedProteinNamesFilePath);

                if (!sortProteinNamesSuccess)
                {
                    return false;
                }

                string proteinNameSummaryFilePath = baseDataFileName + "_ProteinNameSummary.txt";

                // We can now safely delete some files to free up disk space
                try
                {
                    System.Threading.Thread.Sleep(100);
                    File.Delete(proteinNamesUnsortedFilePath);
                }
                catch (Exception ex)
                {
                    // Ignore errors here
                }

                // Look for duplicate protein names
                // In addition, create a new file with all protein names plus also two new columns: First_Protein_For_Hash and Duplicate_Name

                using (var sortedProteinNamesReader = new StreamReader(new FileStream(sortedProteinNamesFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)))
                using (var proteinNameSummaryWriter = new StreamWriter(new FileStream(proteinNameSummaryFilePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)))
                {
                    // Skip the header line
                    sortedProteinNamesReader.ReadLine();

                    var proteinNameHeaderColumns = new List<string>()
                    {
                        PROTEIN_NAME_COLUMN,
                        SEQUENCE_HASH_COLUMN,
                        "First_Protein_For_Hash",
                        "Duplicate_Name"
                    };

                    proteinNameSummaryWriter.WriteLine(FlattenList(proteinNameHeaderColumns));

                    string lastProtein = string.Empty;
                    bool warningShown = false;
                    int duplicateCount = 0;

                    linesRead = 0;
                    lastStatus = DateTime.UtcNow;
                    progressDotShown = false;

                    while (!sortedProteinNamesReader.EndOfStream)
                    {
                        string dataLine = sortedProteinNamesReader.ReadLine();
                        var dataValues = dataLine.Split('\t');

                        string currentProtein = dataValues[0];
                        string sequenceHash = dataValues[1];

                        bool firstProteinForHash = preloadedProteinNamesToKeep.Contains(currentProtein);
                        bool duplicateProtein = false;

                        if (string.Equals(lastProtein, currentProtein))
                        {
                            duplicateProtein = true;

                            if (!warningShown)
                            {
                                ShowMessage("WARNING: the protein hash file has duplicate protein names: " + Path.GetFileName(proteinHashFilePath));
                                warningShown = true;
                            }

                            duplicateCount += 1;

                            if (duplicateCount < 10)
                            {
                                ShowMessage("  ... duplicate protein name: " + currentProtein);
                            }
                            else if (duplicateCount == 10)
                            {
                                ShowMessage("  ... additional duplicates will not be shown ...");
                            }
                        }

                        lastProtein = string.Copy(currentProtein);

                        var dataToWrite = new List<string>()
                        {
                            currentProtein,
                            sequenceHash,
                            BoolToStringInt(firstProteinForHash),
                            BoolToStringInt(duplicateProtein)
                        };

                        proteinNameSummaryWriter.WriteLine(FlattenList(dataToWrite));

                        if (linesRead % 10000 == 0)
                        {
                            if (DateTime.UtcNow.Subtract(lastStatus).TotalSeconds >= 10)
                            {
                                Console.Write((linesRead / (double)proteinNamesUnsortedCount * 100.0).ToString("0") + "% ");
                                progressDotShown = true;
                                lastStatus = DateTime.UtcNow;
                            }
                        }
                    }

                    if (duplicateCount > 0)
                    {
                        ShowMessage("WARNING: Found " + duplicateCount.ToString("#,##0") + " duplicate protein names in the protein hash file");
                    }
                }

                if (progressDotShown)
                {
                    Console.WriteLine();
                }

                Console.WriteLine();

                return true;
            }
            catch (Exception ex)
            {
                OnErrorEvent("Error in LoadExistingProteinHashFile", ex);
                return false;
            }
        }

        public bool LoadParameterFileSettings(string parameterFilePath)
        {
            var settingsFile = new XmlSettingsFileAccessor();

            var customRulesLoaded = false;
            bool success;

            string characterList;

            try
            {
                if (parameterFilePath == null || parameterFilePath.Length == 0)
                {
                    // No parameter file specified; nothing to load
                    return true;
                }

                if (!File.Exists(parameterFilePath))
                {
                    // See if parameterFilePath points to a file in the same directory as the application
                    parameterFilePath = Path.Combine(
                        Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location),
                        Path.GetFileName(parameterFilePath));

                    if (!File.Exists(parameterFilePath))
                    {
                        SetBaseClassErrorCode(ProcessFilesErrorCodes.ParameterFileNotFound);
                        return false;
                    }
                }

                if (settingsFile.LoadSettings(parameterFilePath))
                {
                    if (!settingsFile.SectionPresent(XML_SECTION_OPTIONS))
                    {
                        OnWarningEvent("The node '<section name=\"" + XML_SECTION_OPTIONS + "\"> was not found in the parameter file: " + parameterFilePath);
                        SetBaseClassErrorCode(ProcessFilesErrorCodes.InvalidParameterFile);
                        return false;
                    }
                    else
                    {
                        // Read customized settings

                        SetOptionSwitch(SwitchOptions.AddMissingLineFeedAtEOF,
                            settingsFile.GetParam(XML_SECTION_OPTIONS, "AddMissingLinefeedAtEOF",
                            GetOptionSwitchValue(SwitchOptions.AddMissingLineFeedAtEOF)));
                        SetOptionSwitch(SwitchOptions.AllowAsteriskInResidues,
                            settingsFile.GetParam(XML_SECTION_OPTIONS, "AllowAsteriskInResidues",
                            GetOptionSwitchValue(SwitchOptions.AllowAsteriskInResidues)));
                        SetOptionSwitch(SwitchOptions.AllowDashInResidues,
                            settingsFile.GetParam(XML_SECTION_OPTIONS, "AllowDashInResidues",
                            GetOptionSwitchValue(SwitchOptions.AllowDashInResidues)));
                        SetOptionSwitch(SwitchOptions.CheckForDuplicateProteinNames,
                            settingsFile.GetParam(XML_SECTION_OPTIONS, "CheckForDuplicateProteinNames",
                            GetOptionSwitchValue(SwitchOptions.CheckForDuplicateProteinNames)));
                        SetOptionSwitch(SwitchOptions.CheckForDuplicateProteinSequences,
                            settingsFile.GetParam(XML_SECTION_OPTIONS, "CheckForDuplicateProteinSequences",
                            GetOptionSwitchValue(SwitchOptions.CheckForDuplicateProteinSequences)));

                        SetOptionSwitch(SwitchOptions.SaveProteinSequenceHashInfoFiles,
                            settingsFile.GetParam(XML_SECTION_OPTIONS, "SaveProteinSequenceHashInfoFiles",
                            GetOptionSwitchValue(SwitchOptions.SaveProteinSequenceHashInfoFiles)));

                        SetOptionSwitch(SwitchOptions.SaveBasicProteinHashInfoFile,
                            settingsFile.GetParam(XML_SECTION_OPTIONS, "SaveBasicProteinHashInfoFile",
                            GetOptionSwitchValue(SwitchOptions.SaveBasicProteinHashInfoFile)));

                        MaximumFileErrorsToTrack = settingsFile.GetParam(XML_SECTION_OPTIONS,
                            "MaximumFileErrorsToTrack", MaximumFileErrorsToTrack);
                        MinimumProteinNameLength = settingsFile.GetParam(XML_SECTION_OPTIONS,
                            "MinimumProteinNameLength", MinimumProteinNameLength);
                        MaximumProteinNameLength = settingsFile.GetParam(XML_SECTION_OPTIONS,
                            "MaximumProteinNameLength", MaximumProteinNameLength);
                        MaximumResiduesPerLine = settingsFile.GetParam(XML_SECTION_OPTIONS,
                            "MaximumResiduesPerLine", MaximumResiduesPerLine);

                        SetOptionSwitch(SwitchOptions.WarnBlankLinesBetweenProteins,
                            settingsFile.GetParam(XML_SECTION_OPTIONS, "WarnBlankLinesBetweenProteins",
                            GetOptionSwitchValue(SwitchOptions.WarnBlankLinesBetweenProteins)));
                        SetOptionSwitch(SwitchOptions.WarnLineStartsWithSpace,
                            settingsFile.GetParam(XML_SECTION_OPTIONS, "WarnLineStartsWithSpace",
                            GetOptionSwitchValue(SwitchOptions.WarnLineStartsWithSpace)));

                        SetOptionSwitch(SwitchOptions.OutputToStatsFile,
                            settingsFile.GetParam(XML_SECTION_OPTIONS, "OutputToStatsFile",
                            GetOptionSwitchValue(SwitchOptions.OutputToStatsFile)));

                        SetOptionSwitch(SwitchOptions.NormalizeFileLineEndCharacters,
                            settingsFile.GetParam(XML_SECTION_OPTIONS, "NormalizeFileLineEndCharacters",
                            GetOptionSwitchValue(SwitchOptions.NormalizeFileLineEndCharacters)));

                        if (!settingsFile.SectionPresent(XML_SECTION_FIXED_FASTA_FILE_OPTIONS))
                        {
                            // "ValidateFastaFixedFASTAFileOptions" section not present
                            // Only read the settings for GenerateFixedFASTAFile and SplitOutMultipleRefsInProteinName

                            SetOptionSwitch(SwitchOptions.GenerateFixedFASTAFile,
                                settingsFile.GetParam(XML_SECTION_OPTIONS, "GenerateFixedFASTAFile",
                                GetOptionSwitchValue(SwitchOptions.GenerateFixedFASTAFile)));

                            SetOptionSwitch(SwitchOptions.SplitOutMultipleRefsInProteinName,
                                settingsFile.GetParam(XML_SECTION_OPTIONS, "SplitOutMultipleRefsInProteinName",
                                GetOptionSwitchValue(SwitchOptions.SplitOutMultipleRefsInProteinName)));
                        }
                        else
                        {
                            // "ValidateFastaFixedFASTAFileOptions" section is present

                            SetOptionSwitch(SwitchOptions.GenerateFixedFASTAFile,
                                settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "GenerateFixedFASTAFile",
                                GetOptionSwitchValue(SwitchOptions.GenerateFixedFASTAFile)));

                            SetOptionSwitch(SwitchOptions.SplitOutMultipleRefsInProteinName,
                                settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "SplitOutMultipleRefsInProteinName",
                                GetOptionSwitchValue(SwitchOptions.SplitOutMultipleRefsInProteinName)));

                            SetOptionSwitch(SwitchOptions.FixedFastaRenameDuplicateNameProteins,
                                settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "RenameDuplicateNameProteins",
                                GetOptionSwitchValue(SwitchOptions.FixedFastaRenameDuplicateNameProteins)));

                            SetOptionSwitch(SwitchOptions.FixedFastaKeepDuplicateNamedProteins,
                                settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "KeepDuplicateNamedProteins",
                                GetOptionSwitchValue(SwitchOptions.FixedFastaKeepDuplicateNamedProteins)));

                            SetOptionSwitch(SwitchOptions.FixedFastaConsolidateDuplicateProteinSeqs,
                                settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ConsolidateDuplicateProteinSeqs",
                                GetOptionSwitchValue(SwitchOptions.FixedFastaConsolidateDuplicateProteinSeqs)));

                            SetOptionSwitch(SwitchOptions.FixedFastaConsolidateDupsIgnoreILDiff,
                                settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ConsolidateDupsIgnoreILDiff",
                                GetOptionSwitchValue(SwitchOptions.FixedFastaConsolidateDupsIgnoreILDiff)));

                            SetOptionSwitch(SwitchOptions.FixedFastaTruncateLongProteinNames,
                                settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "TruncateLongProteinNames",
                                GetOptionSwitchValue(SwitchOptions.FixedFastaTruncateLongProteinNames)));

                            SetOptionSwitch(SwitchOptions.FixedFastaSplitOutMultipleRefsForKnownAccession,
                                settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "SplitOutMultipleRefsForKnownAccession",
                                GetOptionSwitchValue(SwitchOptions.FixedFastaSplitOutMultipleRefsForKnownAccession)));

                            SetOptionSwitch(SwitchOptions.FixedFastaWrapLongResidueLines,
                                settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "WrapLongResidueLines",
                                GetOptionSwitchValue(SwitchOptions.FixedFastaWrapLongResidueLines)));

                            SetOptionSwitch(SwitchOptions.FixedFastaRemoveInvalidResidues,
                                settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "RemoveInvalidResidues",
                                GetOptionSwitchValue(SwitchOptions.FixedFastaRemoveInvalidResidues)));

                            // Look for the special character lists
                            // If defined, then update the default values
                            characterList = settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "LongProteinNameSplitChars", string.Empty);
                            if (characterList != null && characterList.Length > 0)
                            {
                                // Update mFixedFastaOptions.LongProteinNameSplitChars with characterList
                                LongProteinNameSplitChars = characterList;
                            }

                            characterList = settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ProteinNameInvalidCharsToRemove", string.Empty);
                            if (characterList != null && characterList.Length > 0)
                            {
                                // Update mFixedFastaOptions.ProteinNameInvalidCharsToRemove with characterList
                                ProteinNameInvalidCharsToRemove = characterList;
                            }

                            characterList = settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ProteinNameFirstRefSepChars", string.Empty);
                            if (characterList != null && characterList.Length > 0)
                            {
                                // Update mProteinNameFirstRefSepChars
                                ProteinNameFirstRefSepChars = characterList;
                            }

                            characterList = settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ProteinNameSubsequentRefSepChars", string.Empty);
                            if (characterList != null && characterList.Length > 0)
                            {
                                // Update mProteinNameSubsequentRefSepChars
                                ProteinNameSubsequentRefSepChars = characterList;
                            }
                        }

                        // Read the custom rules
                        // If all of the sections are missing, then use the default rules
                        customRulesLoaded = false;

                        success = ReadRulesFromParameterFile(settingsFile, XML_SECTION_FASTA_HEADER_LINE_RULES, ref mHeaderLineRules);
                        customRulesLoaded = customRulesLoaded || success;

                        success = ReadRulesFromParameterFile(settingsFile, XML_SECTION_FASTA_PROTEIN_NAME_RULES, ref mProteinNameRules);
                        customRulesLoaded = customRulesLoaded || success;

                        success = ReadRulesFromParameterFile(settingsFile, XML_SECTION_FASTA_PROTEIN_DESCRIPTION_RULES, ref mProteinDescriptionRules);
                        customRulesLoaded = customRulesLoaded || success;

                        success = ReadRulesFromParameterFile(settingsFile, XML_SECTION_FASTA_PROTEIN_SEQUENCE_RULES, ref mProteinSequenceRules);
                        customRulesLoaded = customRulesLoaded || success;
                    }
                }
            }
            catch (Exception ex)
            {
                OnErrorEvent("Error in LoadParameterFileSettings", ex);
                return false;
            }

            if (!customRulesLoaded)
            {
                SetDefaultRules();
            }

            return true;
        }

        public string LookupMessageDescription(int errorMessageCode)
        {
            return LookupMessageDescription(errorMessageCode, null);
        }

        public string LookupMessageDescription(int errorMessageCode, string extraInfo)
        {
            string message;
            bool matchFound;

            switch (errorMessageCode)
            {
                // Error messages
                case (int)eMessageCodeConstants.ProteinNameIsTooLong:
                    message = "Protein name is longer than the maximum allowed length of " + mMaximumProteinNameLength.ToString() + " characters";
                    break;
                //case (int)eMessageCodeConstants.ProteinNameContainsInvalidCharacters:
                //    message = "Protein name contains invalid characters";
                //    if (!specifiedInvalidProteinNameChars)
                //    {
                //        message += " (should contain letters, numbers, period, dash, underscore, colon, comma, or vertical bar)";
                //        specifiedInvalidProteinNameChars = true;
                //    }
                //    break;
                case (int)eMessageCodeConstants.LineStartsWithSpace:
                    message = "Found a line starting with a space";
                    break;
                //case (int)eMessageCodeConstants.RightArrowFollowedBySpace:
                //    message = "Space found directly after the > symbol";
                //    break;
                //case (int)eMessageCodeConstants.RightArrowFollowedByTab:
                //    message = "Tab character found directly after the > symbol";
                //    break;
                //case (int)eMessageCodeConstants.RightArrowButNoProteinName:
                //    message = "Line starts with > but does not contain a protein name";
                //    break;
                case (int)eMessageCodeConstants.BlankLineBetweenProteinNameAndResidues:
                    message = "A blank line was found between the protein name and its residues";
                    break;
                case (int)eMessageCodeConstants.BlankLineInMiddleOfResidues:
                    message = "A blank line was found in the middle of the residue block for the protein";
                    break;
                case (int)eMessageCodeConstants.ResiduesFoundWithoutProteinHeader:
                    message = "Residues were found, but a protein header didn't precede them";
                    break;
                //case (int)eMessageCodeConstants.ResiduesWithAsterisk:
                //    message = "An asterisk was found in the residues";
                //    break;
                //case (int)eMessageCodeConstants.ResiduesWithSpace:
                //    message = "A space was found in the residues";
                //    break;
                //case (int)eMessageCodeConstants.ResiduesWithTab:
                //    message = "A tab character was found in the residues";
                //    break;
                //case (int)eMessageCodeConstants.ResiduesWithInvalidCharacters:
                //    message = "Invalid residues found";
                //    if (!specifiedResidueChars)
                //    {
                //        if (mAllowAsteriskInResidues)
                //        {
                //            message += " (should be any capital letter except J, plus *)";
                //        }
                //        else
                //        {
                //            message += " (should be any capital letter except J)";
                //        }
                //        specifiedResidueChars = true;
                //    }
                //    break;
                case (int)eMessageCodeConstants.ProteinEntriesNotFound:
                    message = "File does not contain any protein entries";
                    break;
                case (int)eMessageCodeConstants.FinalProteinEntryMissingResidues:
                    message = "The last entry in the file is a protein header line, but there is no protein sequence line after it";
                    break;
                case (int)eMessageCodeConstants.FileDoesNotEndWithLinefeed:
                    message = "File does not end in a blank line; this is a problem for Sequest";
                    break;
                case (int)eMessageCodeConstants.DuplicateProteinName:
                    message = "Duplicate protein name found";
                    break;

                // Warning messages
                case (int)eMessageCodeConstants.ProteinNameIsTooShort:
                    message = "Protein name is shorter than the minimum suggested length of " + mMinimumProteinNameLength.ToString() + " characters";
                    break;
                //case (int)eMessageCodeConstants.ProteinNameContainsComma:
                //    message = "Protein name contains a comma";
                //    break;
                //case (int)eMessageCodeConstants.ProteinNameContainsVerticalBars:
                //    message = "Protein name contains two or more vertical bars";
                //    break;
                //case (int)eMessageCodeConstants.ProteinNameContainsWarningCharacters:
                //    message = "Protein name contains undesirable characters";
                //    break;
                //case (int)eMessageCodeConstants.ProteinNameWithoutDescription:
                //    message = "Line contains a protein name, but not a description";
                //    break;
                case (int)eMessageCodeConstants.BlankLineBeforeProteinName:
                    message = "Blank line found before the protein name; this is acceptable, but not preferred";
                    break;
                // case (int)eMessageCodeConstants.ProteinNameAndDescriptionSeparatedByTab:
                //    message = "Protein name is separated from the protein description by a tab";
                //    break;
                case (int)eMessageCodeConstants.ResiduesLineTooLong:
                    message = "Residues line is longer than the suggested maximum length of " + mMaximumResiduesPerLine.ToString() + " characters";
                    break;
                //case (int)eMessageCodeConstants.ProteinDescriptionWithTab:
                //    message = "Protein description contains a tab character";
                //    break;
                //case (int)eMessageCodeConstants.ProteinDescriptionWithQuotationMark:
                //    message = "Protein description contains a quotation mark";
                //    break;
                //case (int)eMessageCodeConstants.ProteinDescriptionWithEscapedSlash:
                //    message = "Protein description contains escaped slash: \/";
                //    break;
                //case (int)eMessageCodeConstants.ProteinDescriptionWithUndesirableCharacter:
                //    message = "Protein description contains undesirable characters";
                //    break;
                //case (int)eMessageCodeConstants.ResiduesLineTooLong:
                //    message = "Residues line is longer than the suggested maximum length of " + mMaximumResiduesPerLine.ToString + " characters";
                //    break;
                //case (int)eMessageCodeConstants.ResiduesLineContainsU:
                //    message = "Residues line contains U (selenocysteine); this residue is unsupported by Sequest";
                //    break;

                case (int)eMessageCodeConstants.DuplicateProteinSequence:
                    message = "Duplicate protein sequences found";
                    break;
                case (int)eMessageCodeConstants.RenamedProtein:
                    message = "Renamed protein because duplicate name";
                    break;
                case (int)eMessageCodeConstants.ProteinRemovedSinceDuplicateSequence:
                    message = "Removed protein since duplicate sequence";
                    break;
                case (int)eMessageCodeConstants.DuplicateProteinNameRetained:
                    message = "Duplicate protein retained in fixed file";
                    break;
                case (int)eMessageCodeConstants.UnspecifiedError:
                    message = "Unspecified error";
                    break;
                default:
                    message = "Unspecified error";

                    // Search the custom rules for the given code
                    matchFound = SearchRulesForID(mHeaderLineRules, errorMessageCode, out message);

                    if (!matchFound)
                    {
                        matchFound = SearchRulesForID(mProteinNameRules, errorMessageCode, out message);
                    }

                    if (!matchFound)
                    {
                        matchFound = SearchRulesForID(mProteinDescriptionRules, errorMessageCode, out message);
                    }

                    if (!matchFound)
                    {
                        SearchRulesForID(mProteinSequenceRules, errorMessageCode, out message);
                    }

                    break;
            }

            if (!string.IsNullOrWhiteSpace(extraInfo))
            {
                message += " (" + extraInfo + ")";
            }

            return message;
        }

        private string LookupMessageType(eMsgTypeConstants EntryType)
        {
            switch (EntryType)
            {
                case eMsgTypeConstants.ErrorMsg:
                    return "Error";
                case eMsgTypeConstants.WarningMsg:
                    return "Warning";
                default:
                    return "Status";
            }
        }

        /// <summary>
        /// Validate a single fasta file
        /// </summary>
        /// <returns>True if success; false if a fatal error</returns>
        /// <remarks>
        /// Note that .ProcessFile returns True if a file is successfully processed (even if errors are found)
        /// Used by clsCustomValidateFastaFiles
        /// </remarks>
        protected bool SimpleProcessFile(string inputFilePath)
        {
            return ProcessFile(inputFilePath, null, null, false);
        }

        /// <summary>
        /// Main processing function
        /// </summary>
        /// <param name="inputFilePath"></param>
        /// <param name="outputFolderPath"></param>
        /// <param name="parameterFilePath"></param>
        /// <param name="resetErrorCode"></param>
        /// <returns>True if success, False if failure</returns>
        public override bool ProcessFile(
            string inputFilePath,
            string outputFolderPath,
            string parameterFilePath,
            bool resetErrorCode)
        {
            FileInfo ioFile;
            StreamWriter swStatsOutFile;

            string inputFilePathFull;
            string statusMessage;

            if (resetErrorCode)
            {
                SetLocalErrorCode(eValidateFastaFileErrorCodes.NoError);
            }

            if (!LoadParameterFileSettings(parameterFilePath))
            {
                statusMessage = "Parameter file load error: " + parameterFilePath;
                OnWarningEvent(statusMessage);
                if (ErrorCode == ProcessFilesErrorCodes.NoError)
                {
                    SetBaseClassErrorCode(ProcessFilesErrorCodes.InvalidParameterFile);
                }

                return false;
            }

            try
            {
                if (inputFilePath == null || inputFilePath.Length == 0)
                {
                    ShowWarning("Input file name is empty");
                    SetBaseClassErrorCode(ProcessFilesErrorCodes.InvalidInputFilePath);
                    return false;
                }
                else
                {
                    Console.WriteLine();
                    ShowMessage("Parsing " + Path.GetFileName(inputFilePath));

                    if (!CleanupFilePaths(ref inputFilePath, ref outputFolderPath))
                    {
                        SetBaseClassErrorCode(ProcessFilesErrorCodes.FilePathError);
                        return false;
                    }
                    else
                    {
                        // List of protein names to keep
                        // Keys are protein names, values are the number of entries written to the fixed fasta file for the given protein name
                        clsNestedStringIntList preloadedProteinNamesToKeep = null;

                        if (!string.IsNullOrEmpty(ExistingProteinHashFile))
                        {
                            bool loadSuccess = LoadExistingProteinHashFile(ExistingProteinHashFile, out preloadedProteinNamesToKeep);
                            if (!loadSuccess)
                            {
                                return false;
                            }
                        }

                        try
                        {
                            // Obtain the full path to the input file
                            ioFile = new FileInfo(inputFilePath);
                            inputFilePathFull = ioFile.FullName;

                            bool success = AnalyzeFastaFile(inputFilePathFull, preloadedProteinNamesToKeep);

                            if (success)
                            {
                                ReportResults(outputFolderPath, mOutputToStatsFile);
                                DeleteTempFiles();
                                return true;
                            }
                            else
                            {
                                if (mOutputToStatsFile)
                                {
                                    mStatsFilePath = ConstructStatsFilePath(outputFolderPath);
                                    swStatsOutFile = new StreamWriter(mStatsFilePath, true);
                                    swStatsOutFile.WriteLine(GetTimeStamp() + "\t" +
                                        "Error parsing " +
                                        Path.GetFileName(inputFilePath) + ": " + GetErrorMessage());
                                    swStatsOutFile.Close();
                                }
                                else
                                {
                                    ShowMessage("Error parsing " +
                                        Path.GetFileName(inputFilePath) +
                                        ": " + GetErrorMessage());
                                }

                                return false;
                            }
                        }
                        catch (Exception ex)
                        {
                            OnErrorEvent("Error calling AnalyzeFastaFile", ex);
                            return false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                OnErrorEvent("Error in ProcessFile", ex);
                return false;
            }
        }

        private readonly char[] extraCharsToTrim = {'|', ' '};

        private void PrependExtraTextToProteinDescription(string extraProteinNameText, ref string proteinDescription)
        {
            if (extraProteinNameText != null && extraProteinNameText.Length > 0)
            {
                // If extraProteinNameText ends in a vertical bar and/or space, them remove them
                extraProteinNameText = extraProteinNameText.TrimEnd(extraCharsToTrim);

                if (proteinDescription != null && proteinDescription.Length > 0)
                {
                    if (proteinDescription[0] == ' ' || proteinDescription[0] == '|')
                    {
                        proteinDescription = extraProteinNameText + proteinDescription;
                    }
                    else
                    {
                        proteinDescription = extraProteinNameText + " " + proteinDescription;
                    }
                }
                else
                {
                    proteinDescription = string.Copy(extraProteinNameText);
                }
            }
        }

        private void ProcessResiduesForPreviousProtein(
            string proteinName,
            StringBuilder sbCurrentResidues,
            clsNestedStringDictionary<int> proteinSequenceHashes,
            ref int proteinSequenceHashCount,
            ref clsProteinHashInfo[] proteinSeqHashInfo,
            bool consolidateDupsIgnoreILDiff,
            TextWriter fixedFastaWriter,
            int currentValidResidueLineLengthMax,
            TextWriter sequenceHashWriter)
        {
            int wrapLength;

            int index;
            int length;

            // Check for and remove any asterisks at the end of the residues
            while (sbCurrentResidues.Length > 0 && sbCurrentResidues[sbCurrentResidues.Length - 1] == '*')
                sbCurrentResidues.Remove(sbCurrentResidues.Length - 1, 1);

            if (sbCurrentResidues.Length > 0)
            {
                // Remove any spaces from the residues

                if (mCheckForDuplicateProteinSequences || mSaveBasicProteinHashInfoFile)
                {
                    // Process the previous protein entry to store a hash of the protein sequence
                    ProcessSequenceHashInfo(
                        proteinName, sbCurrentResidues,
                        proteinSequenceHashes,
                        ref proteinSequenceHashCount, ref proteinSeqHashInfo,
                        consolidateDupsIgnoreILDiff, sequenceHashWriter);
                }

                if (mGenerateFixedFastaFile && mFixedFastaOptions.WrapLongResidueLines)
                {
                    // Write out the residues
                    // Wrap the lines at currentValidResidueLineLengthMax characters (but do not allow to be longer than mMaximumResiduesPerLine residues)

                    wrapLength = currentValidResidueLineLengthMax;
                    if (wrapLength <= 0 || wrapLength > mMaximumResiduesPerLine)
                    {
                        wrapLength = mMaximumResiduesPerLine;
                    }

                    if (wrapLength < 10)
                    {
                        // Do not allow wrapLength to be less than 10
                        wrapLength = 10;
                    }

                    index = 0;
                    int proteinResidueCount = sbCurrentResidues.Length;
                    while (index < sbCurrentResidues.Length)
                    {
                        length = Math.Min(wrapLength, proteinResidueCount - index);
                        fixedFastaWriter.WriteLine(sbCurrentResidues.ToString(index, length));
                        index += wrapLength;
                    }
                }

                sbCurrentResidues.Length = 0;
            }
        }

        private void ProcessSequenceHashInfo(
            string proteinName,
            StringBuilder sbCurrentResidues,
            clsNestedStringDictionary<int> proteinSequenceHashes,
            ref int proteinSequenceHashCount,
            ref clsProteinHashInfo[] proteinSeqHashInfo,
            bool consolidateDupsIgnoreILDiff,
            TextWriter sequenceHashWriter)
        {
            string computedHash;

            try
            {
                if (sbCurrentResidues.Length > 0)
                {
                    // Compute the hash value for sbCurrentResidues
                    computedHash = ComputeProteinHash(sbCurrentResidues, consolidateDupsIgnoreILDiff);

                    if (sequenceHashWriter != null)
                    {
                        var dataValues = new List<string>()
                        {
                            mProteinCount.ToString(),
                            proteinName,
                            sbCurrentResidues.Length.ToString(),
                            computedHash
                        };

                        sequenceHashWriter.WriteLine(FlattenList(dataValues));
                    }

                    if (mCheckForDuplicateProteinSequences && proteinSequenceHashes != null)
                    {
                        // See if proteinSequenceHashes contains hash
                        int seqHashLookupPointer;
                        if (proteinSequenceHashes.TryGetValue(computedHash, out seqHashLookupPointer))
                        {

                            // Value exists; update the entry in proteinSeqHashInfo
                            CachedSequenceHashInfoUpdate(proteinSeqHashInfo[seqHashLookupPointer], proteinName);
                        }
                        else
                        {
                            // Value not yet present; add it
                            CachedSequenceHashInfoUpdateAppend(
                                ref proteinSequenceHashCount, ref proteinSeqHashInfo,
                                computedHash, sbCurrentResidues, proteinName);

                            proteinSequenceHashes.Add(computedHash, proteinSequenceHashCount);
                            proteinSequenceHashCount += 1;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Error caught; pass it up to the calling function
                ShowMessage(ex.Message);
                throw;
            }
        }

        private void CachedSequenceHashInfoUpdate(clsProteinHashInfo proteinSeqHashInfo, string proteinName)
        {
            if ((proteinSeqHashInfo.ProteinNameFirst ?? "") == (proteinName ?? ""))
            {
                proteinSeqHashInfo.DuplicateProteinNameCount += 1;
            }
            else
            {
                proteinSeqHashInfo.AddAdditionalProtein(proteinName);
            }
        }

        private void CachedSequenceHashInfoUpdateAppend(
            ref int proteinSequenceHashCount,
            ref clsProteinHashInfo[] proteinSeqHashInfo,
            string computedHash,
            StringBuilder sbCurrentResidues,
            string proteinName)
        {
            if (proteinSequenceHashCount >= proteinSeqHashInfo.Length)
            {
                // Need to reserve more space in proteinSeqHashInfo
                if (proteinSeqHashInfo.Length < 1000000)
                {
                    var oldProteinSeqHashInfo = proteinSeqHashInfo;
                    proteinSeqHashInfo = new clsProteinHashInfo[(proteinSeqHashInfo.Length * 2)];
                    Array.Copy(oldProteinSeqHashInfo, proteinSeqHashInfo, Math.Min(proteinSeqHashInfo.Length * 2, oldProteinSeqHashInfo.Length));
                }
                else
                {
                    var oldProteinSeqHashInfo1 = proteinSeqHashInfo;
                    proteinSeqHashInfo = new clsProteinHashInfo[((int)Math.Round(proteinSeqHashInfo.Length * 1.2))];
                    Array.Copy(oldProteinSeqHashInfo1, proteinSeqHashInfo, Math.Min((int)Math.Round(proteinSeqHashInfo.Length * 1.2), oldProteinSeqHashInfo1.Length));
                }
            }

            var newProteinHashInfo = new clsProteinHashInfo(computedHash, sbCurrentResidues, proteinName);
            proteinSeqHashInfo[proteinSequenceHashCount] = newProteinHashInfo;
        }

        private bool ReadRulesFromParameterFile(
            XmlSettingsFileAccessor settingsFile,
            string sectionName,
            ref udtRuleDefinitionType[] rules)
        {
            // Returns True if the section named sectionName is present and if it contains an item with keyName = "RuleCount"
            // Note: even if RuleCount = 0, this function will return True

            bool success = false;
            int ruleCount;
            int ruleNumber;

            string ruleBase;

            udtRuleDefinitionType newRule;

            ruleCount = settingsFile.GetParam(sectionName, XML_OPTION_ENTRY_RULE_COUNT, -1);

            if (ruleCount >= 0)
            {
                ClearRulesDataStructure(ref rules);

                for (ruleNumber = 1; ruleNumber <= ruleCount; ruleNumber++)
                {
                    ruleBase = "Rule" + ruleNumber.ToString();

                    newRule.MatchRegEx = settingsFile.GetParam(sectionName, ruleBase + "MatchRegEx", string.Empty);

                    if (newRule.MatchRegEx.Length > 0)
                    {
                        // Only read the rule settings if MatchRegEx contains 1 or more characters

                        newRule.MatchIndicatesProblem = settingsFile.GetParam(sectionName, ruleBase + "MatchIndicatesProblem", true);
                        newRule.MessageWhenProblem = settingsFile.GetParam(sectionName, ruleBase + "MessageWhenProblem", "Error found with RegEx " + newRule.MatchRegEx);
                        newRule.Severity = settingsFile.GetParam(sectionName, ruleBase + "Severity", (short)3);
                        newRule.DisplayMatchAsExtraInfo = settingsFile.GetParam(sectionName, ruleBase + "DisplayMatchAsExtraInfo", false);

                        SetRule(ref rules, newRule.MatchRegEx, newRule.MatchIndicatesProblem, newRule.MessageWhenProblem, newRule.Severity, newRule.DisplayMatchAsExtraInfo);
                    }
                }

                success = true;
            }

            return success;
        }

        private void RecordFastaFileError(
            int lineNumber,
            int charIndex,
            string proteinName,
            int errorMessageCode)
        {
            RecordFastaFileError(lineNumber, charIndex, proteinName,
                errorMessageCode, string.Empty, string.Empty);
        }

        private void RecordFastaFileError(
            int lineNumber,
            int charIndex,
            string proteinName,
            int errorMessageCode,
            string extraInfo,
            string context)
        {
            RecordFastaFileProblemWork(
                ref mFileErrorStats,
                ref mFileErrorCount,
                ref mFileErrors,
                lineNumber,
                charIndex,
                proteinName,
                errorMessageCode,
                extraInfo,
                context);
        }

        private void RecordFastaFileWarning(
            int lineNumber,
            int charIndex,
            string proteinName,
            int warningMessageCode)
        {
            RecordFastaFileWarning(
                lineNumber,
                charIndex,
                proteinName,
                warningMessageCode,
                string.Empty,
                string.Empty);
        }

        private void RecordFastaFileWarning(
            int lineNumber,
            int charIndex,
            string proteinName,
            int warningMessageCode,
            string extraInfo, string context)
        {
            RecordFastaFileProblemWork(ref mFileWarningStats, ref mFileWarningCount,
                ref mFileWarnings, lineNumber, charIndex, proteinName,
                warningMessageCode, extraInfo, context);
        }

        private void RecordFastaFileProblemWork(
            ref udtItemSummaryIndexedType itemSummaryIndexed,
            ref int itemCountSpecified,
            ref udtMsgInfoType[] items,
            int lineNumber,
            int charIndex,
            string proteinName,
            int messageCode,
            string extraInfo,
            string context)
        {
            // Note that charIndex is the index in the source string at which the error occurred
            // When storing in .ColNumber, we add 1 to charIndex

            // Lookup the index of the entry with messageCode in itemSummaryIndexed.ErrorStats
            // Add it if not present

            try
            {
                int itemIndex;
                if (!itemSummaryIndexed.MessageCodeToArrayIndex.TryGetValue(messageCode, out itemIndex))
                {
                    if (itemSummaryIndexed.ErrorStats.Length <= 0)
                    {
                        itemSummaryIndexed.ErrorStats = new udtErrorStatsType[2];
                    }
                    else if (itemSummaryIndexed.ErrorStatsCount == itemSummaryIndexed.ErrorStats.Length)
                    {
                        var oldErrorStats = itemSummaryIndexed.ErrorStats;
                        itemSummaryIndexed.ErrorStats = new udtErrorStatsType[(itemSummaryIndexed.ErrorStats.Length * 2)];
                        Array.Copy(oldErrorStats, itemSummaryIndexed.ErrorStats, Math.Min(itemSummaryIndexed.ErrorStats.Length * 2, oldErrorStats.Length));
                    }

                    itemIndex = itemSummaryIndexed.ErrorStatsCount;
                    itemSummaryIndexed.ErrorStats[itemIndex].MessageCode = messageCode;
                    itemSummaryIndexed.MessageCodeToArrayIndex.Add(messageCode, itemIndex);
                    itemSummaryIndexed.ErrorStatsCount += 1;
                }

                var errorStats = itemSummaryIndexed.ErrorStats;
                if (errorStats[itemIndex].CountSpecified >= mMaximumFileErrorsToTrack)
                {
                    errorStats[itemIndex].CountUnspecified += 1;
                }
                else
                {
                    if (items.Length <= 0)
                    {
                        // Initially reserve space for 10 errors
                        items = new udtMsgInfoType[11];
                    }
                    else if (itemCountSpecified >= items.Length)
                    {
                        // Double the amount of space reserved for errors
                        var oldItems = items;
                        items = new udtMsgInfoType[(items.Length * 2)];
                        Array.Copy(oldItems, items, Math.Min(items.Length * 2, oldItems.Length));
                    }

                    items[itemCountSpecified].LineNumber = lineNumber;
                    items[itemCountSpecified].ColNumber = charIndex + 1;
                    if (proteinName == null)
                    {
                        items[itemCountSpecified].ProteinName = string.Empty;
                    }
                    else
                    {
                        items[itemCountSpecified].ProteinName = proteinName;
                    }

                    items[itemCountSpecified].MessageCode = messageCode;
                    if (extraInfo == null)
                    {
                        items[itemCountSpecified].ExtraInfo = string.Empty;
                    }
                    else
                    {
                        items[itemCountSpecified].ExtraInfo = extraInfo;
                    }

                    if (extraInfo == null)
                    {
                        items[itemCountSpecified].Context = string.Empty;
                    }
                    else
                    {
                        items[itemCountSpecified].Context = context;
                    }

                    itemCountSpecified += 1;

                    errorStats[itemIndex].CountSpecified += 1;
                }
            }
            catch (Exception ex)
            {
                // Ignore any errors that occur, but output the error to the console
                OnWarningEvent("Error in RecordFastaFileProblemWork: " + ex.Message);
            }
        }

        private void ReplaceXMLCodesWithText(string parameterFilePath)
        {
            string outputFilePath;
            string timeStamp;
            string lineIn;

            try
            {
                // Define the output file path
                timeStamp = GetTimeStamp().Replace(" ", "_").Replace(":", "_").Replace("/", "_");

                outputFilePath = parameterFilePath + "_" + timeStamp + ".fixed";

                // Open the input file and output file
                using (var srInFile = new StreamReader(parameterFilePath))
                using (var swOutFile = new StreamWriter(new FileStream(outputFilePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)))
                {
                    // Parse each line in the file
                    while (!srInFile.EndOfStream)
                    {
                        lineIn = srInFile.ReadLine();

                        if (lineIn != null)
                        {
                            lineIn = lineIn.Replace("&gt;", ">").Replace("&lt;", "<");
                            swOutFile.WriteLine(lineIn);
                        }
                    }

                    // Close the input and output files
                }

                // Wait 100 msec
                System.Threading.Thread.Sleep(100);

                // Delete the input file
                File.Delete(parameterFilePath);

                // Wait 250 msec
                System.Threading.Thread.Sleep(250);

                // Rename the output file to the input file
                File.Move(outputFilePath, parameterFilePath);
            }
            catch (Exception ex)
            {
                OnErrorEvent("Error in ReplaceXMLCodesWithText", ex);
            }
        }

        private void ReportMemoryUsage()
        {

            // ReSharper disable once VbUnreachableCode
            if (REPORT_DETAILED_MEMORY_USAGE)
            {
                Console.WriteLine(MEM_USAGE_PREFIX + mMemoryUsageLogger.GetMemoryUsageSummary());
            }
            else
            {
                Console.WriteLine(MEM_USAGE_PREFIX + GetProcessMemoryUsageWithTimestamp());
            }
        }

        private void ReportMemoryUsage(
            clsNestedStringIntList preloadedProteinNamesToKeep,
            clsNestedStringDictionary<int> proteinSequenceHashes,
            ICollection<string> proteinNames,
            IEnumerable<clsProteinHashInfo> proteinSeqHashInfo)
        {
            Console.WriteLine();
            ReportMemoryUsage();

            if (preloadedProteinNamesToKeep != null && preloadedProteinNamesToKeep.Count > 0)
            {
                Console.WriteLine(" PreloadedProteinNamesToKeep: {0,12:#,##0} records", preloadedProteinNamesToKeep.Count);
            }

            if (proteinSequenceHashes.Count > 0)
            {
                Console.WriteLine(" ProteinSequenceHashes:  {0,12:#,##0} records", proteinSequenceHashes.Count);
                Console.WriteLine("   {0}", proteinSequenceHashes.GetSizeSummary());
            }

            Console.WriteLine(" ProteinNames:           {0,12:#,##0} records", proteinNames.Count);
            Console.WriteLine(" ProteinSeqHashInfo:       {0,12:#,##0} records", proteinSeqHashInfo.Count());
        }

        private void ReportMemoryUsage(
            clsNestedStringDictionary<int> proteinNameFirst,
            clsNestedStringDictionary<int> proteinsWritten,
            clsNestedStringDictionary<string> duplicateProteinList)
        {
            Console.WriteLine();
            ReportMemoryUsage();
            Console.WriteLine(" ProteinNameFirst:      {0,12:#,##0} records", proteinNameFirst.Count);
            Console.WriteLine("   {0}", proteinNameFirst.GetSizeSummary());
            Console.WriteLine(" ProteinsWritten:       {0,12:#,##0} records", proteinsWritten.Count);
            Console.WriteLine("   {0}", proteinsWritten.GetSizeSummary());
            Console.WriteLine(" DuplicateProteinList:  {0,12:#,##0} records", duplicateProteinList.Count);
            Console.WriteLine("   {0}", duplicateProteinList.GetSizeSummary());
        }

        private void ReportResults(
            string outputFolderPath,
            bool outputToStatsFile)
        {
            ErrorInfoComparerClass iErrorInfoComparerClass;

            string proteinName;

            int index;
            int retryCount;

            bool success;
            bool fileAlreadyExists;

            try
            {
                var outputOptions = new udtOutputOptionsType()
                {
                    OutputToStatsFile = outputToStatsFile,
                    SepChar = "\t"
                };

                try
                {
                    outputOptions.SourceFile = Path.GetFileName(mFastaFilePath);
                }
                catch (Exception ex)
                {
                    outputOptions.SourceFile = "Unknown_filename_due_to_error.fasta";
                }

                if (outputToStatsFile)
                {
                    mStatsFilePath = ConstructStatsFilePath(outputFolderPath);
                    fileAlreadyExists = File.Exists(mStatsFilePath);

                    success = false;
                    retryCount = 0;

                    while (!success && retryCount < 5)
                    {
                        FileStream outStream = null;
                        try
                        {
                            outStream = new FileStream(mStatsFilePath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
                            var outFileWriter = new StreamWriter(outStream);

                            outputOptions.OutFile = outFileWriter;

                            success = true;
                        }
                        catch (Exception ex)
                        {
                            // Failed to open file, wait 1 second, then try again
                            if (outStream != null)
                            {
                                outStream.Close();
                            }

                            retryCount += 1;
                            System.Threading.Thread.Sleep(1000);
                        }
                    }

                    if (success)
                    {
                        outputOptions.SepChar = "\t";
                        if (!fileAlreadyExists)
                        {
                            // Write the header line
                            var headers = new List<string>()
                            {
                                "Date",
                                "SourceFile",
                                "MessageType",
                                "LineNumber",
                                "ColumnNumber",
                                "Description_or_Protein",
                                "Info",
                                "Context"
                            };

                            outputOptions.OutFile.WriteLine(string.Join(outputOptions.SepChar, headers));
                        }
                    }
                    else
                    {
                        outputOptions.SepChar = ", ";
                        outputToStatsFile = false;
                        SetLocalErrorCode(eValidateFastaFileErrorCodes.ErrorCreatingStatsFile);
                    }
                }
                else
                {
                    outputOptions.SepChar = ", ";
                }

                ReportResultAddEntry(
                    outputOptions, eMsgTypeConstants.StatusMsg,
                    "Full path to file", mFastaFilePath);

                ReportResultAddEntry(
                    outputOptions, eMsgTypeConstants.StatusMsg,
                    "Protein count", mProteinCount.ToString("#,##0"));

                ReportResultAddEntry(
                    outputOptions, eMsgTypeConstants.StatusMsg,
                    "Residue count", mResidueCount.ToString("#,##0"));

                if (mFileErrorCount > 0)
                {
                    ReportResultAddEntry(
                        outputOptions, eMsgTypeConstants.ErrorMsg,
                        "Error count", GetErrorWarningCounts(eMsgTypeConstants.ErrorMsg, ErrorWarningCountTypes.Total).ToString());

                    if (mFileErrorCount > 1)
                    {
                        iErrorInfoComparerClass = new ErrorInfoComparerClass();
                        Array.Sort(mFileErrors, 0, mFileErrorCount, iErrorInfoComparerClass);
                    }

                    for (index = 0; index <= mFileErrorCount - 1; index++)
                    {
                        var fileError = mFileErrors[index];
                        if (fileError.ProteinName == null || fileError.ProteinName.Length == 0)
                        {
                            proteinName = "N/A";
                        }
                        else
                        {
                            proteinName = string.Copy(fileError.ProteinName);
                        }

                        string messageDescription = LookupMessageDescription(fileError.MessageCode, fileError.ExtraInfo);

                        ReportResultAddEntry(
                            outputOptions, eMsgTypeConstants.ErrorMsg,
                            fileError.LineNumber,
                            fileError.ColNumber,
                            proteinName,
                            messageDescription,
                            fileError.Context);
                    }
                }

                if (mFileWarningCount > 0)
                {
                    ReportResultAddEntry(
                        outputOptions, eMsgTypeConstants.WarningMsg,
                        "Warning count",
                        GetErrorWarningCounts(eMsgTypeConstants.WarningMsg, ErrorWarningCountTypes.Total).ToString());

                    if (mFileWarningCount > 1)
                    {
                        iErrorInfoComparerClass = new ErrorInfoComparerClass();
                        Array.Sort(mFileWarnings, 0, mFileWarningCount, iErrorInfoComparerClass);
                    }

                    for (index = 0; index <= mFileWarningCount - 1; index++)
                    {
                        var fileWarning = mFileWarnings[index];
                        if (fileWarning.ProteinName == null || fileWarning.ProteinName.Length == 0)
                        {
                            proteinName = "N/A";
                        }
                        else
                        {
                            proteinName = string.Copy(fileWarning.ProteinName);
                        }

                        ReportResultAddEntry(
                            outputOptions, eMsgTypeConstants.WarningMsg,
                            fileWarning.LineNumber,
                            fileWarning.ColNumber,
                            proteinName,
                            LookupMessageDescription(fileWarning.MessageCode, fileWarning.ExtraInfo),
                            fileWarning.Context);
                    }
                }

                var fastaFile = new FileInfo(mFastaFilePath);

                // # Proteins, # Peptides, FileSizeKB
                ReportResultAddEntry(
                    outputOptions, eMsgTypeConstants.StatusMsg,
                    "Summary line",
                    mProteinCount.ToString() + " proteins, " + mResidueCount.ToString() + " residues, " + (fastaFile.Length / 1024.0).ToString("0") + " KB");

                if (outputToStatsFile && outputOptions.OutFile != null)
                {
                    outputOptions.OutFile.Close();
                }
            }
            catch (Exception ex)
            {
                OnErrorEvent("Error in ReportResults", ex);
            }
        }

        private void ReportResultAddEntry(
            udtOutputOptionsType outputOptions,
            eMsgTypeConstants entryType,
            string descriptionOrProteinName,
            string info,
            string context = "")
        {
            ReportResultAddEntry(
                outputOptions,
                entryType, 0, 0,
                descriptionOrProteinName,
                info,
                context);
        }

        private void ReportResultAddEntry(
            udtOutputOptionsType outputOptions,
            eMsgTypeConstants entryType,
            int lineNumber,
            int colNumber,
            string descriptionOrProteinName,
            string info,
            string context)
        {
            var dataColumns = new List<string>()
            {
                outputOptions.SourceFile,
                LookupMessageType(entryType),
                lineNumber.ToString(),
                colNumber.ToString(),
                descriptionOrProteinName,
                info
            };

            if (context != null && context.Length > 0)
            {
                dataColumns.Add(context);
            }

            string message = string.Join(outputOptions.SepChar, dataColumns);

            if (outputOptions.OutputToStatsFile)
            {
                outputOptions.OutFile.WriteLine(GetTimeStamp() + outputOptions.SepChar + message);
            }
            else
            {
                Console.WriteLine(message);
            }
        }

        private void ResetStructures()
        {
            // This is used to reset the error arrays and stats variables

            mLineCount = 0;
            mProteinCount = 0;
            mResidueCount = 0;

            mFixedFastaStats.TruncatedProteinNameCount = 0;
            mFixedFastaStats.UpdatedResidueLines = 0;
            mFixedFastaStats.ProteinNamesInvalidCharsReplaced = 0;
            mFixedFastaStats.ProteinNamesMultipleRefsRemoved = 0;
            mFixedFastaStats.DuplicateNameProteinsSkipped = 0;
            mFixedFastaStats.DuplicateNameProteinsRenamed = 0;
            mFixedFastaStats.DuplicateSequenceProteinsSkipped = 0;

            mFileErrorCount = 0;
            mFileErrors = new udtMsgInfoType[0];
            ResetItemSummaryStructure(ref mFileErrorStats);

            mFileWarningCount = 0;
            mFileWarnings = new udtMsgInfoType[0];
            ResetItemSummaryStructure(ref mFileWarningStats);

            AbortProcessing = false;
        }

        private void ResetItemSummaryStructure(ref udtItemSummaryIndexedType itemSummary)
        {
            itemSummary.ErrorStatsCount = 0;
            itemSummary.ErrorStats = new udtErrorStatsType[0];
            if (itemSummary.MessageCodeToArrayIndex == null)
            {
                itemSummary.MessageCodeToArrayIndex = new Dictionary<int, int>();
            }
            else
            {
                itemSummary.MessageCodeToArrayIndex.Clear();
            }
        }

        private void SaveRulesToParameterFile(XmlSettingsFileAccessor settingsFile, string sectionName, IList<udtRuleDefinitionType> rules)
        {
            int ruleNumber;
            string ruleBase;

            if (rules == null || rules.Count <= 0)
            {
                settingsFile.SetParam(sectionName, XML_OPTION_ENTRY_RULE_COUNT, 0);
            }
            else
            {
                settingsFile.SetParam(sectionName, XML_OPTION_ENTRY_RULE_COUNT, rules.Count);

                for (ruleNumber = 1; ruleNumber <= rules.Count; ruleNumber++)
                {
                    ruleBase = "Rule" + ruleNumber.ToString();

                    var rule = rules[ruleNumber - 1];
                    settingsFile.SetParam(sectionName, ruleBase + "MatchRegEx", rule.MatchRegEx);
                    settingsFile.SetParam(sectionName, ruleBase + "MatchIndicatesProblem", rule.MatchIndicatesProblem);
                    settingsFile.SetParam(sectionName, ruleBase + "MessageWhenProblem", rule.MessageWhenProblem);
                    settingsFile.SetParam(sectionName, ruleBase + "Severity", rule.Severity);
                    settingsFile.SetParam(sectionName, ruleBase + "DisplayMatchAsExtraInfo", rule.DisplayMatchAsExtraInfo);
                }
            }
        }

        public bool SaveSettingsToParameterFile(string parameterFilePath)
        {
            // Save a model parameter file

            StreamWriter srOutFile;
            var settingsFile = new XmlSettingsFileAccessor();

            try
            {
                if (parameterFilePath == null || parameterFilePath.Length == 0)
                {
                    // No parameter file specified; do not save the settings
                    return true;
                }

                if (!File.Exists(parameterFilePath))
                {
                    // Need to generate a blank XML settings file

                    srOutFile = new StreamWriter(parameterFilePath, false);

                    srOutFile.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
                    srOutFile.WriteLine("<sections>");
                    srOutFile.WriteLine("  <section name=\"" + XML_SECTION_OPTIONS + "\">");
                    srOutFile.WriteLine("  </section>");
                    srOutFile.WriteLine("</sections>");

                    srOutFile.Close();
                }

                settingsFile.LoadSettings(parameterFilePath);

                // Save the general settings

                settingsFile.SetParam(XML_SECTION_OPTIONS, "AddMissingLinefeedAtEOF", GetOptionSwitchValue(SwitchOptions.AddMissingLineFeedAtEOF));
                settingsFile.SetParam(XML_SECTION_OPTIONS, "AllowAsteriskInResidues", GetOptionSwitchValue(SwitchOptions.AllowAsteriskInResidues));
                settingsFile.SetParam(XML_SECTION_OPTIONS, "AllowDashInResidues", GetOptionSwitchValue(SwitchOptions.AllowDashInResidues));

                settingsFile.SetParam(XML_SECTION_OPTIONS, "CheckForDuplicateProteinNames", GetOptionSwitchValue(SwitchOptions.CheckForDuplicateProteinNames));
                settingsFile.SetParam(XML_SECTION_OPTIONS, "CheckForDuplicateProteinSequences", GetOptionSwitchValue(SwitchOptions.CheckForDuplicateProteinSequences));
                settingsFile.SetParam(XML_SECTION_OPTIONS, "SaveProteinSequenceHashInfoFiles", GetOptionSwitchValue(SwitchOptions.SaveProteinSequenceHashInfoFiles));
                settingsFile.SetParam(XML_SECTION_OPTIONS, "SaveBasicProteinHashInfoFile", GetOptionSwitchValue(SwitchOptions.SaveBasicProteinHashInfoFile));

                settingsFile.SetParam(XML_SECTION_OPTIONS, "MaximumFileErrorsToTrack", MaximumFileErrorsToTrack);
                settingsFile.SetParam(XML_SECTION_OPTIONS, "MinimumProteinNameLength", MinimumProteinNameLength);
                settingsFile.SetParam(XML_SECTION_OPTIONS, "MaximumProteinNameLength", MaximumProteinNameLength);
                settingsFile.SetParam(XML_SECTION_OPTIONS, "MaximumResiduesPerLine", MaximumResiduesPerLine);

                settingsFile.SetParam(XML_SECTION_OPTIONS, "WarnBlankLinesBetweenProteins", GetOptionSwitchValue(SwitchOptions.WarnBlankLinesBetweenProteins));
                settingsFile.SetParam(XML_SECTION_OPTIONS, "WarnLineStartsWithSpace", GetOptionSwitchValue(SwitchOptions.WarnLineStartsWithSpace));
                settingsFile.SetParam(XML_SECTION_OPTIONS, "OutputToStatsFile", GetOptionSwitchValue(SwitchOptions.OutputToStatsFile));
                settingsFile.SetParam(XML_SECTION_OPTIONS, "NormalizeFileLineEndCharacters", GetOptionSwitchValue(SwitchOptions.NormalizeFileLineEndCharacters));

                settingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "GenerateFixedFASTAFile", GetOptionSwitchValue(SwitchOptions.GenerateFixedFASTAFile));
                settingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "SplitOutMultipleRefsInProteinName", GetOptionSwitchValue(SwitchOptions.SplitOutMultipleRefsInProteinName));

                settingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "RenameDuplicateNameProteins", GetOptionSwitchValue(SwitchOptions.FixedFastaRenameDuplicateNameProteins));
                settingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "KeepDuplicateNamedProteins", GetOptionSwitchValue(SwitchOptions.FixedFastaKeepDuplicateNamedProteins));

                settingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ConsolidateDuplicateProteinSeqs", GetOptionSwitchValue(SwitchOptions.FixedFastaConsolidateDuplicateProteinSeqs));
                settingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ConsolidateDupsIgnoreILDiff", GetOptionSwitchValue(SwitchOptions.FixedFastaConsolidateDupsIgnoreILDiff));

                settingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "TruncateLongProteinNames", GetOptionSwitchValue(SwitchOptions.FixedFastaTruncateLongProteinNames));
                settingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "SplitOutMultipleRefsForKnownAccession", GetOptionSwitchValue(SwitchOptions.FixedFastaSplitOutMultipleRefsForKnownAccession));
                settingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "WrapLongResidueLines", GetOptionSwitchValue(SwitchOptions.FixedFastaWrapLongResidueLines));
                settingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "RemoveInvalidResidues", GetOptionSwitchValue(SwitchOptions.FixedFastaRemoveInvalidResidues));

                settingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "LongProteinNameSplitChars", LongProteinNameSplitChars);
                settingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ProteinNameInvalidCharsToRemove", ProteinNameInvalidCharsToRemove);
                settingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ProteinNameFirstRefSepChars", ProteinNameFirstRefSepChars);
                settingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ProteinNameSubsequentRefSepChars", ProteinNameSubsequentRefSepChars);

                // Save the rules
                SaveRulesToParameterFile(settingsFile, XML_SECTION_FASTA_HEADER_LINE_RULES, mHeaderLineRules);
                SaveRulesToParameterFile(settingsFile, XML_SECTION_FASTA_PROTEIN_NAME_RULES, mProteinNameRules);
                SaveRulesToParameterFile(settingsFile, XML_SECTION_FASTA_PROTEIN_DESCRIPTION_RULES, mProteinDescriptionRules);
                SaveRulesToParameterFile(settingsFile, XML_SECTION_FASTA_PROTEIN_SEQUENCE_RULES, mProteinSequenceRules);

                // Commit the new settings to disk
                settingsFile.SaveSettings();

                // Need to re-open the parameter file and replace instances of "&gt;" with ">" and "&lt;" with "<"
                ReplaceXMLCodesWithText(parameterFilePath);
            }
            catch (Exception ex)
            {
                OnErrorEvent("Error in SaveSettingsToParameterFile", ex);
                return false;
            }

            return true;
        }

        private bool SearchRulesForID(
            IList<udtRuleDefinitionType> rules,
            int errorMessageCode,
            out string message)
        {
            if (rules != null)
            {
                for (int index = 0; index <= rules.Count - 1; index++)
                {
                    if (rules[index].CustomRuleID == errorMessageCode)
                    {
                        message = rules[index].MessageWhenProblem;
                        return true;
                    }
                }
            }

            message = null;
            return false;
        }

        /// <summary>
        /// Updates the validation rules using the current options
        /// </summary>
        /// <remarks>Call this function after setting new options using SetOptionSwitch</remarks>
        public void SetDefaultRules()
        {
            ClearAllRules();

            // For the rules, severity level 1 to 4 is warning; severity 5 or higher is an error

            // Header line errors
            SetRule(RuleTypes.HeaderLine, @"^>[ \t]*$", true, "Line starts with > but does not contain a protein name", DEFAULT_ERROR_SEVERITY);
            SetRule(RuleTypes.HeaderLine, @"^>[ \t].+", true, "Space or tab found directly after the > symbol", DEFAULT_ERROR_SEVERITY);

            // Header line warnings
            SetRule(RuleTypes.HeaderLine, @"^>[^ \t]+[ \t]*$", true, MESSAGE_TEXT_PROTEIN_DESCRIPTION_MISSING, DEFAULT_WARNING_SEVERITY);
            SetRule(RuleTypes.HeaderLine, @"^>[^ \t]+\t", true, "Protein name is separated from the protein description by a tab", DEFAULT_WARNING_SEVERITY);

            // Protein Name error characters
            string allowedChars = @"A-Za-z0-9.\-_:,\|/()\[\]\=\+#";

            if (mAllowAllSymbolsInProteinNames)
            {
                allowedChars += @"!@$%^&*<>?,\\";
            }

            string allowedCharsMatchSpec = "[^" + allowedChars + "]";

            SetRule(RuleTypes.ProteinName, allowedCharsMatchSpec, true, "Protein name contains invalid characters", DEFAULT_ERROR_SEVERITY, true);

            // Protein name warnings

            // Note that .*? changes .* from being greedy to being lazy
            SetRule(RuleTypes.ProteinName, "[:|].*?[:|;].*?[:|;]", true, "Protein name contains 3 or more vertical bars", DEFAULT_WARNING_SEVERITY + 1, true);

            if (!mAllowAllSymbolsInProteinNames)
            {
                SetRule(RuleTypes.ProteinName, @"[/()\[\],]", true, "Protein name contains undesirable characters", DEFAULT_WARNING_SEVERITY, true);
            }

            // Protein description warnings
            SetRule(RuleTypes.ProteinDescription, "\"", true, "Protein description contains a quotation mark", DEFAULT_WARNING_SEVERITY);
            SetRule(RuleTypes.ProteinDescription, @"\t", true, "Protein description contains a tab character", DEFAULT_WARNING_SEVERITY);
            SetRule(RuleTypes.ProteinDescription, @"\\/", true, @"Protein description contains an escaped slash: \/", DEFAULT_WARNING_SEVERITY);
            SetRule(RuleTypes.ProteinDescription, @"[\x00-\x08\x0E-\x1F]", true, "Protein description contains an escape code character", DEFAULT_ERROR_SEVERITY);
            SetRule(RuleTypes.ProteinDescription, ".{900,}", true, MESSAGE_TEXT_PROTEIN_DESCRIPTION_TOO_LONG, DEFAULT_WARNING_SEVERITY + 1, false);

            // Protein sequence errors
            SetRule(RuleTypes.ProteinSequence, @"[ \t]", true, "A space or tab was found in the residues", DEFAULT_ERROR_SEVERITY);

            if (!mAllowAsteriskInResidues)
            {
                SetRule(RuleTypes.ProteinSequence, @"\*", true, MESSAGE_TEXT_ASTERISK_IN_RESIDUES, DEFAULT_ERROR_SEVERITY);
            }

            if (!mAllowDashInResidues)
            {
                SetRule(RuleTypes.ProteinSequence, @"\-", true, MESSAGE_TEXT_DASH_IN_RESIDUES, DEFAULT_ERROR_SEVERITY);
            }

            // Note: we look for a space, tab, asterisk, and dash with separate rules (defined above)
            // Thus they are "allowed" by this RegEx, even though we may flag them as warnings with a different RegEx
            // We look for non-standard amino acids with warning rules (defined below)
            SetRule(RuleTypes.ProteinSequence, @"[^A-Z \t\*\-]", true, "Invalid residues found", DEFAULT_ERROR_SEVERITY, true);

            // Protein residue warnings
            // MS-GF+ treats these residues as stop characters(meaning no identified peptide will ever contain B, J, O, U, X, or Z)

            // SEQUEST uses mass 114.53494 for B (average of N and D)
            SetRule(RuleTypes.ProteinSequence, "B", true, "Residues line contains B (non-standard amino acid for N or D)", DEFAULT_WARNING_SEVERITY - 1);

            // Unsupported by SEQUEST
            SetRule(RuleTypes.ProteinSequence, "J", true, "Residues line contains J (non-standard amino acid)", DEFAULT_WARNING_SEVERITY - 1);

            // SEQUEST uses mass 114.07931 for O
            SetRule(RuleTypes.ProteinSequence, "O", true, "Residues line contains O (non-standard amino acid, ornithine)", DEFAULT_WARNING_SEVERITY - 1);

            // Unsupported by SEQUEST
            SetRule(RuleTypes.ProteinSequence, "U", true, "Residues line contains U (non-standard amino acid, selenocysteine)", DEFAULT_WARNING_SEVERITY);

            // SEQUEST uses mass 113.08406 for X (same as L and I)
            SetRule(RuleTypes.ProteinSequence, "X", true, "Residues line contains X (non-standard amino acid for L or I)", DEFAULT_WARNING_SEVERITY - 1);

            // SEQUEST uses mass 128.55059 for Z (average of Q and E)
            SetRule(RuleTypes.ProteinSequence, "Z", true, "Residues line contains Z (non-standard amino acid for Q or E)", DEFAULT_WARNING_SEVERITY - 1);
        }

        private void SetLocalErrorCode(eValidateFastaFileErrorCodes eNewErrorCode)
        {
            SetLocalErrorCode(eNewErrorCode, false);
        }

        private void SetLocalErrorCode(
            eValidateFastaFileErrorCodes eNewErrorCode,
            bool leaveExistingErrorCodeUnchanged)
        {
            if (leaveExistingErrorCodeUnchanged && mLocalErrorCode != eValidateFastaFileErrorCodes.NoError)
            {
                // An error code is already defined; do not change it
            }
            else
            {
                mLocalErrorCode = eNewErrorCode;

                if (eNewErrorCode == eValidateFastaFileErrorCodes.NoError)
                {
                    if (ErrorCode == ProcessFilesErrorCodes.LocalizedError)
                    {
                        SetBaseClassErrorCode(ProcessFilesErrorCodes.NoError);
                    }
                }
                else
                {
                    SetBaseClassErrorCode(ProcessFilesErrorCodes.LocalizedError);
                }
            }
        }

        private void SetRule(
            RuleTypes ruleType,
            string regexToMatch,
            bool doesMatchIndicateProblem,
            string problemReturnMessage,
            short severityLevel)
        {
            SetRule(
                ruleType, regexToMatch,
                doesMatchIndicateProblem,
                problemReturnMessage,
                severityLevel, false);
        }

        private void SetRule(
            RuleTypes ruleType,
            string regexToMatch,
            bool doesMatchIndicateProblem,
            string problemReturnMessage,
            short severityLevel,
            bool displayMatchAsExtraInfo)
        {
            switch (ruleType)
            {
                case RuleTypes.HeaderLine:
                    SetRule(ref mHeaderLineRules, regexToMatch, doesMatchIndicateProblem, problemReturnMessage, severityLevel, displayMatchAsExtraInfo);
                    break;
                case RuleTypes.ProteinDescription:
                    SetRule(ref mProteinDescriptionRules, regexToMatch, doesMatchIndicateProblem, problemReturnMessage, severityLevel, displayMatchAsExtraInfo);
                    break;
                case RuleTypes.ProteinName:
                    SetRule(ref mProteinNameRules, regexToMatch, doesMatchIndicateProblem, problemReturnMessage, severityLevel, displayMatchAsExtraInfo);
                    break;
                case RuleTypes.ProteinSequence:
                    SetRule(ref mProteinSequenceRules, regexToMatch, doesMatchIndicateProblem, problemReturnMessage, severityLevel, displayMatchAsExtraInfo);
                    break;
            }
        }

        private void SetRule(
            ref udtRuleDefinitionType[] rules,
            string matchRegEx,
            bool matchIndicatesProblem,
            string messageWhenProblem,
            short severity,
            bool displayMatchAsExtraInfo)
        {
            if (rules == null || rules.Length == 0)
            {
                rules = new udtRuleDefinitionType[1];
            }
            else
            {
                var oldRules = rules;
                rules = new udtRuleDefinitionType[rules.Length + 1];
                Array.Copy(oldRules, rules, Math.Min(rules.Length + 1, oldRules.Length));
            }

            rules[rules.Length - 1].MatchRegEx = matchRegEx;
            rules[rules.Length - 1].MatchIndicatesProblem = matchIndicatesProblem;
            rules[rules.Length - 1].MessageWhenProblem = messageWhenProblem;
            rules[rules.Length - 1].Severity = severity;
            rules[rules.Length - 1].DisplayMatchAsExtraInfo = displayMatchAsExtraInfo;
            rules[rules.Length - 1].CustomRuleID = mMasterCustomRuleID;

            mMasterCustomRuleID += 1;
        }

        private bool SortFile(FileInfo proteinHashFile, int sortColumnIndex, string sortedFilePath)
        {
            var sortUtility = new FlexibleFileSortUtility.TextFileSorter();

            mSortUtilityErrorMessage = string.Empty;

            sortUtility.WorkingDirectoryPath = proteinHashFile.Directory.FullName;
            sortUtility.HasHeaderLine = true;
            sortUtility.ColumnDelimiter = "\t";
            sortUtility.MaxFileSizeMBForInMemorySort = 250;
            sortUtility.ChunkSizeMB = 250;
            sortUtility.SortColumn = sortColumnIndex + 1;
            sortUtility.SortColumnIsNumeric = false;

            // The sort utility uses CompareOrdinal (StringComparison.Ordinal)
            sortUtility.IgnoreCase = false;

            RegisterEvents(sortUtility);
            sortUtility.ErrorEvent += mSortUtility_ErrorEvent;

            bool success = sortUtility.SortFile(proteinHashFile.FullName, sortedFilePath);

            if (success)
            {
                Console.WriteLine();
                return true;
            }

            if (string.IsNullOrWhiteSpace(mSortUtilityErrorMessage))
            {
                ShowErrorMessage("Unknown error sorting " + proteinHashFile.Name);
            }
            else
            {
                ShowErrorMessage("Sort error: " + mSortUtilityErrorMessage);
            }

            Console.WriteLine();
            return false;
        }

        private void SplitFastaProteinHeaderLine(
            string headerLine,
            out string proteinName,
            out string proteinDescription,
            out int descriptionStartIndex)
        {
            proteinDescription = string.Empty;
            descriptionStartIndex = 0;

            // Make sure the protein name and description are valid
            // Find the first space and/or tab
            int spaceIndex = GetBestSpaceIndex(headerLine);

            // At this point, spaceIndex will contain the location of the space or tab separating the protein name and description
            // However, if the space or tab is directly after the > sign, then we cannot continue (if this is the case, then spaceIndex will be 1)
            if (spaceIndex > 1)
            {
                proteinName = headerLine.Substring(1, spaceIndex - 1);
                proteinDescription = headerLine.Substring(spaceIndex + 1);
                descriptionStartIndex = spaceIndex;
            }
            // Line does not contain a description
            else if (spaceIndex <= 0)
            {
                if (headerLine.Trim().Length <= 1)
                {
                    proteinName = string.Empty;
                }
                else
                {
                    // The line contains a protein name, but not a description
                    proteinName = headerLine.Substring(1);
                }
            }
            else
            {
                // Space or tab found directly after the > symbol
                proteinName = string.Empty;
            }
        }

        private bool VerifyLinefeedAtEOF(string strInputFilePath, bool blnAddCrLfIfMissing)
        {
            var blnNeedToAddCrLf = false;
            bool blnSuccess;

            try
            {
                // Open the input file and validate that the final characters are CrLf, simply CR, or simply LF
                using (var fsInFile = new FileStream(strInputFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    if (fsInFile.Length > 2)
                    {
                        fsInFile.Seek(-1, SeekOrigin.End);

                        int lastByte = fsInFile.ReadByte();

                        if (lastByte == 10 || lastByte == 13)
                        {
                            // File ends in a linefeed or carriage return character; that's good
                            blnNeedToAddCrLf = false;
                        }
                        else
                        {
                            blnNeedToAddCrLf = true;
                        }
                    }

                    if (blnNeedToAddCrLf)
                    {
                        if (blnAddCrLfIfMissing)
                        {
                            ShowMessage("Appending CrLf return to: " + Path.GetFileName(strInputFilePath));
                            fsInFile.WriteByte(13);

                            fsInFile.WriteByte(10);
                        }
                    }
                }

                blnSuccess = true;
            }
            catch (Exception ex)
            {
                SetLocalErrorCode(eValidateFastaFileErrorCodes.ErrorVerifyingLinefeedAtEOF);
                blnSuccess = false;
            }

            return blnSuccess;
        }

        private readonly Regex reAdditionalProtein = new Regex(@"(.+)-[a-z]\d*", RegexOptions.Compiled);

        private void WriteCachedProtein(
            string cachedProteinName,
            string cachedProteinDescription,
            TextWriter consolidatedFastaWriter,
            IList<clsProteinHashInfo> proteinSeqHashInfo,
            StringBuilder sbCachedProteinResidueLines,
            StringBuilder sbCachedProteinResidues,
            bool consolidateDuplicateProteinSeqsInFasta,
            bool consolidateDupsIgnoreILDiff,
            clsNestedStringDictionary<int> proteinNameFirst,
            clsNestedStringDictionary<string> duplicateProteinList,
            int lineCountRead,
            clsNestedStringDictionary<int> proteinsWritten)
        {
            string masterProteinName = string.Empty;
            string masterProteinInfo;

            string proteinHash;
            string lineOut = string.Empty;
            Match reMatch;

            bool keepProtein;
            bool skipDupProtein;

            var additionalProteinNames = new List<string>();

            int seqIndex;
            if (proteinNameFirst.TryGetValue(cachedProteinName, out seqIndex))
            {
                // cachedProteinName was found in proteinNameFirst

                int cachedSeqIndex;
                if (proteinsWritten.TryGetValue(cachedProteinName, out cachedSeqIndex))
                {
                    if (mFixedFastaOptions.KeepDuplicateNamedProteinsUnlessMatchingSequence)
                    {
                        // Keep this protein if its sequence hash differs from the first protein with this name
                        proteinHash = ComputeProteinHash(sbCachedProteinResidues, consolidateDupsIgnoreILDiff);
                        if ((proteinSeqHashInfo[seqIndex].SequenceHash ?? "") != (proteinHash ?? ""))
                        {
                            RecordFastaFileWarning(lineCountRead, 1, cachedProteinName, (int)eMessageCodeConstants.DuplicateProteinNameRetained);
                            keepProtein = true;
                        }
                        else
                        {
                            keepProtein = false;
                        }
                    }
                    else
                    {
                        keepProtein = false;
                    }
                }
                else
                {
                    keepProtein = true;
                    proteinsWritten.Add(cachedProteinName, seqIndex);
                }

                if (keepProtein && seqIndex >= 0)
                {
                    if (proteinSeqHashInfo[seqIndex].AdditionalProteins.Count() > 0)
                    {
                        // The protein has duplicate proteins
                        // Construct a list of the duplicate protein names

                        additionalProteinNames.Clear();
                        foreach (string additionalProtein in proteinSeqHashInfo[seqIndex].AdditionalProteins)
                        {
                            // Add the additional protein name if it is not of the form "BaseName-b", "BaseName-c", etc.
                            skipDupProtein = false;

                            if (additionalProtein == null)
                            {
                                skipDupProtein = true;
                            }
                            else if ((additionalProtein.ToLower() ?? "") == (cachedProteinName.ToLower() ?? ""))
                            {
                                // Names match; do not add to the list
                                skipDupProtein = true;
                            }
                            else
                            {
                                // Check whether additionalProtein looks like one of the following
                                // ProteinX-b
                                // ProteinX-a2
                                // ProteinX-d3
                                reMatch = reAdditionalProtein.Match(additionalProtein);

                                if (reMatch.Success)
                                {
                                    if ((cachedProteinName.ToLower() ?? "") == (reMatch.Groups[1].Value.ToLower() ?? ""))
                                    {
                                        // Base names match; do not add to the list
                                        // For example, ProteinX and ProteinX-b
                                        skipDupProtein = true;
                                    }
                                }
                            }

                            if (!skipDupProtein)
                            {
                                additionalProteinNames.Add(additionalProtein);
                            }
                        }

                        if (additionalProteinNames.Count > 0 && consolidateDuplicateProteinSeqsInFasta)
                        {
                            // Append the duplicate protein names to the description
                            // However, do not let the description get over 7995 characters in length
                            string updatedDescription = cachedProteinDescription + "; Duplicate proteins: " + FlattenArray(additionalProteinNames, ',');
                            if (updatedDescription.Length > MAX_PROTEIN_DESCRIPTION_LENGTH)
                            {
                                updatedDescription = updatedDescription.Substring(0, MAX_PROTEIN_DESCRIPTION_LENGTH - 3) + "...";
                            }

                            lineOut = ConstructFastaHeaderLine(cachedProteinName, updatedDescription);
                        }
                    }
                }
            }
            else
            {
                keepProtein = false;
                mFixedFastaStats.DuplicateSequenceProteinsSkipped += 1;

                if (!duplicateProteinList.TryGetValue(cachedProteinName, out masterProteinName))
                {
                    masterProteinInfo = "same as ??";
                }
                else
                {
                    masterProteinInfo = "same as " + masterProteinName;
                }

                RecordFastaFileWarning(lineCountRead, 0, cachedProteinName, (int)eMessageCodeConstants.ProteinRemovedSinceDuplicateSequence, masterProteinInfo, string.Empty);
            }

            if (keepProtein)
            {
                if (string.IsNullOrEmpty(lineOut))
                {
                    lineOut = ConstructFastaHeaderLine(cachedProteinName, cachedProteinDescription);
                }

                consolidatedFastaWriter.WriteLine(lineOut);
                consolidatedFastaWriter.Write(sbCachedProteinResidueLines.ToString());
            }
        }

        private void WriteCachedProteinHashMetadata(
            TextWriter proteinNamesToKeepWriter,
            TextWriter uniqueProteinSeqsWriter,
            TextWriter uniqueProteinSeqDuplicateWriter,
            int currentSequenceIndex,
            string sequenceHash,
            string sequenceLength,
            ICollection<string> proteinNames)
        {
            if (proteinNames.Count < 1)
            {
                throw new Exception("proteinNames is empty in WriteCachedProteinHashMetadata; this indicates a logic error");
            }

            string firstProtein = proteinNames.ElementAtOrDefault(0);
            string duplicateProteinList;

            // Append the first protein name to the _ProteinsToKeep.tmp file
            proteinNamesToKeepWriter.WriteLine(firstProtein);

            if (proteinNames.Count == 1)
            {
                duplicateProteinList = string.Empty;
            }
            else
            {
                var duplicateProteins = new StringBuilder();

                for (int proteinIndex = 1; proteinIndex <= proteinNames.Count - 1; proteinIndex++)
                {
                    if (duplicateProteins.Length > 0)
                    {
                        duplicateProteins.Append(",");
                    }

                    duplicateProteins.Append(proteinNames.ElementAtOrDefault(proteinIndex));

                    var dataValuesSeqDuplicate = new List<string>()
                    {
                        currentSequenceIndex.ToString(),
                        firstProtein,
                        sequenceLength,
                        proteinNames.ElementAtOrDefault(proteinIndex)
                    };
                    uniqueProteinSeqDuplicateWriter.WriteLine(FlattenList(dataValuesSeqDuplicate));
                }

                duplicateProteinList = duplicateProteins.ToString();
            }

            var dataValues = new List<string>()
            {
                currentSequenceIndex.ToString(),
                proteinNames.ElementAtOrDefault(0),
                sequenceLength,
                sequenceHash,
                proteinNames.Count.ToString(),
                duplicateProteinList
            };

            uniqueProteinSeqsWriter.WriteLine(FlattenList(dataValues));
        }

        #region "Event Handlers"

        private void mSortUtility_ErrorEvent(string message, Exception ex)
        {
            mSortUtilityErrorMessage = message;
        }

        #endregion

        // IComparer class to allow comparison of udtMsgInfoType items
        private class ErrorInfoComparerClass : IComparer<udtMsgInfoType>
        {
            public int Compare(udtMsgInfoType x, udtMsgInfoType y)
            {
                var errorInfo1 = x;
                var errorInfo2 = y;

                if (errorInfo1.MessageCode > errorInfo2.MessageCode)
                {
                    return 1;
                }
                else if (errorInfo1.MessageCode < errorInfo2.MessageCode)
                {
                    return -1;
                }
                else if (errorInfo1.LineNumber > errorInfo2.LineNumber)
                {
                    return 1;
                }
                else if (errorInfo1.LineNumber < errorInfo2.LineNumber)
                {
                    return -1;
                }
                else if (errorInfo1.ColNumber > errorInfo2.ColNumber)
                {
                    return 1;
                }
                else if (errorInfo1.ColNumber < errorInfo2.ColNumber)
                {
                    return -1;
                }
                else
                {
                    return 0;
                }
            }
        }
    }
}