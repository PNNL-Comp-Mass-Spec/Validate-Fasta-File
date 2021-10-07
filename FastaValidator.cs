using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using PRISM;

namespace ValidateFastaFile
{
    // Ignore Spelling: A-Za-z, Diff, Dups, gi, jgi, Lf, Mem, ornithine, pre, selenocysteine, Sep, seqs, Validator, varchar, yyyy-MM-dd

    /// <summary>
    /// Old FASTA file validator class name
    /// </summary>
    [Obsolete("Renamed to 'FastaValidator'", true)]
    // ReSharper disable once InconsistentNaming
    public class clsValidateFastaFile : FastaValidator
    {
        /// <summary>
        /// Parameterless constructor
        /// </summary>
        public clsValidateFastaFile() : base() { }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="parameterFilePath"></param>
        public clsValidateFastaFile(string parameterFilePath) : base(parameterFilePath) { }
    }

    /// <summary>
    /// This class will read a protein FASTA file and validate its contents
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
    public class FastaValidator : PRISM.FileProcessor.ProcessFilesBase
    {
        /// <summary>
        /// Constructor
        /// </summary>
        public FastaValidator()
        {
            mFileDate = "August 5, 2021";
            InitializeLocalVariables();
        }

        /// <summary>
        /// Constructor that takes a parameter file
        /// </summary>
        /// <param name="parameterFilePath"></param>
        public FastaValidator(string parameterFilePath) : this()
        {
            LoadParameterFileSettings(parameterFilePath);
        }

        private const int DEFAULT_MINIMUM_PROTEIN_NAME_LENGTH = 3;

        /// <summary>
        /// The maximum suggested value when using SEQUEST is 34 characters
        /// In contrast, MS-GF+ supports long protein names
        /// </summary>
        public const int DEFAULT_MAXIMUM_PROTEIN_NAME_LENGTH = 60;

        private const int DEFAULT_MAXIMUM_RESIDUES_PER_LINE = 120;

        /// <summary>
        /// Default protein line start character
        /// </summary>
        public const char DEFAULT_PROTEIN_LINE_START_CHAR = '>';

        /// <summary>
        /// Default long protein name split char
        /// </summary>
        public const char DEFAULT_LONG_PROTEIN_NAME_SPLIT_CHAR = '|';

        /// <summary>
        /// Default protein name first reference split chars
        /// </summary>
        public const string DEFAULT_PROTEIN_NAME_FIRST_REF_SEP_CHARS = ":|";

        /// <summary>
        /// Default protein name subsequent reference separation chars
        /// </summary>
        public const string DEFAULT_PROTEIN_NAME_SUBSEQUENT_REF_SEP_CHARS = ":|;";

        private const char INVALID_PROTEIN_NAME_CHAR_REPLACEMENT = '_';

        private const int CUSTOM_RULE_ID_START = 1000;

        private const int DEFAULT_CONTEXT_LENGTH = 13;

        /// <summary>
        /// Protein description missing message
        /// </summary>
        public const string MESSAGE_TEXT_PROTEIN_DESCRIPTION_MISSING = "Line contains a protein name, but not a description";

        /// <summary>
        /// Protein description too long message
        /// </summary>
        public const string MESSAGE_TEXT_PROTEIN_DESCRIPTION_TOO_LONG = "Protein description is over 900 characters long";

        /// <summary>
        /// Asterisk found in the residues message
        /// </summary>
        public const string MESSAGE_TEXT_ASTERISK_IN_RESIDUES = "An asterisk was found in the residues";

        /// <summary>
        /// Dash found in the residues message
        /// </summary>
        public const string MESSAGE_TEXT_DASH_IN_RESIDUES = "A dash was found in the residues";

        /// <summary>
        /// Option section name in the XML parameter file
        /// </summary>
        public const string XML_SECTION_OPTIONS = "ValidateFastaFileOptions";

        /// <summary>
        /// Fixed FASTA file options section in the XML parameter file
        /// </summary>
        public const string XML_SECTION_FIXED_FASTA_FILE_OPTIONS = "ValidateFastaFixedFASTAFileOptions";

        /// <summary>
        /// Fixed FASTA file options section in the XML parameter file
        /// </summary>
        public const string XML_SECTION_FASTA_HEADER_LINE_RULES = "ValidateFastaHeaderLineRules";

        /// <summary>
        /// Protein name rules section in the XML parameter file
        /// </summary>
        public const string XML_SECTION_FASTA_PROTEIN_NAME_RULES = "ValidateFastaProteinNameRules";

        /// <summary>
        /// Protein description rules section in the XML parameter file
        /// </summary>
        public const string XML_SECTION_FASTA_PROTEIN_DESCRIPTION_RULES = "ValidateFastaProteinDescriptionRules";

        /// <summary>
        /// Protein sequence rules section in the XML parameter file
        /// </summary>
        public const string XML_SECTION_FASTA_PROTEIN_SEQUENCE_RULES = "ValidateFastaProteinSequenceRules";

        /// <summary>
        /// RuleCount element name
        /// </summary>
        public const string XML_OPTION_ENTRY_RULE_COUNT = "RuleCount";

        /// <summary>
        /// Maximum protein description length
        /// </summary>
        /// <remarks>
        /// The value of 7995 is chosen because the maximum varchar() value in SQL Server is varchar(8000)
        /// and we want to prevent truncation errors when importing protein names and descriptions into SQL Server
        /// </remarks>
        public const int MAX_PROTEIN_DESCRIPTION_LENGTH = 7995;

        private const string MEM_USAGE_PREFIX = "MemUsage: ";
        private const bool REPORT_DETAILED_MEMORY_USAGE = false;

        private const string PROTEIN_NAME_COLUMN = "Protein_Name";
        private const string SEQUENCE_LENGTH_COLUMN = "Sequence_Length";
        private const string SEQUENCE_HASH_COLUMN = "Sequence_Hash";
        private const string PROTEIN_HASHES_FILENAME_SUFFIX = "_ProteinHashes.txt";

        private const int DEFAULT_WARNING_SEVERITY = 3;
        private const int DEFAULT_ERROR_SEVERITY = 7;

        /// <summary>
        /// Message code constants
        /// </summary>
        /// <remarks>
        /// Custom rules start with message code CUSTOM_RULE_ID_START=1000, and therefore
        /// the values in enum MessageCodeConstants should all be less than CUSTOM_RULE_ID_START
        /// </remarks>
        public enum MessageCodeConstants
        {
#pragma warning disable 1591
            UnspecifiedError = 0,

            // Error messages
            ProteinNameIsTooLong = 1,
            LineStartsWithSpace = 2,
            // RightArrowFollowedBySpace = 3,
            // RightArrowFollowedByTab = 4,
            // RightArrowButNoProteinName = 5,
            BlankLineBetweenProteinNameAndResidues = 6,
            BlankLineInMiddleOfResidues = 7,
            ResiduesFoundWithoutProteinHeader = 8,
            ProteinEntriesNotFound = 9,
            FinalProteinEntryMissingResidues = 10,
            FileDoesNotEndWithLinefeed = 11,
            DuplicateProteinName = 12,

            // Warning messages
            ProteinNameIsTooShort = 13,
            // ProteinNameContainsVerticalBars = 14,
            // ProteinNameContainsWarningCharacters = 21,
            // ProteinNameWithoutDescription = 14,
            BlankLineBeforeProteinName = 15,
            // ProteinNameAndDescriptionSeparatedByTab = 16,
            // ProteinDescriptionWithTab = 25,
            // ProteinDescriptionWithQuotationMark = 26,
            // ProteinDescriptionWithEscapedSlash = 27,
            // ProteinDescriptionWithUndesirableCharacter = 28,
            ResiduesLineTooLong = 17,
            // ResiduesLineContainsU = 30,
            DuplicateProteinSequence = 18,
            RenamedProtein = 19,
            ProteinRemovedSinceDuplicateSequence = 20,
            DuplicateProteinNameRetained = 21
#pragma warning restore 1591
        }

        /// <summary>
        /// Error message info
        /// </summary>
        public class MsgInfo : IComparable<MsgInfo>
        {
            /// <summary>
            /// Line number of this error in the FASTA file
            /// </summary>
            public int LineNumber { get; }

            /// <summary>
            /// Column number of this error in the FASTA file
            /// </summary>
            public int ColNumber { get; }

            /// <summary>
            /// Column number of this error in the FASTA file
            /// </summary>
            public string ProteinName { get; }

            /// <summary>
            /// Error message code
            /// </summary>
            public int MessageCode { get; }

            /// <summary>
            /// Extra info about this error
            /// </summary>
            public string ExtraInfo { get; }

            /// <summary>
            /// Error message context
            /// </summary>
            public string Context { get; }

            /// <summary>
            /// Constructor that takes line number, column number, etc.
            /// </summary>
            /// <param name="lineNumber"></param>
            /// <param name="colNumber"></param>
            /// <param name="proteinName"></param>
            /// <param name="messageCode"></param>
            /// <param name="extraInfo"></param>
            /// <param name="context"></param>
            public MsgInfo(int lineNumber, int colNumber, string proteinName, int messageCode, string extraInfo, string context)
            {
                LineNumber = lineNumber;
                ColNumber = colNumber;
                ProteinName = proteinName;
                MessageCode = messageCode;
                ExtraInfo = extraInfo;
                Context = context;
            }

            /// <summary>
            /// Parameterless constructor
            /// </summary>
            public MsgInfo()
            {
                ProteinName = string.Empty;
                ExtraInfo = string.Empty;
                Context = string.Empty;
            }

            /// <summary>
            /// Return a string describing this error
            /// </summary>
            public override string ToString()
            {
                return string.Format("Line {0}, protein {1}, code {2}: {3}", LineNumber, ProteinName, MessageCode, ExtraInfo);
            }

            /// <summary>
            /// Compare one instance of this class to another
            /// </summary>
            /// <param name="other"></param>
            /// <returns>0 if the two instances match, otherwise -1 or 1 based on sort order</returns>
            public int CompareTo(MsgInfo other)
            {
                if (ReferenceEquals(this, other))
                    return 0;
                if (other is null)
                    return 1;

                var messageCodeComparison = MessageCode.CompareTo(other.MessageCode);
                if (messageCodeComparison != 0)
                    return messageCodeComparison;

                var lineNumberComparison = LineNumber.CompareTo(other.LineNumber);
                if (lineNumberComparison != 0)
                    return lineNumberComparison;

                return ColNumber.CompareTo(other.ColNumber);
            }
        }

        /// <summary>
        /// Options for reporting results
        /// </summary>
        public class OutputOptions
        {
            /// <summary>
            /// Filename of the FASTA file examined
            /// </summary>
            public string SourceFile { get; set; }

            /// <summary>
            /// When true, write message stats to a file
            /// </summary>
            public bool OutputToStatsFile { get; }

            /// <summary>
            /// Output file path
            /// </summary>
            public StreamWriter OutFile { get; set; }

            /// <summary>
            /// Column separation character
            /// </summary>
            public string SepChar { get; set; }

            /// <summary>
            /// Constructor
            /// </summary>
            /// <param name="outputToStatsFile"></param>
            /// <param name="sepChar"></param>
            public OutputOptions(bool outputToStatsFile, string sepChar)
            {
                OutputToStatsFile = outputToStatsFile;
                SepChar = sepChar;
            }

            /// <summary>
            /// Return the name of the FASTA file being analyzed
            /// </summary>
            public override string ToString()
            {
                return SourceFile;
            }
        }

#pragma warning disable 1591

        /// <summary>
        /// Validation rule types
        /// </summary>
        public enum RuleTypes
        {
            HeaderLine,
            ProteinName,
            ProteinDescription,
            ProteinSequence
        }

        /// <summary>
        /// Option switches
        /// </summary>
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

        /// <summary>
        /// Fixed FASTA stat categories
        /// </summary>
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

        /// <summary>
        /// Error warning count types
        /// </summary>
        public enum ErrorWarningCountTypes
        {
            Specified,
            Unspecified,
            Total
        }

        /// <summary>
        /// Message type constants
        /// </summary>
        public enum MsgTypeConstants
        {
            ErrorMsg = 0,
            WarningMsg = 1,
            StatusMsg = 2
        }

        /// <summary>
        /// Validation error codes
        /// </summary>
        public enum ValidateFastaFileErrorCodes
        {
            NoError = 0,
            OptionsSectionNotFound = 1,
            ErrorReadingInputFile = 2,
            ErrorCreatingStatsFile = 4,
            ErrorVerifyingLinefeedAtEOF = 8,
            UnspecifiedError = -1
        }

#pragma warning restore 1591

        /// <summary>
        /// Line ending characters
        /// </summary>
        public enum LineEndingCharacters
        {
            /// <summary>
            /// Windows
            /// </summary>
            CRLF,

            /// <summary>
            /// Old style Mac
            /// </summary>
            CR,

            /// <summary>
            /// Unix, Linux, OS X
            /// </summary>
            LF,

            /// <summary>
            /// Oddball (Just for completeness!)
            /// </summary>
            LFCR
        }

        /// <summary>
        /// Error stats container
        /// </summary>
        private class ErrorStats
        {
            /// <summary>
            /// Error code
            /// </summary>
            /// <remarks>Custom rules start with message code CUSTOM_RULE_ID_START</remarks>
            private int MessageCode { get; }

            /// <summary>
            /// Number of times detailed information about this error was stored in mFileErrors
            /// </summary>
            public int CountSpecified { get; set; }

            /// <summary>
            /// Number of additional occurrences of this error (where details were not stored in mFileErrors)
            /// </summary>
            public int CountUnspecified { get; set; }

            /// <summary>
            /// Constructor
            /// </summary>
            /// <param name="messageCode"></param>
            public ErrorStats(int messageCode)
            {
                MessageCode = messageCode;
            }

            /// <summary>
            /// Return the message code, count specified, and count unspecified
            /// </summary>
            public override string ToString()
            {
                return MessageCode + ": " + CountSpecified + " specified, " + CountUnspecified + " unspecified";
            }
        }

        /// <summary>
        /// Container for tracking errors and warnings
        /// </summary>
        private class MsgInfosAndSummary
        {
            /// <summary>
            /// Error messages
            /// </summary>
            public List<MsgInfo> Messages { get; } = new();

            /// <summary>
            /// Stats dictionary
            /// </summary>
            public Dictionary<int, ErrorStats> MessageCodeToErrorStats { get; } = new();

            /// <summary>
            /// Number of items in Messages
            /// </summary>
            [Obsolete("Use Messages.Count")]
            public int Count => Messages.Count;

            /// <summary>
            /// Clear cached messages
            /// </summary>
            public void Reset()
            {
                Messages.Clear();
                MessageCodeToErrorStats.Clear();
            }

            /// <summary>
            /// Return the sum of CountSpecified for all tracked messages
            /// </summary>
            // ReSharper disable once UnusedMember.Local
            public int ComputeTotalSpecifiedCount()
            {
                return MessageCodeToErrorStats.Values.Sum(stat => stat.CountSpecified);
            }

            /// <summary>
            /// Return the sum of CountUnspecified for all tracked messages
            /// </summary>
            public int ComputeTotalUnspecifiedCount()
            {
                return MessageCodeToErrorStats.Values.Sum(stat => stat.CountUnspecified);
            }
        }

        /// <summary>
        /// Validation rule definition container
        /// </summary>
        private class RuleDefinition
        {
            /// <summary>
            /// Rule RegEx
            /// </summary>
            public string MatchRegEx { get; }

            /// <summary>
            /// True means text matching the RegEx means a problem; false means if text doesn't match the RegEx, then that means a problem
            /// </summary>
            public bool MatchIndicatesProblem { get; set; }

            /// <summary>
            /// Message to display if a problem is present
            /// </summary>
            public string MessageWhenProblem { get; set; }

            /// <summary>
            /// 0 is lowest severity, 9 is highest severity; value >= 5 means error
            /// </summary>
            public short Severity { get; set; }

            /// <summary>
            /// If true, the matching text is stored as the context info
            /// </summary>
            public bool DisplayMatchAsExtraInfo { get; set; }

            /// <summary>
            /// Custom Rule ID
            /// </summary>
            /// <remarks>This value is auto-assigned</remarks>
            public int CustomRuleID { get; set; }

            /// <summary>
            /// Constructor
            /// </summary>
            /// <param name="matchRegEx"></param>
            public RuleDefinition(string matchRegEx)
            {
                MatchRegEx = matchRegEx;
            }

            /// <summary>
            /// Return the rule ID and message to display if a problem is present
            /// </summary>
            public override string ToString()
            {
                return CustomRuleID + ": " + MessageWhenProblem;
            }
        }

        /// <summary>
        /// Extended rule definition container
        /// </summary>
        private class RuleDefinitionExtended
        {
            /// <summary>
            /// Parent rule definition
            /// </summary>
            public RuleDefinition RuleDefinition { get; }

            public Regex MatchRegEx { get; }

            /// <summary>
            /// True if the rule is valid, false if a problem
            /// </summary>
            // ReSharper disable once UnusedAutoPropertyAccessor.Local
            public bool Valid { get; set; }

            /// <summary>
            /// Constructor
            /// </summary>
            /// <param name="ruleDefinition"></param>
            /// <param name="regexRule"></param>
            public RuleDefinitionExtended(RuleDefinition ruleDefinition, Regex regexRule)
            {
                RuleDefinition = ruleDefinition;
                MatchRegEx = regexRule;
            }

            /// <summary>
            /// Return the rule ID and message to display if a problem is present
            /// </summary>
            public override string ToString()
            {
                return RuleDefinition.CustomRuleID + ": " + RuleDefinition.MessageWhenProblem;
            }
        }

        /// <summary>
        /// Options container
        /// </summary>
        private class FixedFastaOptions
        {
            /// <summary>
            /// Split out multiple refs in protein name
            /// </summary>
            public bool SplitOutMultipleRefsInProteinName { get; set; }

            /// <summary>
            /// Split out multiple refs for known accession
            /// </summary>
            public bool SplitOutMultipleRefsForKnownAccession { get; set; }

            /// <summary>
            /// Long protein name split chars
            /// </summary>
            public char[] LongProteinNameSplitChars { get; set; }

            /// <summary>
            /// Protein name invalid chars to remove
            /// </summary>
            public char[] ProteinNameInvalidCharsToRemove { get; set; }

            /// <summary>
            /// Rename proteins with duplicate names
            /// </summary>
            public bool RenameProteinsWithDuplicateNames { get; set; }

            /// <summary>
            /// Keep duplicate named proteins unless matching sequence
            /// </summary>
            /// <remarks>Ignored if RenameProteinsWithDuplicateNames=true or ConsolidateProteinsWithDuplicateSeqs=true</remarks>
            public bool KeepDuplicateNamedProteinsUnlessMatchingSequence { get; set; }

            /// <summary>
            /// Consolidate proteins with duplicate sequences
            /// </summary>
            public bool ConsolidateProteinsWithDuplicateSeqs { get; set; }

            /// <summary>
            /// Ignore I/L differences when consolidating duplicates
            /// </summary>
            public bool ConsolidateDupsIgnoreILDiff { get; set; }

            /// <summary>
            /// Truncate long protein names
            /// </summary>
            public bool TruncateLongProteinNames { get; set; }

            /// <summary>
            /// Wrap long residue lines
            /// </summary>
            public bool WrapLongResidueLines { get; set; }

            /// <summary>
            /// Remove invalid residues
            /// </summary>
            public bool RemoveInvalidResidues { get; set; }

            /// <summary>
            /// Constructor
            /// </summary>
            public FixedFastaOptions()
            {
                LongProteinNameSplitChars = new[] { DEFAULT_LONG_PROTEIN_NAME_SPLIT_CHAR };

                // Default to an empty character array for invalid characters
                ProteinNameInvalidCharsToRemove = new char[] { };
            }
        }

        /// <summary>
        /// Fixed FASTA stats
        /// </summary>
        private class FixedFastaStats
        {
            /// <summary>
            /// Truncated protein name count
            /// </summary>
            public int TruncatedProteinNameCount { get; set; }

            /// <summary>
            /// Updated residue lines
            /// </summary>
            public int UpdatedResidueLines { get; set; }

            /// <summary>
            /// Protein names invalid chars replaced
            /// </summary>
            public int ProteinNamesInvalidCharsReplaced { get; set; }

            /// <summary>
            /// Protein names multiple refs removed
            /// </summary>
            public int ProteinNamesMultipleRefsRemoved { get; set; }

            /// <summary>
            /// Duplicate name proteins skipped
            /// </summary>
            public int DuplicateNameProteinsSkipped { get; set; }

            /// <summary>
            /// Duplicate name proteins renamed
            /// </summary>
            public int DuplicateNameProteinsRenamed { get; set; }

            /// <summary>
            /// Duplicate sequence proteins skipped
            /// </summary>
            public int DuplicateSequenceProteinsSkipped { get; set; }

            /// <summary>
            /// Constructor
            /// </summary>
            public FixedFastaStats()
            {
                Reset();
            }

            /// <summary>
            /// Reset all counts to 0
            /// </summary>
            public void Reset()
            {
                TruncatedProteinNameCount = 0;
                UpdatedResidueLines = 0;
                ProteinNamesInvalidCharsReplaced = 0;
                ProteinNamesMultipleRefsRemoved = 0;
                DuplicateNameProteinsSkipped = 0;
                DuplicateNameProteinsRenamed = 0;
                DuplicateSequenceProteinsSkipped = 0;
            }

            /// <summary>
            /// Get the specified statistic
            /// </summary>
            /// <param name="statCategory"></param>
            public int GetStat(FixedFASTAFileValues statCategory)
            {
                return statCategory switch
                {
                    FixedFASTAFileValues.DuplicateProteinNamesSkippedCount => DuplicateNameProteinsSkipped,
                    FixedFASTAFileValues.ProteinNamesInvalidCharsReplaced => ProteinNamesInvalidCharsReplaced,
                    FixedFASTAFileValues.ProteinNamesMultipleRefsRemoved => ProteinNamesMultipleRefsRemoved,
                    FixedFASTAFileValues.TruncatedProteinNameCount => TruncatedProteinNameCount,
                    FixedFASTAFileValues.UpdatedResidueLines => UpdatedResidueLines,
                    FixedFASTAFileValues.DuplicateProteinNamesRenamedCount => DuplicateNameProteinsRenamed,
                    FixedFASTAFileValues.DuplicateProteinSeqsSkippedCount => DuplicateSequenceProteinsSkipped,
                    _ => 0,
                };
            }
        }

        private class ProteinNameTruncationRegex
        {
            /// <summary>
            /// Extracts IPI:IPI00048500.11 from IPI:IPI00048500.11|ref|23848934 <br />
            /// Second matching group contains everything after the first vertical bar
            /// </summary>
            public Regex MatchIPI { get; }

            /// <summary>
            /// Extracts gi|169602219 from gi|169602219|ref|XP_001794531.1| <br />
            /// Second matching group contains everything after the second vertical bar
            /// </summary>
            public Regex MatchGI { get; }

            /// <summary>
            /// Extracts jgi|Batde5|906240 from jgi|Batde5|90624|GP3.061830 <br />
            /// Second matching group contains everything after the third vertical bar
            /// </summary>
            public Regex MatchJGI { get; }

            /// <summary>
            /// Extracts bob|234384 from bob|234384|ref|483293, or bob|845832 from bob|845832;ref|384923 <br />
            /// Second matching group contains everything after the separator following the first matched group
            /// </summary>
            public Regex MatchGeneric { get; }

            /// <summary>
            /// Matches jgi|Batde5|23435 ; it requires that there be a number after the second bar <br />
            /// Contains no matching groups
            /// </summary>
            public Regex MatchJGIBaseAndID { get; }

            /// <summary>
            /// Extracts the separator set following the first separator in the string
            /// </summary>
            public Regex MatchDoubleBarOrColonAndBar { get; }

            /// <summary>
            /// Constructor
            /// </summary>
            /// <param name="proteinNameFirstRefSepChars"></param>
            /// <param name="proteinNameSubsequentRefSepChars"></param>
            public ProteinNameTruncationRegex(char[] proteinNameFirstRefSepChars, char[] proteinNameSubsequentRefSepChars)
            {
                // Note that each of these RegEx tests contain two groups with captured text:

                // The following will extract IPI:IPI00048500.11 from IPI:IPI00048500.11|ref|23848934
                MatchIPI =
                    new Regex(@"^(IPI:IPI[\w.]{2,})\|(.+)",
                    RegexOptions.Singleline | RegexOptions.Compiled);

                // The following will extract gi|169602219 from gi|169602219|ref|XP_001794531.1|
                MatchGI =
                    new Regex(@"^(gi\|\d+)\|(.+)",
                    RegexOptions.Singleline | RegexOptions.Compiled);

                // The following will extract jgi|Batde5|906240 from jgi|Batde5|90624|GP3.061830
                MatchJGI =
                    new Regex(@"^(jgi\|[^|]+\|[^|]+)\|(.+)",
                    RegexOptions.Singleline | RegexOptions.Compiled);

                // The following will extract bob|234384 from  bob|234384|ref|483293
                // or bob|845832 from  bob|845832;ref|384923
                MatchGeneric =
                    new Regex(@"^(\w{2,}[" +
                        new string(proteinNameFirstRefSepChars) + @"][\w\d._]{2,})[" +
                        new string(proteinNameSubsequentRefSepChars) + "](.+)",
                        RegexOptions.Singleline | RegexOptions.Compiled);
                // The following matches jgi|Batde5|23435 ; it requires that there be a number after the second bar
                MatchJGIBaseAndID =
                    new Regex(@"^jgi\|[^|]+\|\d+",
                        RegexOptions.Singleline | RegexOptions.Compiled);

                // Note that this RegEx contains a group with captured text:
                MatchDoubleBarOrColonAndBar =
                    new Regex("[" +
                        new string(proteinNameFirstRefSepChars) + "][^" +
                        new string(proteinNameSubsequentRefSepChars) + "]*([" +
                        new string(proteinNameSubsequentRefSepChars) + "])",
                        RegexOptions.Singleline | RegexOptions.Compiled);
            }
        }

        /// <summary>
        /// FASTA file path being examined
        /// </summary>
        /// <remarks>Used by CustomValidateFastaFiles</remarks>
        protected string mFastaFilePath;

        private readonly FixedFastaStats mFixedFastaStats = new();

        private readonly MsgInfosAndSummary mFileErrors = new();
        private readonly MsgInfosAndSummary mFileWarnings = new();

        private readonly List<RuleDefinition> mHeaderLineRules = new();
        private readonly List<RuleDefinition> mProteinNameRules = new();
        private readonly List<RuleDefinition> mProteinDescriptionRules = new();
        private readonly List<RuleDefinition> mProteinSequenceRules = new();
        private int mMasterCustomRuleID = CUSTOM_RULE_ID_START;

        private char[] mProteinNameFirstRefSepChars;
        private char[] mProteinNameSubsequentRefSepChars;

        /// <summary>
        /// This array has a space and a non-breaking space
        /// </summary>
        private readonly char[] mProteinAccessionSepChars = { ' ', '\x00a0' };

        private bool mAddMissingLinefeedAtEOF;
        private bool mCheckForDuplicateProteinNames;

        /// <summary>
        /// Check for duplicate protein sequences
        /// </summary>
        /// <remarks>
        /// This will be set to True if mSaveProteinSequenceHashInfoFiles is True
        /// or mFixedFastaOptions.ConsolidateProteinsWithDuplicateSeqs is True
        /// </remarks>
        private bool mCheckForDuplicateProteinSequences;

        /// <summary>
        /// Maximum number of errors per type to track
        /// </summary>
        private int mMaximumFileErrorsToTrack;
        private int mMinimumProteinNameLength;
        private int mMaximumProteinNameLength;
        private int mMaximumResiduesPerLine;

        /// <summary>
        /// Options used when mGenerateFixedFastaFile is True
        /// </summary>
        private readonly FixedFastaOptions mFixedFastaOptions = new();

        private bool mOutputToStatsFile;
        private string mStatsFilePath;

        private bool mGenerateFixedFastaFile;
        private bool mSaveProteinSequenceHashInfoFiles;

        /// <summary>
        /// When true, create a text file that will contain the protein name and sequence hash for each protein.
        /// This option will not store protein names and/or hashes in memory, and is thus useful for processing
        /// huge .Fasta files to determine duplicate proteins.
        /// </summary>
        private bool mSaveBasicProteinHashInfoFile;

        private bool mAllowAsteriskInResidues;
        private bool mAllowDashInResidues;
        private bool mAllowAllSymbolsInProteinNames;

        private bool mWarnBlankLinesBetweenProteins;
        private bool mWarnLineStartsWithSpace;
        private bool mNormalizeFileLineEndCharacters;

        /// <summary>
        /// The number of characters at the start of key strings to use when adding items to NestedStringDictionary instances
        /// </summary>
        /// <remarks>
        /// If this value is too short, all of the items added to the NestedStringDictionary instance
        /// will be tracked by the same dictionary, which could result in a dictionary surpassing the 2 GB boundary
        /// </remarks>
        private byte mProteinNameSpannerCharLength = 1;

        private ValidateFastaFileErrorCodes mLocalErrorCode;

        private MemoryUsageLogger mMemoryUsageLogger;

        private float mProcessMemoryUsageMBAtStart;

        private string mSortUtilityErrorMessage;

        private List<string> mTempFilesToDelete;

        /// <summary>
        /// Gets a processing option
        /// </summary>
        /// <param name="switchName"></param>
        /// <remarks>Be sure to call SetDefaultRules() after setting all of the options</remarks>
        [Obsolete("Use GetOptionSwitchValue instead", true)]
        // ReSharper disable once UnusedMember.Global
        public bool get_OptionSwitch(SwitchOptions switchName)
        {
            return GetOptionSwitchValue(switchName);
        }

        /// <summary>
        /// Sets a processing option
        /// </summary>
        /// <param name="switchName"></param>
        /// <param name="value"></param>
        [Obsolete("Use SetOptionSwitch instead", true)]
        // ReSharper disable once UnusedMember.Global
        public void set_OptionSwitch(SwitchOptions switchName, bool value)
        {
            SetOptionSwitch(switchName, value);
        }

        /// <summary>
        /// Set a processing option
        /// </summary>
        /// <param name="switchName"></param>
        /// <param name="state"></param>
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

        /// <summary>
        /// Get a processing option
        /// </summary>
        /// <param name="switchName"></param>
        public bool GetOptionSwitchValue(SwitchOptions switchName)
        {
            return switchName switch
            {
                SwitchOptions.AddMissingLineFeedAtEOF => mAddMissingLinefeedAtEOF,
                SwitchOptions.AllowAsteriskInResidues => mAllowAsteriskInResidues,
                SwitchOptions.CheckForDuplicateProteinNames => mCheckForDuplicateProteinNames,
                SwitchOptions.GenerateFixedFASTAFile => mGenerateFixedFastaFile,
                SwitchOptions.OutputToStatsFile => mOutputToStatsFile,
                SwitchOptions.SplitOutMultipleRefsInProteinName => mFixedFastaOptions.SplitOutMultipleRefsInProteinName,
                SwitchOptions.WarnBlankLinesBetweenProteins => mWarnBlankLinesBetweenProteins,
                SwitchOptions.WarnLineStartsWithSpace => mWarnLineStartsWithSpace,
                SwitchOptions.NormalizeFileLineEndCharacters => mNormalizeFileLineEndCharacters,
                SwitchOptions.CheckForDuplicateProteinSequences => mCheckForDuplicateProteinSequences,
                SwitchOptions.FixedFastaRenameDuplicateNameProteins => mFixedFastaOptions.RenameProteinsWithDuplicateNames,
                SwitchOptions.FixedFastaKeepDuplicateNamedProteins => mFixedFastaOptions.KeepDuplicateNamedProteinsUnlessMatchingSequence,
                SwitchOptions.SaveProteinSequenceHashInfoFiles => mSaveProteinSequenceHashInfoFiles,
                SwitchOptions.FixedFastaConsolidateDuplicateProteinSeqs => mFixedFastaOptions.ConsolidateProteinsWithDuplicateSeqs,
                SwitchOptions.FixedFastaConsolidateDupsIgnoreILDiff => mFixedFastaOptions.ConsolidateDupsIgnoreILDiff,
                SwitchOptions.FixedFastaTruncateLongProteinNames => mFixedFastaOptions.TruncateLongProteinNames,
                SwitchOptions.FixedFastaSplitOutMultipleRefsForKnownAccession => mFixedFastaOptions.SplitOutMultipleRefsForKnownAccession,
                SwitchOptions.FixedFastaWrapLongResidueLines => mFixedFastaOptions.WrapLongResidueLines,
                SwitchOptions.FixedFastaRemoveInvalidResidues => mFixedFastaOptions.RemoveInvalidResidues,
                SwitchOptions.SaveBasicProteinHashInfoFile => mSaveBasicProteinHashInfoFile,
                SwitchOptions.AllowDashInResidues => mAllowDashInResidues,
                SwitchOptions.AllowAllSymbolsInProteinNames => mAllowAllSymbolsInProteinNames,
                _ => false,
            };
        }

        /// <summary>
        /// Get error warning counts for the given message type and count type
        /// </summary>
        /// <param name="messageType"></param>
        /// <param name="countType"></param>
        [Obsolete("Use GetErrorWarningCounts instead", true)]
        // ReSharper disable once UnusedMember.Global
        public int get_ErrorWarningCounts(
            MsgTypeConstants messageType,
            ErrorWarningCountTypes countType)
        {
            return GetErrorWarningCounts(messageType, countType);
        }

        /// <summary>
        /// Get error warning counts for the given message type and count type
        /// </summary>
        /// <param name="messageType"></param>
        /// <param name="countType"></param>
        public int GetErrorWarningCounts(
            MsgTypeConstants messageType,
            ErrorWarningCountTypes countType)
        {
            MsgInfosAndSummary msgSet;
            switch (messageType)
            {
                case MsgTypeConstants.ErrorMsg:
                    msgSet = mFileErrors;
                    break;

                case MsgTypeConstants.WarningMsg:
                    msgSet = mFileWarnings;
                    break;

                case MsgTypeConstants.StatusMsg:
                    return 0;

                default:
                    return 0;
            }

            var count = 0;
            if (countType == ErrorWarningCountTypes.Specified || countType == ErrorWarningCountTypes.Total)
            {
                count += msgSet.Messages.Count;
            }

            if (countType == ErrorWarningCountTypes.Unspecified || countType == ErrorWarningCountTypes.Total)
            {
                count += msgSet.ComputeTotalUnspecifiedCount();
            }

            return count;
        }

        /// <summary>
        /// Get fixed FASTA file stats for the given stat category
        /// </summary>
        /// <param name="statCategory"></param>
        [Obsolete("Use GetFixedFASTAFileStats instead", true)]
        // ReSharper disable once UnusedMember.Global
        public int get_FixedFASTAFileStats(FixedFASTAFileValues statCategory)
        {
            return GetFixedFASTAFileStats(statCategory);
        }

        /// <summary>
        /// Get fixed FASTA file stats for the given stat category
        /// </summary>
        /// <param name="statCategory"></param>
        public int GetFixedFASTAFileStats(FixedFASTAFileValues statCategory)
        {
            return mFixedFastaStats.GetStat(statCategory);
        }

        /// <summary>
        /// Number of proteins in the FASTA file
        /// </summary>
        public int ProteinCount { get; private set; }

        /// <summary>
        /// Number of lines in the FASTA file
        /// </summary>
        public int LineCount { get; private set; }

        /// <summary>
        /// Local error code
        /// </summary>
        // ReSharper disable once UnusedMember.Global
        public ValidateFastaFileErrorCodes LocalErrorCode => mLocalErrorCode;

        /// <summary>
        /// Number of residues in the FASTA file
        /// </summary>
        public long ResidueCount { get; private set; }

        /// <summary>
        /// FASTA file path
        /// </summary>
        // ReSharper disable once UnusedMember.Global
        public string FastaFilePath => mFastaFilePath;

        /// <summary>
        /// Get error message text by index in mErrors
        /// </summary>
        /// <param name="index"></param>
        /// <param name="valueSeparator"></param>
        [Obsolete("Use GetErrorMessageTextByIndex", true)]
        // ReSharper disable once UnusedMember.Global
        public string get_ErrorMessageTextByIndex(int index, string valueSeparator)
        {
            return GetErrorMessageTextByIndex(index, valueSeparator);
        }

        /// <summary>
        /// Get error message text by index in mErrors
        /// </summary>
        /// <param name="index"></param>
        /// <param name="valueSeparator"></param>
        public string GetErrorMessageTextByIndex(int index, string valueSeparator)
        {
            return GetFileErrorTextByIndex(index, valueSeparator);
        }

        /// <summary>
        /// Get warning message text by index in mFileWarnings
        /// </summary>
        /// <param name="index"></param>
        /// <param name="valueSeparator"></param>
        [Obsolete("Use GetWarningMessageTextByIndex", true)]
        // ReSharper disable once UnusedMember.Global
        public string get_WarningMessageTextByIndex(int index, string valueSeparator)
        {
            return GetWarningMessageTextByIndex(index, valueSeparator);
        }

        /// <summary>
        /// Get warning message text by index in mFileWarnings
        /// </summary>
        /// <param name="index"></param>
        /// <param name="valueSeparator"></param>
        public string GetWarningMessageTextByIndex(int index, string valueSeparator)
        {
            return GetFileWarningTextByIndex(index, valueSeparator);
        }

        /// <summary>
        /// Get error details, by index in mErrors
        /// </summary>
        /// <param name="errorIndex"></param>
        [Obsolete("Use GetErrorsByIndex", true)]
        // ReSharper disable once UnusedMember.Global
        public MsgInfo get_ErrorsByIndex(int errorIndex)
        {
            return GetErrorsByIndex(errorIndex);
        }

        /// <summary>
        /// Get error details, by index in mErrors
        /// </summary>
        /// <param name="errorIndex"></param>
        public MsgInfo GetErrorsByIndex(int errorIndex)
        {
            return GetFileErrorByIndex(errorIndex);
        }

        /// <summary>
        /// Get warning details, by index in mWarnings
        /// </summary>
        /// <param name="warningIndex"></param>
        [Obsolete("Use GetWarningsByIndex", true)]
        // ReSharper disable once UnusedMember.Global
        public MsgInfo get_WarningsByIndex(int warningIndex)
        {
            return GetWarningsByIndex(warningIndex);
        }

        /// <summary>
        /// Get warning details, by index in mWarnings
        /// </summary>
        /// <param name="warningIndex"></param>
        public MsgInfo GetWarningsByIndex(int warningIndex)
        {
            return GetFileWarningByIndex(warningIndex);
        }

        /// <summary>
        /// Existing protein hash file to load into memory instead of computing new hash values while reading the FASTA file
        /// </summary>
        public string ExistingProteinHashFile { get; set; }

        /// <summary>
        /// Maximum number of errors to track
        /// </summary>
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

        /// <summary>
        /// Maximum protein name length
        /// </summary>
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

        /// <summary>
        /// Minimum protein name length
        /// </summary>
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

        /// <summary>
        /// Maximum residues per line
        /// </summary>
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

        /// <summary>
        /// Protein line start character
        /// </summary>
        public char ProteinLineStartChar { get; set; }

        /// <summary>
        /// Stats file path
        /// </summary>
        // ReSharper disable once UnusedMember.Global
        public string StatsFilePath => mStatsFilePath ?? string.Empty;

        /// <summary>
        /// Invalid characters to remove from protein names
        /// </summary>
        public string ProteinNameInvalidCharsToRemove
        {
            get => new(mFixedFastaOptions.ProteinNameInvalidCharsToRemove);
            set
            {
                // Check for and remove any spaces from Value, since
                // a space does not make sense for an invalid protein name character
                value = value?.Replace(" ", string.Empty) ?? string.Empty;
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

        /// <summary>
        /// Separation characters used to find the first reference in a protein name
        /// </summary>
        public string ProteinNameFirstRefSepChars
        {
            get => new(mProteinNameFirstRefSepChars);
            set
            {
                // Check for and remove any spaces from Value, since
                // a space does not make sense for a separation character
                value = value?.Replace(" ", string.Empty) ?? string.Empty;
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

        /// <summary>
        /// Separation characters used to find the additional references in a protein name
        /// </summary>
        public string ProteinNameSubsequentRefSepChars
        {
            get => new(mProteinNameSubsequentRefSepChars);
            set
            {
                // Check for and remove any spaces from Value, since
                // a space does not make sense for a separation character
                value = value?.Replace(" ", string.Empty) ?? string.Empty;
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

        /// <summary>
        /// Long protein name split characters
        /// </summary>
        public string LongProteinNameSplitChars
        {
            get => new(mFixedFastaOptions.LongProteinNameSplitChars);
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

        /// <summary>
        /// List of warnings
        /// </summary>
        // ReSharper disable once UnusedMember.Global
        public List<MsgInfo> FileWarningList => GetFileWarnings();

        /// <summary>
        /// List of errors
        /// </summary>
        // ReSharper disable once UnusedMember.Global
        public List<MsgInfo> FileErrorList => GetFileErrors();

        #region "Events and delegates"

        /// <summary>
        /// Progress completed event
        /// </summary>
        public event ProgressCompletedEventHandler ProgressCompleted;

        /// <summary>
        /// Progress completed event handler
        /// </summary>
        public delegate void ProgressCompletedEventHandler();

        /// <summary>
        /// Event raised when a new FASTA file is created, with normalized line endings
        /// </summary>
        public event WroteLineEndNormalizedFASTAEventHandler WroteLineEndNormalizedFASTA;

        /// <summary>
        /// Event handler for event WroteLineEndNormalizedFASTA
        /// </summary>
        /// <param name="newFilePath"></param>
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

        #endregion

        /// <summary>
        /// Examine the given FASTA file to look for problems.
        /// Optionally create a new, fixed FASTA file
        /// Optionally also consolidate proteins with duplicate sequences
        /// </summary>
        /// <param name="fastaFilePathToCheck"></param>
        /// <param name="preloadedProteinNamesToKeep">
        /// Preloaded list of protein names to include in the fixed FASTA file
        /// Keys are protein names, values are the number of entries written to the fixed FASTA file for the given protein name
        /// </param>
        /// <returns>True if the file was successfully analyzed (even if errors were found)</returns>
        /// <remarks>Assumes fastaFilePathToCheck exists</remarks>
        private bool AnalyzeFastaFile(string fastaFilePathToCheck, NestedStringIntList preloadedProteinNamesToKeep)
        {
            StreamWriter fixedFastaWriter = null;
            StreamWriter sequenceHashWriter = null;

            var fastaFilePathOut = "UndefinedFilePath.xyz";

            bool success;
            var exceptionCaught = false;

            var consolidateDuplicateProteinSeqsInFasta = false;
            var keepDuplicateNamedProteinsUnlessMatchingSequence = false;
            var consolidateDupsIgnoreILDiff = false;

            try
            {
                // Reset the data structures and variables
                ResetStructures();
                var proteinSeqHashInfo = new List<ProteinHashInfo>(1);

                // This is a dictionary of dictionaries, with one dictionary for each letter or number that a SHA-1 hash could start with
                // This dictionary of dictionaries provides a quick lookup for existing protein hashes
                // This dictionary is not used if preloadedProteinNamesToKeep contains data
                const int SPANNER_CHAR_LENGTH = 1;
                var proteinSequenceHashes = new NestedStringDictionary<int>(false, SPANNER_CHAR_LENGTH);

                var usingPreloadedProteinNames = false;

                if (preloadedProteinNamesToKeep?.Count > 0)
                {
                    // Auto enable/disable some options
                    mSaveBasicProteinHashInfoFile = false;
                    mCheckForDuplicateProteinSequences = false;

                    // Auto-enable creating a fixed FASTA file
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
                        LineEndingCharacters.CRLF);

                    if (!string.IsNullOrWhiteSpace(mFastaFilePath) &&
                        !string.IsNullOrWhiteSpace(fastaFilePathToCheck) &&
                        !mFastaFilePath.Equals(fastaFilePathToCheck))
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

                var proteinHeaderFound = false;
                var processingResidueBlock = false;
                var blankLineProcessed = false;

                var proteinName = string.Empty;
                var currentResidues = new StringBuilder();

                // Initialize the RegEx objects
                // This object contains multiple RegEx, with the following capabilities:
                // Note that each of these RegEx tests contain two groups with captured text:
                // reMatchIPI extracts IPI:IPI00048500.11 from IPI:IPI00048500.11|ref|23848934
                // reMatchGI extracts gi|169602219 from gi|169602219|ref|XP_001794531.1|
                // reMatchJGI extracts jgi|Batde5|906240 from jgi|Batde5|90624|GP3.061830
                // reMatchGeneric extracts bob|234384 from  bob|234384|ref|483293
                //                         or bob|845832 from  bob|845832;ref|384923
                // reMatchJGIBaseAndID matches jgi|Batde5|23435 ; it requires that there be a number after the second bar
                // reMatchDoubleBarOrColonAndBar contains a group with captured text
                var reProteinNameTruncation = new ProteinNameTruncationRegex(mProteinNameFirstRefSepChars, mProteinNameSubsequentRefSepChars);

                // Non-letter characters in residues
                var allowedResidueChars = "A-Z";
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
                    mFixedFastaOptions.LongProteinNameSplitChars = new[] { DEFAULT_LONG_PROTEIN_NAME_SPLIT_CHAR };
                }

                // Initialize the rule details UDTs, which contain a RegEx object for each rule
                InitializeRuleDetails(mHeaderLineRules, out var headerLineRuleDetails);
                InitializeRuleDetails(mProteinNameRules, out var proteinNameRuleDetails);
                InitializeRuleDetails(mProteinDescriptionRules, out var proteinDescriptionRuleDetails);
                InitializeRuleDetails(mProteinSequenceRules, out var proteinSequenceRuleDetails);

                // Open the file and read, at most, the first 100,000 characters to see if it contains CrLf or just Lf
                var terminatorSize = DetermineLineTerminatorSize(fastaFilePathToCheck);

                // Pre-scan a portion of the FASTA file to determine the appropriate value for mProteinNameSpannerCharLength
                AutoDetermineFastaProteinNameSpannerCharLength(mFastaFilePath, terminatorSize);

                // Open the input file
                var startTime = DateTime.UtcNow;

                using (var fastaReader = new StreamReader(new FileStream(fastaFilePathToCheck, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)))
                {
                    // Optionally, create the output FASTA file
                    if (mGenerateFixedFastaFile)
                    {
                        try
                        {
                            var outputDirectory = GetOutputDirectory(Path.GetDirectoryName(fastaFilePathToCheck) ?? string.Empty);

                            fastaFilePathOut = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(fastaFilePathToCheck) + "_new.fasta");

                            fixedFastaWriter = new StreamWriter(new FileStream(fastaFilePathOut, FileMode.Create, FileAccess.Write, FileShare.ReadWrite));
                        }
                        catch (Exception ex)
                        {
                            // Error opening output file
                            RecordFastaFileError(0, 0, string.Empty, (int)MessageCodeConstants.UnspecifiedError,
                                "Error creating output file " + fastaFilePathOut + ": " + ex.Message, string.Empty);
                            OnErrorEvent("Error creating output file (Create _new.fasta)", ex);
                            return false;
                        }
                    }

                    // Optionally, open the Sequence Hash file
                    if (mSaveBasicProteinHashInfoFile)
                    {
                        var basicProteinHashInfoFilePath = "<undefined>";

                        try
                        {
                            var outputDirectory = GetOutputDirectory(Path.GetDirectoryName(fastaFilePathToCheck) ?? string.Empty);

                            basicProteinHashInfoFilePath = Path.Combine(
                                outputDirectory,
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
                            RecordFastaFileError(0, 0, string.Empty, (int)MessageCodeConstants.UnspecifiedError,
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
                        proteinSeqHashInfo.Clear();
                        proteinSeqHashInfo.Capacity = 100;
                    }

                    // Parse each line in the file
                    long bytesRead = 0;
                    var lastMemoryUsageReport = DateTime.UtcNow;

                    // Note: This value is updated only if the line length is < mMaximumResiduesPerLine
                    var currentValidResidueLineLengthMax = 0;
                    var processingDuplicateOrInvalidProtein = false;

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
                                "indicating a corrupt FASTA file", LineCount + 1), ex);
                            exceptionCaught = true;
                            break;
                        }
                        catch (Exception ex)
                        {
                            OnErrorEvent(string.Format("Error in AnalyzeFastaFile reading line {0}", LineCount + 1), ex);
                            exceptionCaught = true;
                            break;
                        }

                        bytesRead += lineIn.Length + terminatorSize;

                        if (LineCount % 250 == 0)
                        {
                            if (AbortProcessing)
                                break;

                            if (DateTime.UtcNow.Subtract(lastProgressReport).TotalSeconds >= 0.5)
                            {
                                lastProgressReport = DateTime.UtcNow;
                                var percentComplete = (float)(bytesRead / (double)fastaReader.BaseStream.Length * 100.0);
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

                        LineCount++;

                        if (lineIn.Length > 10000000)
                        {
                            RecordFastaFileError(LineCount, 0, proteinName, (int)MessageCodeConstants.ResiduesLineTooLong, "Line is over 10 million residues long; skipping", string.Empty);
                            continue;
                        }

                        if (lineIn.Length > 1000000)
                        {
                            RecordFastaFileWarning(LineCount, 0, proteinName, (int)MessageCodeConstants.ResiduesLineTooLong, "Line is over 1 million residues long; this is very suspicious", string.Empty);
                        }
                        else if (lineIn.Length > 100000)
                        {
                            RecordFastaFileWarning(LineCount, 0, proteinName, (int)MessageCodeConstants.ResiduesLineTooLong, "Line is over 1 million residues long; this could indicate a problem", string.Empty);
                        }

                        if (lineIn == null)
                            continue;

                        if (lineIn.Trim().Length == 0)
                        {
                            // We typically only want blank lines at the end of the FASTA file or between two protein entries
                            blankLineProcessed = true;
                            continue;
                        }

                        if (lineIn[0] == ' ')
                        {
                            if (mWarnLineStartsWithSpace)
                            {
                                RecordFastaFileError(LineCount, 0, string.Empty,
                                                     (int)MessageCodeConstants.LineStartsWithSpace, string.Empty, ExtractContext(lineIn, 0));
                            }
                        }

                        // Note: Only trim the start of the line; do not trim the end of the line since SEQUEST incorrectly notates the peptide terminal state if a residue has a space after it
                        lineIn = lineIn.TrimStart();

                        if (lineIn[0] == ProteinLineStartChar)
                        {
                            // Protein entry

                            if (currentResidues.Length > 0)
                            {
                                ProcessResiduesForPreviousProtein(
                                    proteinName, currentResidues,
                                    proteinSequenceHashes,
                                    proteinSeqHashInfo,
                                    consolidateDupsIgnoreILDiff,
                                    fixedFastaWriter,
                                    currentValidResidueLineLengthMax,
                                    sequenceHashWriter);

                                currentValidResidueLineLengthMax = 0;
                            }

                            // Now process this protein entry
                            ProteinCount++;
                            proteinHeaderFound = true;
                            processingResidueBlock = false;

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
                                    RecordFastaFileWarning(LineCount, 0, proteinName, (int)MessageCodeConstants.BlankLineBeforeProteinName);
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
                                        RecordFastaFileError(LineCount, 0, proteinName, (int)MessageCodeConstants.BlankLineBetweenProteinNameAndResidues);
                                    }
                                }
                                else
                                {
                                    RecordFastaFileError(LineCount, 0, string.Empty, (int)MessageCodeConstants.ResiduesFoundWithoutProteinHeader);
                                }

                                processingResidueBlock = true;
                            }
                            else if (blankLineProcessed)
                            {
                                RecordFastaFileError(LineCount, 0, proteinName, (int)MessageCodeConstants.BlankLineInMiddleOfResidues);
                            }

                            var newResidueCount = lineIn.Length;
                            ResidueCount += newResidueCount;

                            // Check the line length; raise a warning if longer than suggested
                            if (newResidueCount > mMaximumResiduesPerLine)
                            {
                                RecordFastaFileWarning(LineCount, 0, proteinName, (int)MessageCodeConstants.ResiduesLineTooLong, newResidueCount.ToString(), string.Empty);
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
                                        mFixedFastaStats.UpdatedResidueLines++;
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
                                        currentResidues.Append(residuesClean);
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

                    if (currentResidues.Length > 0)
                    {
                        ProcessResiduesForPreviousProtein(
                            proteinName, currentResidues,
                            proteinSequenceHashes,
                            proteinSeqHashInfo,
                            consolidateDupsIgnoreILDiff,
                            fixedFastaWriter, currentValidResidueLineLengthMax,
                            sequenceHashWriter);
                    }

                    if (mCheckForDuplicateProteinSequences)
                    {
                        // Step through proteinSeqHashInfo and look for duplicate sequences
                        foreach (var hashItem in proteinSeqHashInfo)
                        {
                            if (hashItem.AdditionalProteins.Any())
                            {
                                var proteinHashInfo = hashItem;
                                RecordFastaFileWarning(LineCount, 0, proteinHashInfo.ProteinNameFirst, (int)MessageCodeConstants.DuplicateProteinSequence,
                                    proteinHashInfo.ProteinNameFirst + ", " + FlattenArray(proteinHashInfo.AdditionalProteins, ','), proteinHashInfo.SequenceStart);
                            }
                        }
                    }

                    var memoryUsageMB = MemoryUsageLogger.GetProcessMemoryUsageMB();
                    if (memoryUsageMB > mProcessMemoryUsageMBAtStart * 4 ||
                        memoryUsageMB - mProcessMemoryUsageMBAtStart > 50)
                    {
                        ReportMemoryUsage(preloadedProteinNamesToKeep, proteinSequenceHashes, proteinNames, proteinSeqHashInfo);
                    }
                }

                var totalSeconds = DateTime.UtcNow.Subtract(startTime).TotalSeconds;
                if (totalSeconds > 5)
                {
                    var linesPerSecond = (int)Math.Round(LineCount / totalSeconds);
                    Console.WriteLine();
                    ShowMessage(string.Format(
                        "Processing complete after {0:N0} seconds; read {1:N0} lines/second",
                        totalSeconds, linesPerSecond));
                }

                // Close the output files
                fixedFastaWriter?.Close();
                sequenceHashWriter?.Close();

                if (ProteinCount == 0)
                {
                    RecordFastaFileError(LineCount, 0, string.Empty, (int)MessageCodeConstants.ProteinEntriesNotFound);
                }
                else if (proteinHeaderFound)
                {
                    RecordFastaFileError(LineCount, 0, proteinName, (int)MessageCodeConstants.FinalProteinEntryMissingResidues);
                }
                else if (!blankLineProcessed)
                {
                    // File does not end in multiple blank lines; need to re-open it using a binary reader and check the last two characters to make sure they're valid
                    System.Threading.Thread.Sleep(100);

                    if (!VerifyLinefeedAtEOF(fastaFilePathToCheck, mAddMissingLinefeedAtEOF))
                    {
                        RecordFastaFileError(LineCount, 0, string.Empty, (int)MessageCodeConstants.FileDoesNotEndWithLinefeed);
                    }
                }

                if (usingPreloadedProteinNames)
                {
                    // Report stats on the number of proteins read, the number written, and any that had duplicate protein names in the original FASTA file
                    var nameCountNotFound = 0;
                    var duplicateProteinNameCount = 0;
                    var proteinCountWritten = 0;
                    var preloadedProteinNameCount = 0;

                    foreach (var spanningKey in preloadedProteinNamesToKeep.GetSpanningKeys())
                    {
                        var proteinsForKey = preloadedProteinNamesToKeep.GetListForSpanningKey(spanningKey);
                        preloadedProteinNameCount += proteinsForKey.Count;

                        foreach (var proteinEntry in proteinsForKey)
                        {
                            if (proteinEntry.Value == 0)
                            {
                                nameCountNotFound++;
                            }
                            else
                            {
                                proteinCountWritten++;
                                if (proteinEntry.Value > 1)
                                {
                                    duplicateProteinNameCount++;
                                }
                            }
                        }
                    }

                    Console.WriteLine();
                    if (proteinCountWritten == preloadedProteinNameCount)
                    {
                        ShowMessage("Fixed FASTA has all " + proteinCountWritten.ToString("#,##0") + " proteins determined from the pre-existing protein hash file");
                    }
                    else
                    {
                        ShowMessage("Fixed FASTA has " + proteinCountWritten.ToString("#,##0") + " of the " + preloadedProteinNameCount.ToString("#,##0") + " proteins determined from the pre-existing protein hash file");
                    }

                    if (nameCountNotFound > 0)
                    {
                        ShowMessage("WARNING: " + nameCountNotFound.ToString("#,##0") + " protein names were in the protein name list to keep, but were not found in the FASTA file");
                    }

                    if (duplicateProteinNameCount > 0)
                    {
                        ShowMessage("WARNING: " + duplicateProteinNameCount.ToString("#,##0") + " protein names were present multiple times in the FASTA file; duplicate entries were skipped");
                    }

                    success = !exceptionCaught;
                }
                else if (mSaveProteinSequenceHashInfoFiles)
                {
                    var percentComplete = 98.0F;
                    if (consolidateDuplicateProteinSeqsInFasta || keepDuplicateNamedProteinsUnlessMatchingSequence)
                    {
                        percentComplete = percentComplete * 3 / 4;
                    }

                    UpdateProgress("Validating FASTA File (" + Math.Round(percentComplete, 0) + "% Done)", percentComplete);

                    var hashInfoSuccess = AnalyzeFastaSaveHashInfo(
                        fastaFilePathToCheck,
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

                    OnProgressComplete();
                }
            }
            catch (Exception ex)
            {
                OnErrorEvent(string.Format("Error in AnalyzeFastaFile reading line {0}", LineCount), ex);
                success = false;
            }
            finally
            {
                // These close statements will typically be redundant,
                // However, if an exception occurs, they will be needed to close the files

                fixedFastaWriter?.Close();

                // ReSharper disable once ConstantConditionalAccessQualifier
                sequenceHashWriter?.Close();
            }

            return success;
        }

        private void AnalyzeFastaProcessProteinHeader(
            TextWriter fixedFastaWriter,
            string lineIn,
            out string proteinName,
            out bool processingDuplicateOrInvalidProtein,
            NestedStringIntList preloadedProteinNamesToKeep,
            ISet<string> proteinNames,
            IList<RuleDefinitionExtended> headerLineRuleDetails,
            IList<RuleDefinitionExtended> proteinNameRuleDetails,
            IList<RuleDefinitionExtended> proteinDescriptionRuleDetails,
            ProteinNameTruncationRegex reProteinNameTruncation)
        {
            var skipDuplicateProtein = false;

            proteinName = string.Empty;
            processingDuplicateOrInvalidProtein = true;

            try
            {
                SplitFastaProteinHeaderLine(lineIn, out proteinName, out var proteinDescription, out var descriptionStartIndex);

                processingDuplicateOrInvalidProtein = proteinName.Length == 0;

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
                        RecordFastaFileWarning(LineCount, 1, proteinName,
                            (int)MessageCodeConstants.ProteinNameIsTooShort, proteinName.Length.ToString(), string.Empty);
                    }
                    else if (proteinName.Length > mMaximumProteinNameLength)
                    {
                        RecordFastaFileError(LineCount, 1, proteinName,
                            (int)MessageCodeConstants.ProteinNameIsTooLong, proteinName.Length.ToString(), string.Empty);
                    }

                    // Test the protein name rules
                    EvaluateRules(proteinNameRuleDetails, proteinName, proteinName, 1, lineIn, DEFAULT_CONTEXT_LENGTH);

                    if (preloadedProteinNamesToKeep?.Count > 0)
                    {
                        // See if preloadedProteinNamesToKeep contains proteinName
                        var matchCount = preloadedProteinNamesToKeep.GetValueForItem(proteinName, -1);

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
                                proteinDescription = proteinDescription.TrimStart('|', ' ');
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
                            proteinName = ExamineProteinName(ref proteinName, proteinNames, out skipDuplicateProtein, out processingDuplicateOrInvalidProtein);

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
                RecordFastaFileError(0, 0, string.Empty, (int)MessageCodeConstants.UnspecifiedError,
                    "Error parsing protein header line '" + lineIn + "': " + ex.Message, string.Empty);
                OnErrorEvent("Error parsing protein header line", ex);
            }
        }

        private bool AnalyzeFastaSaveHashInfo(
            string fastaFilePathToCheck,
            IList<ProteinHashInfo> proteinSeqHashInfo,
            bool consolidateDuplicateProteinSeqsInFasta,
            bool consolidateDupsIgnoreILDiff,
            bool keepDuplicateNamedProteinsUnlessMatchingSequence,
            string fastaFilePathOut)
        {
            StreamWriter uniqueProteinSeqsWriter;
            StreamWriter duplicateProteinMapWriter = null;

            var uniqueProteinSeqsFileOut = string.Empty;
            var duplicateProteinMappingFileOut = string.Empty;

            bool success;

            var duplicateProteinSeqsFound = false;

            try
            {
                var outputDirectory = GetOutputDirectory(Path.GetDirectoryName(fastaFilePathToCheck) ?? string.Empty);

                uniqueProteinSeqsFileOut = Path.Combine(
                    outputDirectory,
                    Path.GetFileNameWithoutExtension(fastaFilePathToCheck) + "_UniqueProteinSeqs.txt");

                // Create uniqueProteinSeqsWriter
                uniqueProteinSeqsWriter = new StreamWriter(new FileStream(uniqueProteinSeqsFileOut, FileMode.Create, FileAccess.Write, FileShare.ReadWrite));
            }
            catch (Exception ex)
            {
                // Error opening output file
                RecordFastaFileError(0, 0, string.Empty, (int)MessageCodeConstants.UnspecifiedError,
                    "Error creating output file " + uniqueProteinSeqsFileOut + ": " + ex.Message, string.Empty);
                OnErrorEvent("Error creating output file (SaveHashInfo to _UniqueProteinSeqs.txt)", ex);
                return false;
            }

            try
            {
                // Define the path to the protein mapping file, but don't create it yet; just delete it if it exists
                // We'll only create it if two or more proteins have the same protein sequence

                var outputDirectory = GetOutputDirectory(Path.GetDirectoryName(fastaFilePathToCheck) ?? string.Empty);

                duplicateProteinMappingFileOut = Path.Combine(
                    outputDirectory,
                    Path.GetFileNameWithoutExtension(fastaFilePathToCheck) + "_UniqueProteinSeqDuplicates.txt");

                // Look for duplicateProteinMappingFileOut and erase it if it exists
                if (File.Exists(duplicateProteinMappingFileOut))
                {
                    File.Delete(duplicateProteinMappingFileOut);
                }
            }
            catch (Exception ex)
            {
                // Error deleting output file
                RecordFastaFileError(0, 0, string.Empty, (int)MessageCodeConstants.UnspecifiedError,
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

                uniqueProteinSeqsWriter.WriteLine(FlattenList(headerColumns));

                for (var index = 0; index < proteinSeqHashInfo.Count; index++)
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

                    uniqueProteinSeqsWriter.WriteLine(FlattenList(dataValues));

                    if (proteinHashInfo.AdditionalProteins.Any())
                    {
                        duplicateProteinSeqsFound = true;

                        if (duplicateProteinMapWriter == null)
                        {
                            // Need to create duplicateProteinMapWriter
                            duplicateProteinMapWriter = new StreamWriter(new FileStream(duplicateProteinMappingFileOut, FileMode.Create, FileAccess.Write, FileShare.ReadWrite));

                            var proteinHeaderColumns = new List<string>()
                            {
                                "Sequence_Index",
                                "Protein_Name_First",
                                SEQUENCE_LENGTH_COLUMN,
                                "Duplicate_Protein"
                            };

                            duplicateProteinMapWriter.WriteLine(FlattenList(proteinHeaderColumns));
                        }

                        foreach (var additionalProtein in proteinHashInfo.AdditionalProteins)
                        {
                            if (string.IsNullOrWhiteSpace(additionalProtein))
                                continue;

                            var proteinDataValues = new List<string>()
                            {
                                (index + 1).ToString(),
                                proteinHashInfo.ProteinNameFirst,
                                proteinHashInfo.SequenceLength.ToString(),
                                additionalProtein
                            };

                            duplicateProteinMapWriter.WriteLine(FlattenList(proteinDataValues));
                        }
                    }
                    else if (proteinHashInfo.DuplicateProteinNameCount > 0)
                    {
                        duplicateProteinSeqsFound = true;
                    }
                }

                uniqueProteinSeqsWriter.Close();
                duplicateProteinMapWriter?.Close();

                success = true;
            }
            catch (Exception ex)
            {
                // Error writing results
                RecordFastaFileError(0, 0, string.Empty, (int)MessageCodeConstants.UnspecifiedError,
                    "Error writing results to " + uniqueProteinSeqsFileOut + " or " + duplicateProteinMappingFileOut + ": " + ex.Message, string.Empty);
                OnErrorEvent("Error writing results to " + uniqueProteinSeqsFileOut + " or " + duplicateProteinMappingFileOut, ex);
                success = false;
            }

            if (success && proteinSeqHashInfo.Count > 0 && duplicateProteinSeqsFound)
            {
                if (consolidateDuplicateProteinSeqsInFasta || keepDuplicateNamedProteinsUnlessMatchingSequence)
                {
                    success = CorrectForDuplicateProteinSeqsInFasta(
                        consolidateDuplicateProteinSeqsInFasta,
                        consolidateDupsIgnoreILDiff,
                        fastaFilePathOut,
                        proteinSeqHashInfo);
                }
            }

            return success;
        }

        /// <summary>
        /// Pre-scan a portion of the FASTA file to determine the appropriate value for mProteinNameSpannerCharLength
        /// </summary>
        /// <param name="fastaFilePathToTest">FASTA file to examine</param>
        /// <param name="terminatorSize">Linefeed length (1 for LF or 2 for CRLF)</param>
        /// <remarks>
        /// Reads 50 MB chunks from 10 sections of the FASTA file (or the entire FASTA file if under 500 MB in size)
        /// Keeps track of the portion of protein names in common between adjacent proteins
        /// Uses this information to determine an appropriate value for mProteinNameSpannerCharLength
        /// </remarks>
        private void AutoDetermineFastaProteinNameSpannerCharLength(string fastaFilePathToTest, int terminatorSize)
        {
            const int PARTS_TO_SAMPLE = 10;
            const int KILOBYTES_PER_SAMPLE = 51200;

            var proteinStartLetters = new Dictionary<string, int>();
            var startTime = DateTime.UtcNow;
            var showStats = false;

            var fastaFile = new FileInfo(fastaFilePathToTest);
            if (!fastaFile.Exists)
                return;
            var fullScanLengthBytes = 1024L * PARTS_TO_SAMPLE * KILOBYTES_PER_SAMPLE;
            long linesReadTotal = 0;

            if (fastaFile.Length < fullScanLengthBytes)
            {
                fullScanLengthBytes = fastaFile.Length;
                linesReadTotal = AutoDetermineFastaProteinNameSpannerCharLength(fastaFile, terminatorSize, proteinStartLetters, 0, fastaFile.Length);
            }
            else
            {
                var stepSizeBytes = (long)Math.Round(fastaFile.Length / (double)PARTS_TO_SAMPLE, 0);

                for (long byteOffsetStart = 0; byteOffsetStart <= fastaFile.Length; byteOffsetStart += stepSizeBytes)
                {
                    var linesRead = AutoDetermineFastaProteinNameSpannerCharLength(fastaFile, terminatorSize, proteinStartLetters, byteOffsetStart, KILOBYTES_PER_SAMPLE * 1024);

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
                var preScanProteinCount = (from item in proteinStartLetters select item.Value).Sum();

                if (showStats)
                {
                    var percentFileProcessed = fullScanLengthBytes / (double)fastaFile.Length * 100;

                    ShowMessage(string.Format(
                        "  parsed {0:0}% of the file, reading {1:#,##0} lines and finding {2:#,##0} proteins",
                        percentFileProcessed, linesReadTotal, preScanProteinCount));
                }

                // Determine the appropriate spanner length given the observation counts of the base names
                mProteinNameSpannerCharLength = NestedStringIntList.DetermineSpannerLengthUsingStartLetterStats(proteinStartLetters);
            }

            ShowMessage("Using ProteinNameSpannerCharLength = " + mProteinNameSpannerCharLength);
            Console.WriteLine();
        }

        /// <summary>
        /// Read a portion of the FASTA file, comparing adjacent protein names and keeping track of the name portions in common
        /// </summary>
        /// <param name="fastaFile"></param>
        /// <param name="terminatorSize"></param>
        /// <param name="proteinStartLetters"></param>
        /// <param name="startOffset"></param>
        /// <param name="bytesToRead"></param>
        /// <returns>The number of lines read</returns>
        private long AutoDetermineFastaProteinNameSpannerCharLength(
            FileInfo fastaFile,
            int terminatorSize,
            IDictionary<string, int> proteinStartLetters,
            long startOffset,
            long bytesToRead)
        {
            var linesRead = 0L;

            try
            {
                var previousProteinLength = 0;
                var previousProteinName = string.Empty;

                if (startOffset >= fastaFile.Length)
                {
                    ShowMessage("Ignoring byte offset of " + startOffset +
                                " in AutoDetermineProteinNameSpannerCharLength since past the end of the file " +
                                "(" + fastaFile.Length + " bytes");
                    return 0;
                }

                long bytesRead = 0;

                var firstLineDiscarded = startOffset == 0;

                using var inStream = new FileStream(fastaFile.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                {
                    Position = startOffset
                };

                using var reader = new StreamReader(inStream);

                while (!reader.EndOfStream)
                {
                    var lineIn = reader.ReadLine();
                    bytesRead += terminatorSize;
                    linesRead++;

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

                    if (lineIn[0] != ProteinLineStartChar)
                    {
                        continue;
                    }

                    // Make sure the protein name and description are valid
                    // Find the first space and/or tab
                    var spaceIndex = GetBestSpaceIndex(lineIn);
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

                    var currentNameLength = proteinName.Length;
                    var charIndex = 0;

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

                        charIndex++;
                    }

                    var charsInCommon = charIndex;
                    if (charsInCommon > 0)
                    {
                        var baseName = previousProteinName.Substring(0, charsInCommon);

                        if (proteinStartLetters.TryGetValue(baseName, out var matchCount))
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
            ProteinNameTruncationRegex reProteinNameTruncation)
        {
            Match match;
            string newProteinName;
            int charIndex;
            string extraProteinNameText;

            var multipleRefsSplitOutFromKnownAccession = false;

            // Auto-fix potential errors in the protein name

            // Possibly truncate to mMaximumProteinNameLength characters
            var proteinNameTooLong = proteinName.Length > mMaximumProteinNameLength;

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

                match = reProteinNameTruncation.MatchIPI.Match(proteinName);
                if (match.Success)
                {
                    multipleRefsSplitOutFromKnownAccession = true;
                }
                else
                {
                    // IPI didn't match; try gi
                    match = reProteinNameTruncation.MatchGI.Match(proteinName);
                }

                if (match.Success)
                {
                    multipleRefsSplitOutFromKnownAccession = true;
                }
                else
                {
                    // GI didn't match; try jgi
                    match = reProteinNameTruncation.MatchJGI.Match(proteinName);
                }

                if (match.Success)
                {
                    multipleRefsSplitOutFromKnownAccession = true;
                }
                // jgi didn't match; try generic (text separated by a series of colons or bars),
                // but only if the name is too long
                else if (mFixedFastaOptions.TruncateLongProteinNames && proteinNameTooLong)
                {
                    match = reProteinNameTruncation.MatchGeneric.Match(proteinName);
                }

                if (match.Success)
                {
                    // Truncate the protein name, but move the truncated portion into the next group
                    newProteinName = match.Groups[1].Value;
                    extraProteinNameText = match.Groups[2].Value;
                }
                else if (mFixedFastaOptions.TruncateLongProteinNames && proteinNameTooLong)
                {
                    // Name is too long, but it didn't match the known patterns
                    // Find the last occurrence of mFixedFastaOptions.LongProteinNameSplitChars (default is vertical bar)
                    // and truncate the text following the match
                    // Repeat the process until the protein name length >= mMaximumProteinNameLength

                    // See if any of the characters in proteinNameSplitChars is present after
                    // character 6 but less than character mMaximumProteinNameLength
                    const int minCharIndex = 6;

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
                        mFixedFastaStats.TruncatedProteinNameCount++;
                    }
                    else
                    {
                        mFixedFastaStats.ProteinNamesMultipleRefsRemoved++;
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
                    {
                        // Next, replace any remaining instances of the character with an underscore
                        newProteinName = newProteinName.Replace(invalidChar, INVALID_PROTEIN_NAME_CHAR_REPLACEMENT);
                    }

                    if (proteinName != newProteinName && newProteinName.Length >= 3)
                    {
                        proteinName = string.Copy(newProteinName);
                        mFixedFastaStats.ProteinNamesInvalidCharsReplaced++;
                    }
                }
            }

            if (mFixedFastaOptions.SplitOutMultipleRefsInProteinName && !multipleRefsSplitOutFromKnownAccession)
            {
                // Look for multiple refs in the protein name, but only if we didn't already split out multiple refs above

                match = reProteinNameTruncation.MatchDoubleBarOrColonAndBar.Match(proteinName);
                if (match.Success)
                {
                    // Protein name contains 2 or more vertical bars, or a colon and a bar
                    // Split out the multiple refs and place them in the description
                    // However, jgi names are supposed to have two vertical bars, so we need to treat that data differently

                    extraProteinNameText = string.Empty;

                    match = reProteinNameTruncation.MatchJGIBaseAndID.Match(proteinName);
                    if (match.Success)
                    {
                        // ProteinName is similar to jgi|Organism|00000
                        // Check whether there is any text following the match
                        if (match.Length < proteinName.Length)
                        {
                            // Extra text exists; populate extraProteinNameText
                            extraProteinNameText = proteinName.Substring(match.Length + 1);
                            proteinName = match.ToString();
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
                        mFixedFastaStats.ProteinNamesMultipleRefsRemoved++;
                    }
                }
            }

            // Make sure proteinDescription doesn't start with a | or space
            if (proteinDescription.Length > 0)
            {
                proteinDescription = proteinDescription.TrimStart('|', ' ');
            }

            return proteinName;
        }

        private string BoolToStringInt(bool value)
        {
            return value ? "1" : "0";
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
                    ClearRulesDataStructure(mHeaderLineRules);
                    break;
                case RuleTypes.ProteinDescription:
                    ClearRulesDataStructure(mProteinDescriptionRules);
                    break;
                case RuleTypes.ProteinName:
                    ClearRulesDataStructure(mProteinNameRules);
                    break;
                case RuleTypes.ProteinSequence:
                    ClearRulesDataStructure(mProteinSequenceRules);
                    break;
            }
        }

        private void ClearRulesDataStructure(List<RuleDefinition> rules)
        {
            rules.Clear();
            rules.Capacity = 10;
        }

        /// <summary>
        /// Protein the protein hash for the residues
        /// </summary>
        /// <param name="residues"></param>
        /// <param name="consolidateDupsIgnoreILDiff"></param>
        public string ComputeProteinHash(StringBuilder residues, bool consolidateDupsIgnoreILDiff)
        {
            if (residues.Length > 0)
            {
                // ReSharper disable once ConvertIfStatementToReturnStatement
                if (consolidateDupsIgnoreILDiff)
                {
                    return HashUtilities.ComputeStringHashSha1(residues.ToString().Replace('L', 'I')).ToUpper();
                }
                else
                {
                    return HashUtilities.ComputeStringHashSha1(residues.ToString()).ToUpper();
                }
            }

            return string.Empty;
        }

        /// <summary>
        /// Looks for duplicate proteins in the FASTA file
        /// Creates a new FASTA file that has exact duplicates removed
        /// Will consolidate proteins with the same sequence if consolidateDuplicateProteinSeqsInFasta=True
        /// </summary>
        /// <param name="consolidateDuplicateProteinSeqsInFasta"></param>
        /// <param name="consolidateDupsIgnoreILDiff"></param>
        /// <param name="fixedFastaFilePath"></param>
        /// <param name="proteinSeqHashInfo"></param>
        private bool CorrectForDuplicateProteinSeqsInFasta(
            bool consolidateDuplicateProteinSeqsInFasta,
            bool consolidateDupsIgnoreILDiff,
            string fixedFastaFilePath,
            IList<ProteinHashInfo> proteinSeqHashInfo)
        {
            StreamWriter consolidatedFastaWriter = null;

            int terminatorSize;

            var fixedFastaFilePathTemp = string.Empty;

            var cachedProteinName = string.Empty;
            var cachedProteinDescription = string.Empty;
            var sbCachedProteinResidueLines = new StringBuilder(250);
            var sbCachedProteinResidues = new StringBuilder(250);

            bool success;

            if (proteinSeqHashInfo.Count == 0)
            {
                return true;
            }

            // '''''''''''''''''''''
            // Processing Steps
            // '''''''''''''''''''''
            //
            // Open fixedFastaFilePath with the FASTA file reader
            // Create a new FASTA file with a writer

            // For each protein, check whether it has duplicates
            // If not, just write it out to the new FASTA file

            // If it does have duplicates and it is the master, then append the duplicate protein names to the end of the description for the protein
            // and write out the name, new description, and sequence to the new FASTA file

            // Otherwise, check if it is a duplicate of a master protein
            // If it is, then do not write the name, description, or sequence to the new FASTA file

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
                RecordFastaFileError(0, 0, string.Empty, (int)MessageCodeConstants.UnspecifiedError,
                    "Error renaming " + fixedFastaFilePath + " to " + fixedFastaFilePathTemp + ": " + ex.Message, string.Empty);
                OnErrorEvent("Error renaming fixed FASTA to .tempfixed", ex);
                return false;
            }

            StreamReader fastaReader;

            try
            {
                // Open the file and read, at most, the first 100,000 characters to see if it contains CrLf or just Lf
                terminatorSize = DetermineLineTerminatorSize(fixedFastaFilePathTemp);

                // Open the Fixed FASTA file
                Stream fsInFile = new FileStream(
                    fixedFastaFilePathTemp,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.ReadWrite);

                fastaReader = new StreamReader(fsInFile);
            }
            catch (Exception ex)
            {
                RecordFastaFileError(0, 0, string.Empty, (int)MessageCodeConstants.UnspecifiedError,
                    "Error opening " + fixedFastaFilePathTemp + ": " + ex.Message, string.Empty);
                OnErrorEvent("Error opening fixedFastaFilePathTemp", ex);
                return false;
            }

            try
            {
                // Create the new FASTA file
                consolidatedFastaWriter = new StreamWriter(new FileStream(fixedFastaFilePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite));
            }
            catch (Exception ex)
            {
                RecordFastaFileError(0, 0, string.Empty, (int)MessageCodeConstants.UnspecifiedError,
                    "Error creating consolidated FASTA output file " + fixedFastaFilePath + ": " + ex.Message, string.Empty);
                OnErrorEvent("Error creating consolidated FASTA output file", ex);
            }

            try
            {
                // This list contains the protein names that we will keep; values are the index values pointing into proteinSeqHashInfo
                // If consolidateDuplicateProteinSeqsInFasta=False, this will contain all protein names
                // If consolidateDuplicateProteinSeqsInFasta=True, we only keep the first name found for a given sequence
                // Populate proteinNameFirst with the protein names in proteinSeqHashInfo().ProteinNameFirst
                var proteinNameFirst = new NestedStringDictionary<int>(true, mProteinNameSpannerCharLength);

                // This list contains the names of duplicate proteins; the hash values are the protein names of the master protein that has the same sequence
                // Populate htDuplicateProteinList with the protein names in proteinSeqHashInfo().AdditionalProteins
                var duplicateProteinList = new NestedStringDictionary<string>(true, mProteinNameSpannerCharLength);

                for (var index = 0; index < proteinSeqHashInfo.Count; index++)
                {
                    var proteinHashInfo = proteinSeqHashInfo[index];

                    if (!proteinNameFirst.ContainsKey(proteinHashInfo.ProteinNameFirst))
                    {
                        proteinNameFirst.Add(proteinHashInfo.ProteinNameFirst, index);
                    }
                    else
                    {
                        // .ProteinNameFirst is already present in proteinNameFirst
                        // The fixed FASTA file will only actually contain the first occurrence of .ProteinNameFirst, so we can effectively ignore this entry
                        // but we should increment the DuplicateNameSkipCount

                    }

                    if (proteinHashInfo.AdditionalProteins.Any())
                    {
                        foreach (var additionalProtein in proteinHashInfo.AdditionalProteins)
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
                                // .AdditionalProteins(duplicateIndex) is already present in proteinNameFirst
                                // Increment the DuplicateNameSkipCount
                            }
                        }
                    }
                }

                // This list keeps track of the protein names that have been written out to the new FASTA file
                // Keys are the protein names; values are the index of the entry in proteinSeqHashInfo()
                var proteinsWritten = new NestedStringDictionary<int>(false, mProteinNameSpannerCharLength);

                var lastMemoryUsageReport = DateTime.UtcNow;

                // Parse each line in the file
                var lineCountRead = 0;
                long bytesRead = 0;
                mFixedFastaStats.DuplicateSequenceProteinsSkipped = 0;

                while (!fastaReader.EndOfStream)
                {
                    var lineIn = fastaReader.ReadLine();
                    bytesRead += lineIn.Length + terminatorSize;

                    if (lineCountRead % 50 == 0)
                    {
                        if (AbortProcessing)
                            break;
                        var percentComplete = 75 + (float)(bytesRead / (double)fastaReader.BaseStream.Length * 100.0) / 4;
                        UpdateProgress("Consolidating duplicate proteins to create a new FASTA File (" + Math.Round(percentComplete, 0) + "% Done)", percentComplete);

                        if (DateTime.UtcNow.Subtract(lastMemoryUsageReport).TotalMinutes >= 1)
                        {
                            lastMemoryUsageReport = DateTime.UtcNow;
                            ReportMemoryUsage(proteinNameFirst, proteinsWritten, duplicateProteinList);
                        }
                    }

                    lineCountRead++;

                    if (lineIn != null)
                    {
                        if (lineIn.Trim().Length > 0)
                        {
                            // Note: Trim the start of the line (however, since this is a fixed FASTA file it should not start with a space)
                            lineIn = lineIn.TrimStart();

                            if (lineIn[0] == ProteinLineStartChar)
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
                                SplitFastaProteinHeaderLine(lineIn, out cachedProteinName, out cachedProteinDescription, out _);
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
                RecordFastaFileError(0, 0, string.Empty, (int)MessageCodeConstants.UnspecifiedError,
                    "Error writing to consolidated FASTA file " + fixedFastaFilePath + ": " + ex.Message, string.Empty);
                OnErrorEvent("Error writing to consolidated FASTA file", ex);
                return false;
            }
            finally
            {
                try
                {
                    fastaReader?.Close();
                    consolidatedFastaWriter?.Close();

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
            if (string.IsNullOrWhiteSpace(proteinName))
            {
                proteinName = "????";
            }

            if (string.IsNullOrWhiteSpace(proteinDescription))
            {
                return ProteinLineStartChar + proteinName;
            }

            return ProteinLineStartChar + proteinName + " " + proteinDescription;
        }

        private string ConstructStatsFilePath(string outputDirectoryPath)
        {
            var outFilePath = string.Empty;

            try
            {
                // Record the current time in now
                outFilePath = "FastaFileStats_" + DateTime.Now.ToString("yyyy-MM-dd") + ".txt";

                if (!string.IsNullOrEmpty(outputDirectoryPath))
                {
                    outFilePath = Path.Combine(outputDirectoryPath, outFilePath);
                }
            }
            catch (Exception)
            {
                // Silently ignore this error
                if (string.IsNullOrWhiteSpace(outFilePath))
                {
                    outFilePath = "FastaFileStats.txt";
                }
            }

            return outFilePath;
        }

        private void DeleteTempFiles()
        {
            if (mTempFilesToDelete?.Count > 0)
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
                    catch (Exception)
                    {
                        // Ignore errors here
                    }
                }
            }
        }

        private int DetermineLineTerminatorSize(string inputFilePath)
        {
            var endCharType = DetermineLineTerminatorType(inputFilePath);

            return endCharType switch
            {
                LineEndingCharacters.CR => 1,
                LineEndingCharacters.LF => 1,
                LineEndingCharacters.CRLF => 2,
                LineEndingCharacters.LFCR => 2,
                _ => 2
            };
        }

        private LineEndingCharacters DetermineLineTerminatorType(string inputFilePath)
        {
            var endCharacterType = default(LineEndingCharacters);

            try
            {
                // Open the input file and look for the first carriage return or line feed
                using var fsInFile = new FileStream(inputFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

                while (fsInFile.Position < fsInFile.Length && fsInFile.Position < 100000)
                {
                    var oneByte = fsInFile.ReadByte();

                    if (oneByte == 10)
                    {
                        // Found linefeed
                        if (fsInFile.Position < fsInFile.Length)
                        {
                            oneByte = fsInFile.ReadByte();
                            if (oneByte == 13)
                            {
                                // LfCr
                                endCharacterType = LineEndingCharacters.LFCR;
                            }
                            else
                            {
                                // Lf only
                                endCharacterType = LineEndingCharacters.LF;
                            }
                        }
                        else
                        {
                            endCharacterType = LineEndingCharacters.LF;
                        }

                        break;
                    }

                    if (oneByte == 13)
                    {
                        // Found carriage return
                        if (fsInFile.Position < fsInFile.Length)
                        {
                            oneByte = fsInFile.ReadByte();
                            if (oneByte == 10)
                            {
                                // CrLf
                                endCharacterType = LineEndingCharacters.CRLF;
                            }
                            else
                            {
                                // Cr only
                                endCharacterType = LineEndingCharacters.CR;
                            }
                        }
                        else
                        {
                            endCharacterType = LineEndingCharacters.CR;
                        }

                        break;
                    }
                }
            }
            catch (Exception)
            {
                SetLocalErrorCode(ValidateFastaFileErrorCodes.ErrorVerifyingLinefeedAtEOF);
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
            LineEndingCharacters desiredLineEndCharacterType)
        {
            var newEndChar = "\r\n";

            var endCharType = DetermineLineTerminatorType(pathOfFileToFix);

            var origEndCharCount = 0;

            if (endCharType != desiredLineEndCharacterType)
            {
                switch (desiredLineEndCharacterType)
                {
                    case LineEndingCharacters.CRLF:
                        newEndChar = "\r\n";
                        break;
                    case LineEndingCharacters.CR:
                        newEndChar = "\r";
                        break;
                    case LineEndingCharacters.LF:
                        newEndChar = "\n";
                        break;
                    case LineEndingCharacters.LFCR:
                        newEndChar = "\r\n";
                        break;
                }

                switch (endCharType)
                {
                    case LineEndingCharacters.CR:
                        origEndCharCount = 2;
                        break;
                    case LineEndingCharacters.CRLF:
                        origEndCharCount = 1;
                        break;
                    case LineEndingCharacters.LF:
                        origEndCharCount = 1;
                        break;
                    case LineEndingCharacters.LFCR:
                        origEndCharCount = 2;
                        break;
                }

                string newFilePath;

                if (Path.IsPathRooted(newFileName))
                {
                    newFilePath = newFileName;
                }
                else
                {
                    var outputDirectory = GetOutputDirectory(Path.GetDirectoryName(pathOfFileToFix) ?? string.Empty);

                    newFilePath = Path.Combine(outputDirectory, Path.GetFileName(newFileName));
                }

                var targetFile = new FileInfo(pathOfFileToFix);
                var fileSizeBytes = targetFile.Length;

                var reader = targetFile.OpenText();

                using var writer = new StreamWriter(new FileStream(newFilePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite));

                this.OnProgressUpdate("Normalizing Line Endings...", 0.0F);

                var dataLine = reader.ReadLine();
                long linesRead = 0;
                while (dataLine != null)
                {
                    writer.Write(dataLine);
                    writer.Write(newEndChar);

                    var currentFilePos = dataLine.Length + origEndCharCount;
                    linesRead++;

                    if (linesRead % 1000 == 0)
                    {
                        var percentComplete = currentFilePos / (float)fileSizeBytes * 100;
                        var progressMessage = string.Format("Normalizing Line Endings ({0:F1} % complete", percentComplete);

                        OnProgressUpdate(progressMessage, percentComplete);
                    }

                    dataLine = reader.ReadLine();
                }

                reader.Close();

                return newFilePath;
            }

            return pathOfFileToFix;
        }

        private void EvaluateRules(
            IList<RuleDefinitionExtended> ruleDetails,
            string proteinName,
            string textToTest,
            int testTextOffsetInLine,
            string entireLine,
            int contextLength)
        {
            for (var index = 0; index <= ruleDetails.Count - 1; index++)
            {
                var ruleDetail = ruleDetails[index];

                var match = ruleDetail.MatchRegEx.Match(textToTest);

                if (ruleDetail.RuleDefinition.MatchIndicatesProblem && match.Success ||
                    !(ruleDetail.RuleDefinition.MatchIndicatesProblem && !match.Success))
                {
                    string extraInfo;
                    if (ruleDetail.RuleDefinition.DisplayMatchAsExtraInfo)
                    {
                        extraInfo = match.ToString();
                    }
                    else
                    {
                        extraInfo = string.Empty;
                    }

                    var charIndexOfMatch = testTextOffsetInLine + match.Index;
                    if (ruleDetail.RuleDefinition.Severity >= 5)
                    {
                        RecordFastaFileError(LineCount, charIndexOfMatch, proteinName,
                            ruleDetail.RuleDefinition.CustomRuleID, extraInfo,
                            ExtractContext(entireLine, charIndexOfMatch, contextLength));
                    }
                    else
                    {
                        RecordFastaFileWarning(LineCount, charIndexOfMatch, proteinName,
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
            out bool processingDuplicateOrInvalidProtein)
        {
            var duplicateName = proteinNames.Contains(proteinName);
            skipDuplicateProtein = false;

            if (duplicateName && mGenerateFixedFastaFile)
            {
                if (mFixedFastaOptions.RenameProteinsWithDuplicateNames)
                {
                    var letterToAppend = 'b';
                    var numberToAppend = 0;
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
                                numberToAppend++;
                            }
                            else
                            {
                                letterToAppend = (char)(letterToAppend + 1);
                            }
                        }
                    }
                    while (duplicateName);

                    RecordFastaFileWarning(LineCount, 1, proteinName, (int)MessageCodeConstants.RenamedProtein, "--> " + newProteinName, string.Empty);

                    proteinName = string.Copy(newProteinName);
                    mFixedFastaStats.DuplicateNameProteinsRenamed++;
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
                    RecordFastaFileError(LineCount, 1, proteinName, (int)MessageCodeConstants.DuplicateProteinName);
                    if (mSaveBasicProteinHashInfoFile)
                    {
                        processingDuplicateOrInvalidProtein = false;
                    }
                    else
                    {
                        processingDuplicateOrInvalidProtein = true;
                        mFixedFastaStats.DuplicateNameProteinsSkipped++;
                    }
                }
                else
                {
                    RecordFastaFileWarning(LineCount, 1, proteinName, (int)MessageCodeConstants.DuplicateProteinName);
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

        private string ExtractContext(string text, int startIndex, int contextLength = DEFAULT_CONTEXT_LENGTH)
        {
            // Note that contextLength should be an odd number; if it isn't, we'll add 1 to it

            if (contextLength % 2 == 0)
            {
                contextLength++;
            }
            else if (contextLength < 1)
            {
                contextLength = 1;
            }

            if (text == null)
            {
                return string.Empty;
            }

            if (text.Length <= 1)
            {
                return text;
            }

            // Define the start index for extracting the context from text
            var contextStartIndex = startIndex - (int)Math.Round((contextLength - 1) / (double)2);
            if (contextStartIndex < 0)
                contextStartIndex = 0;

            // Define the end index for extracting the context from text
            var contextEndIndex = Math.Max(startIndex + (int)Math.Round((contextLength - 1) / (double)2), contextStartIndex + contextLength - 1);
            if (contextEndIndex >= text.Length)
            {
                contextEndIndex = text.Length - 1;
            }

            // Return the context portion of text
            return text.Substring(contextStartIndex, contextEndIndex - contextStartIndex + 1);
        }

        private string FlattenArray(IEnumerable<string> items, char sepChar)
        {
            if (items == null)
            {
                return string.Empty;
            }

            return string.Join(sepChar.ToString(), items);
        }

        /// <summary>
        /// Convert a list of strings to a tab-delimited string
        /// </summary>
        /// <param name="dataValues"></param>
        private string FlattenList(IEnumerable<string> dataValues)
        {
            return FlattenArray(dataValues, '\t');
        }

        /// <summary>
        /// Find the first space (or first tab) in the protein header line
        /// </summary>
        /// <param name="headerLine"></param>
        /// <remarks>Used for determining protein name</remarks>
        private int GetBestSpaceIndex(string headerLine)
        {
            var spaceIndex = headerLine.IndexOfAny(mProteinAccessionSepChars);
            if (spaceIndex == 1)
            {
                // Space found directly after the > symbol
                return spaceIndex;
            }

            var tabIndex = headerLine.IndexOf('\t');

            if (tabIndex <= 0)
                return spaceIndex;

            if (tabIndex == 1)
            {
                // Tab character found directly after the > symbol
                return tabIndex;
            }

            // Tab character found; does it separate the protein name and description?
            if (spaceIndex <= 0 || tabIndex < spaceIndex)
            {
                return tabIndex;
            }

            return spaceIndex;
        }

        /// <summary>
        /// Get default file extensions to parse
        /// </summary>
        public override IList<string> GetDefaultExtensionsToParse()
        {
            return new List<string>() { ".fasta", ".faa" };
        }

        /// <summary>
        /// Get the error message
        /// </summary>
        /// <returns>The error message, or an empty string if no error</returns>
        public override string GetErrorMessage()
        {
            if (ErrorCode == ProcessFilesErrorCodes.LocalizedError ||
                ErrorCode == ProcessFilesErrorCodes.NoError)
            {
                return mLocalErrorCode switch
                {
                    ValidateFastaFileErrorCodes.NoError => "",
                    ValidateFastaFileErrorCodes.OptionsSectionNotFound => "The section " + XML_SECTION_OPTIONS + " was not found in the parameter file",
                    ValidateFastaFileErrorCodes.ErrorReadingInputFile => "Error reading input file",
                    ValidateFastaFileErrorCodes.ErrorCreatingStatsFile => "Error creating stats output file",
                    ValidateFastaFileErrorCodes.ErrorVerifyingLinefeedAtEOF => "Error verifying linefeed at end of file",
                    ValidateFastaFileErrorCodes.UnspecifiedError => "Unspecified localized error",
                    _ => "Unknown error state"  // This shouldn't happen
                };
            }

            return GetBaseClassErrorMessage();
        }

        private string GetFileErrorTextByIndex(int fileErrorIndex, string sepChar)
        {
            if (mFileErrors.Messages.Count == 0 || fileErrorIndex < 0 || fileErrorIndex >= mFileErrors.Messages.Count)
            {
                return string.Empty;
            }

            var fileError = mFileErrors.Messages[fileErrorIndex];

            string proteinName;
            if (string.IsNullOrEmpty(fileError.ProteinName))
            {
                proteinName = "N/A";
            }
            else
            {
                proteinName = string.Copy(fileError.ProteinName);
            }

            return LookupMessageType(MsgTypeConstants.ErrorMsg) + sepChar +
                   "Line " + fileError.LineNumber.ToString() + sepChar +
                   "Col " + fileError.ColNumber.ToString() + sepChar +
                   proteinName + sepChar +
                   LookupMessageDescription(fileError.MessageCode, fileError.ExtraInfo) + sepChar + fileError.Context;
        }

        private MsgInfo GetFileErrorByIndex(int fileErrorIndex)
        {
            if (mFileErrors.Messages.Count == 0 || fileErrorIndex < 0 || fileErrorIndex >= mFileErrors.Messages.Count)
            {
                return new MsgInfo();
            }

            return mFileErrors.Messages[fileErrorIndex];
        }

        /// <summary>
        /// Retrieve the errors reported by the validator
        /// </summary>
        /// <remarks>Used by CustomValidateFastaFiles</remarks>
        protected List<MsgInfo> GetFileErrors()
        {
            var fileErrors = new List<MsgInfo>();

            foreach (var fileError in mFileErrors.Messages)
            {
                fileErrors.Add(fileError);
            }

            return fileErrors;
        }

        private string GetFileWarningTextByIndex(int fileWarningIndex, string sepChar)
        {
            if (mFileWarnings.Messages.Count == 0 || fileWarningIndex < 0 || fileWarningIndex >= mFileWarnings.Messages.Count)
            {
                return string.Empty;
            }

            var fileWarning = mFileWarnings.Messages[fileWarningIndex];

            string proteinName;
            if (string.IsNullOrEmpty(fileWarning.ProteinName))
            {
                proteinName = "N/A";
            }
            else
            {
                proteinName = string.Copy(fileWarning.ProteinName);
            }

            return LookupMessageType(MsgTypeConstants.WarningMsg) + sepChar +
                   "Line " + fileWarning.LineNumber.ToString() + sepChar +
                   "Col " + fileWarning.ColNumber.ToString() + sepChar +
                   proteinName + sepChar +
                   LookupMessageDescription(fileWarning.MessageCode, fileWarning.ExtraInfo) +
                   sepChar + fileWarning.Context;
        }

        private MsgInfo GetFileWarningByIndex(int fileWarningIndex)
        {
            if (mFileWarnings.Messages.Count == 0 || fileWarningIndex < 0 || fileWarningIndex >= mFileWarnings.Messages.Count)
            {
                return new MsgInfo();
            }

            return mFileWarnings.Messages[fileWarningIndex];
        }

        /// <summary>
        /// Retrieve the warnings reported by the validator
        /// </summary>
        /// <remarks>Used by CustomValidateFastaFiles</remarks>
        protected List<MsgInfo> GetFileWarnings()
        {
            var fileWarnings = new List<MsgInfo>();

            foreach (var fileWarning in mFileWarnings.Messages)
            {
                fileWarnings.Add(fileWarning);
            }

            return fileWarnings;
        }

        private string GetOutputDirectory(string inputFileDirectoryPath)
        {
            if (string.IsNullOrWhiteSpace(mOutputDirectoryPath))
            {
                return Path.GetDirectoryName(inputFileDirectoryPath) ?? string.Empty;
            }

            return mOutputDirectoryPath;
        }

        private string GetProcessMemoryUsageWithTimestamp()
        {
            return GetTimeStamp() + "\t" + MemoryUsageLogger.GetProcessMemoryUsageMB().ToString("0.0") + " MB in use";
        }

        private string GetTimeStamp()
        {
            // Record the current time
            return DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString();
        }

        private void InitializeLocalVariables()
        {
            mLocalErrorCode = ValidateFastaFileErrorCodes.NoError;

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

            SetDefaultRules();

            ResetStructures();
            mFastaFilePath = string.Empty;

            mMemoryUsageLogger = new MemoryUsageLogger(string.Empty);
            mProcessMemoryUsageMBAtStart = MemoryUsageLogger.GetProcessMemoryUsageMB();

            // ReSharper disable once ConditionIsAlwaysTrueOrFalse
            if (REPORT_DETAILED_MEMORY_USAGE)
            {
#pragma warning disable CS0162 // Unreachable code detected
                // mMemoryUsageMBAtStart = mMemoryUsageLogger.GetFreeMemoryMB()
                Console.WriteLine(MEM_USAGE_PREFIX + mMemoryUsageLogger.GetMemoryUsageHeader());
                Console.WriteLine(MEM_USAGE_PREFIX + mMemoryUsageLogger.GetMemoryUsageSummary());
#pragma warning restore CS0162 // Unreachable code detected
            }

            mTempFilesToDelete = new List<string>();
        }

        private void InitializeRuleDetails(
            IReadOnlyList<RuleDefinition> ruleDefinitions,
            out RuleDefinitionExtended[] ruleDetails)
        {
            if (ruleDefinitions == null || ruleDefinitions.Count == 0)
            {
                ruleDetails = Array.Empty<RuleDefinitionExtended>();
            }
            else
            {
                ruleDetails = new RuleDefinitionExtended[ruleDefinitions.Count];

                for (var index = 0; index <= ruleDefinitions.Count - 1; index++)
                {
                    try
                    {
                        ruleDetails[index] = new RuleDefinitionExtended(
                            ruleDefinitions[index],
                            new Regex(
                                ruleDefinitions[index].MatchRegEx,
                                RegexOptions.Singleline |
                                RegexOptions.Compiled))
                        {
                            Valid = true
                        };
                    }
                    catch (Exception ex)
                    {
                        // Ignore the error, but mark .Valid = false
                        ruleDetails[index].Valid = false;

                        OnWarningEvent(string.Format("Invalid RegEx '{0}: {1} ", ruleDefinitions[index].MatchRegEx ?? "null", ex));
                    }
                }
            }
        }

        private bool LoadExistingProteinHashFile(
            string proteinHashFilePath,
            out NestedStringIntList preloadedProteinNamesToKeep)
        {
            // List of protein names to keep
            // Keys are protein names, values are the number of entries written to the fixed FASTA file for the given protein name
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
                var cachedHeaderLine = string.Empty;

                ShowMessage("Examining pre-existing protein hash file to count the number of entries: " + Path.GetFileName(proteinHashFilePath));
                var lastStatus = DateTime.UtcNow;
                var progressDotShown = false;

                using (var hashFileReader = new StreamReader(new FileStream(proteinHashFile.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)))
                {
                    if (!hashFileReader.EndOfStream)
                    {
                        cachedHeaderLine = hashFileReader.ReadLine();

                        if (!string.IsNullOrWhiteSpace(cachedHeaderLine))
                        {
                            var headerNames = cachedHeaderLine.Split('\t');

                            for (var colIndex = 0; colIndex <= headerNames.Length - 1; colIndex++)
                            {
                                headerInfo.Add(headerNames[colIndex], colIndex);
                            }
                        }

                        proteinHashFileLines = 1;
                    }

                    while (!hashFileReader.EndOfStream)
                    {
                        hashFileReader.ReadLine();
                        proteinHashFileLines++;

                        if (proteinHashFileLines % 10000 != 0)
                            continue;

                        if (DateTime.UtcNow.Subtract(lastStatus).TotalSeconds < 10)
                            continue;

                        Console.Write(".");
                        progressDotShown = true;
                        lastStatus = DateTime.UtcNow;
                    }
                }

                if (progressDotShown)
                    Console.WriteLine();

                if (headerInfo.Count == 0)
                {
                    ShowErrorMessage("Protein hash file is empty (or has an empty header line): " + proteinHashFilePath);
                    return false;
                }

                if (!headerInfo.TryGetValue(SEQUENCE_HASH_COLUMN, out var sequenceHashColumnIndex))
                {
                    ShowErrorMessage("Protein hash file is missing the " + SEQUENCE_HASH_COLUMN + " column: " + proteinHashFilePath);
                    return false;
                }

                if (!headerInfo.TryGetValue(PROTEIN_NAME_COLUMN, out var proteinNameColumnIndex))
                {
                    ShowErrorMessage("Protein hash file is missing the " + PROTEIN_NAME_COLUMN + " column: " + proteinHashFilePath);
                    return false;
                }

                if (!headerInfo.TryGetValue(SEQUENCE_LENGTH_COLUMN, out var sequenceLengthColumnIndex))
                {
                    ShowErrorMessage("Protein hash file is missing the " + SEQUENCE_LENGTH_COLUMN + " column: " + proteinHashFilePath);
                    return false;
                }

                var baseHashFileName = Path.GetFileNameWithoutExtension(proteinHashFile.Name);
                string sortedProteinHashSuffix;
                var proteinHashFilenameSuffixNoExtension = Path.GetFileNameWithoutExtension(PROTEIN_HASHES_FILENAME_SUFFIX);

                if (baseHashFileName.EndsWith(proteinHashFilenameSuffixNoExtension))
                {
                    baseHashFileName = baseHashFileName.Substring(0, baseHashFileName.Length - proteinHashFilenameSuffixNoExtension.Length);
                    sortedProteinHashSuffix = proteinHashFilenameSuffixNoExtension;
                }
                else
                {
                    sortedProteinHashSuffix = string.Empty;
                }

                var baseDataFileName = Path.Combine(proteinHashFile.Directory.FullName, baseHashFileName);

                // Note: do not add sortedProteinHashFilePath to mTempFilesToDelete
                var sortedProteinHashFilePath = baseDataFileName + sortedProteinHashSuffix + "_Sorted.tmp";

                var sortedProteinHashFile = new FileInfo(sortedProteinHashFilePath);
                long sortedHashFileLines = 0;
                var sortRequired = true;

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
                            var headerLine = sortedHashFileReader.ReadLine();
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
                                sortedHashFileLines++;

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
                    var sortHashSuccess = SortFile(proteinHashFile, sequenceHashColumnIndex, sortedProteinHashFilePath);
                    if (!sortHashSuccess)
                    {
                        return false;
                    }
                }

                ShowMessage("Determining the best spanner length for protein names");

                // Examine the protein names in the sequence hash file to determine the appropriate spanner length for the names
                var spannerCharLength = NestedStringIntList.AutoDetermineSpannerCharLength(proteinHashFile, proteinNameColumnIndex, true);
                const bool RAISE_EXCEPTION_IF_ADDED_DATA_NOT_SORTED = true;

                // List of protein names to keep
                preloadedProteinNamesToKeep = new NestedStringIntList(spannerCharLength, RAISE_EXCEPTION_IF_ADDED_DATA_NOT_SORTED);

                lastStatus = DateTime.UtcNow;
                progressDotShown = false;
                Console.WriteLine();
                ShowMessage("Finding the name of the first protein for each protein hash");

                var currentHash = string.Empty;
                var currentHashSeqLength = string.Empty;
                var proteinNamesCurrentHash = new SortedSet<string>();

                var linesRead = 0;
                var proteinNamesUnsortedCount = 0;

                var proteinNamesUnsortedFilePath = baseDataFileName + "_ProteinNamesUnsorted.tmp";
                var proteinNamesToKeepFilePath = baseDataFileName + "_ProteinNamesToKeep.tmp";
                var uniqueProteinSeqsFilePath = baseDataFileName + "_UniqueProteinSeqs.txt";
                var uniqueProteinSeqDuplicateFilePath = baseDataFileName + "_UniqueProteinSeqDuplicates.txt";

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

                    var currentSequenceIndex = 0;

                    while (!sortedHashFileReader.EndOfStream)
                    {
                        var dataLine = sortedHashFileReader.ReadLine();
                        linesRead++;

                        var dataValues = dataLine.Split('\t');

                        var proteinName = dataValues[proteinNameColumnIndex];
                        var proteinHash = dataValues[sequenceHashColumnIndex];

                        proteinNamesUnsortedWriter.WriteLine(FlattenList(new List<string>() { proteinName, proteinHash }));
                        proteinNamesUnsortedCount++;

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
                            currentSequenceIndex++;
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
                var sortedProteinNamesToKeepFilePath = baseDataFileName + "_ProteinsToKeepSorted.tmp";
                mTempFilesToDelete.Add(sortedProteinNamesToKeepFilePath);

                ShowMessage("Sorting the protein names to keep file to create " + Path.GetFileName(sortedProteinNamesToKeepFilePath));
                var sortProteinNamesToKeepSuccess = SortFile(new FileInfo(proteinNamesToKeepFilePath), 0, sortedProteinNamesToKeepFilePath);

                if (!sortProteinNamesToKeepSuccess)
                {
                    return false;
                }

                var lastProteinAdded = string.Empty;

                // Read the sorted protein names to keep file and cache the protein names in memory
                using (var sortedNamesFileReader = new StreamReader(new FileStream(sortedProteinNamesToKeepFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)))
                {
                    // Skip the header line
                    sortedNamesFileReader.ReadLine();

                    while (!sortedNamesFileReader.EndOfStream)
                    {
                        var proteinName = sortedNamesFileReader.ReadLine();

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
                var sortedProteinNamesFilePath = baseDataFileName + "_ProteinNamesSorted.tmp";
                mTempFilesToDelete.Add(sortedProteinNamesFilePath);

                Console.WriteLine();
                ShowMessage("Sorting the protein names file to create " + Path.GetFileName(sortedProteinNamesFilePath));
                var sortProteinNamesSuccess = SortFile(new FileInfo(proteinNamesUnsortedFilePath), 0, sortedProteinNamesFilePath);

                if (!sortProteinNamesSuccess)
                {
                    return false;
                }

                var proteinNameSummaryFilePath = baseDataFileName + "_ProteinNameSummary.txt";

                // We can now safely delete some files to free up disk space
                try
                {
                    System.Threading.Thread.Sleep(100);
                    File.Delete(proteinNamesUnsortedFilePath);
                }
                catch (Exception)
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

                    var lastProtein = string.Empty;
                    var warningShown = false;
                    var duplicateCount = 0;

                    linesRead = 0;
                    lastStatus = DateTime.UtcNow;
                    progressDotShown = false;

                    while (!sortedProteinNamesReader.EndOfStream)
                    {
                        var dataLine = sortedProteinNamesReader.ReadLine();
                        var dataValues = dataLine.Split('\t');

                        linesRead++;

                        if (dataValues.Length < 2)
                        {
                            OnWarningEvent(string.Format("Skipping line {0} in the protein hash file since it does not have 2 columns", linesRead));
                            continue;
                        }

                        var currentProtein = dataValues[0];
                        var sequenceHash = dataValues[1];

                        var firstProteinForHash = preloadedProteinNamesToKeep.Contains(currentProtein);
                        var duplicateProtein = false;

                        if (string.Equals(lastProtein, currentProtein))
                        {
                            duplicateProtein = true;

                            if (!warningShown)
                            {
                                ShowMessage("WARNING: the protein hash file has duplicate protein names: " + Path.GetFileName(proteinHashFilePath));
                                warningShown = true;
                            }

                            duplicateCount++;

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

        /// <summary>
        /// Load settings from a parameter file
        /// </summary>
        /// <param name="parameterFilePath"></param>
        public bool LoadParameterFileSettings(string parameterFilePath)
        {
            var settingsFile = new XmlSettingsFileAccessor();

            var customRulesLoaded = false;

            try
            {
                if (string.IsNullOrEmpty(parameterFilePath))
                {
                    // No parameter file specified; nothing to load
                    return true;
                }

                if (!File.Exists(parameterFilePath))
                {
                    // See if parameterFilePath points to a file in the same directory as the application
                    parameterFilePath = Path.Combine(
                        Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) ?? string.Empty,
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
                        var characterList = settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "LongProteinNameSplitChars", string.Empty);
                        if (!string.IsNullOrEmpty(characterList))
                        {
                            // Update mFixedFastaOptions.LongProteinNameSplitChars with characterList
                            LongProteinNameSplitChars = characterList;
                        }

                        characterList = settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ProteinNameInvalidCharsToRemove", string.Empty);
                        if (!string.IsNullOrEmpty(characterList))
                        {
                            // Update mFixedFastaOptions.ProteinNameInvalidCharsToRemove with characterList
                            ProteinNameInvalidCharsToRemove = characterList;
                        }

                        characterList = settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ProteinNameFirstRefSepChars", string.Empty);
                        if (!string.IsNullOrEmpty(characterList))
                        {
                            // Update mProteinNameFirstRefSepChars
                            ProteinNameFirstRefSepChars = characterList;
                        }

                        characterList = settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ProteinNameSubsequentRefSepChars", string.Empty);
                        if (!string.IsNullOrEmpty(characterList))
                        {
                            // Update mProteinNameSubsequentRefSepChars
                            ProteinNameSubsequentRefSepChars = characterList;
                        }
                    }

                    // Read the custom rules
                    // If all of the sections are missing, use the default rules

                    var success = ReadRulesFromParameterFile(settingsFile, XML_SECTION_FASTA_HEADER_LINE_RULES, mHeaderLineRules);
                    customRulesLoaded = success;

                    success = ReadRulesFromParameterFile(settingsFile, XML_SECTION_FASTA_PROTEIN_NAME_RULES, mProteinNameRules);
                    customRulesLoaded = customRulesLoaded || success;

                    success = ReadRulesFromParameterFile(settingsFile, XML_SECTION_FASTA_PROTEIN_DESCRIPTION_RULES, mProteinDescriptionRules);
                    customRulesLoaded = customRulesLoaded || success;

                    success = ReadRulesFromParameterFile(settingsFile, XML_SECTION_FASTA_PROTEIN_SEQUENCE_RULES, mProteinSequenceRules);
                    customRulesLoaded = customRulesLoaded || success;
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

        /// <summary>
        /// Obtain the message description by message code
        /// </summary>
        /// <param name="errorMessageCode"></param>
        // ReSharper disable once UnusedMember.Global
        public string LookupMessageDescription(int errorMessageCode)
        {
            return LookupMessageDescription(errorMessageCode, null);
        }

        /// <summary>
        /// Obtain the message description by message code, appending extraInfo
        /// </summary>
        /// <param name="errorMessageCode"></param>
        /// <param name="extraInfo"></param>
        public string LookupMessageDescription(int errorMessageCode, string extraInfo)
        {
            string message;

            switch (errorMessageCode)
            {
                // Error messages
                case (int)MessageCodeConstants.ProteinNameIsTooLong:
                    message = "Protein name is longer than the maximum allowed length of " + mMaximumProteinNameLength.ToString() + " characters";
                    break;
                //case (int)MessageCodeConstants.ProteinNameContainsInvalidCharacters:
                //    message = "Protein name contains invalid characters";
                //    if (!specifiedInvalidProteinNameChars)
                //    {
                //        message += " (should contain letters, numbers, period, dash, underscore, colon, comma, or vertical bar)";
                //        specifiedInvalidProteinNameChars = true;
                //    }
                //    break;
                case (int)MessageCodeConstants.LineStartsWithSpace:
                    message = "Found a line starting with a space";
                    break;
                //case (int)MessageCodeConstants.RightArrowFollowedBySpace:
                //    message = "Space found directly after the > symbol";
                //    break;
                //case (int)MessageCodeConstants.RightArrowFollowedByTab:
                //    message = "Tab character found directly after the > symbol";
                //    break;
                //case (int)MessageCodeConstants.RightArrowButNoProteinName:
                //    message = "Line starts with > but does not contain a protein name";
                //    break;
                case (int)MessageCodeConstants.BlankLineBetweenProteinNameAndResidues:
                    message = "A blank line was found between the protein name and its residues";
                    break;
                case (int)MessageCodeConstants.BlankLineInMiddleOfResidues:
                    message = "A blank line was found in the middle of the residue block for the protein";
                    break;
                case (int)MessageCodeConstants.ResiduesFoundWithoutProteinHeader:
                    message = "Residues were found, but a protein header didn't precede them";
                    break;
                //case (int)MessageCodeConstants.ResiduesWithAsterisk:
                //    message = "An asterisk was found in the residues";
                //    break;
                //case (int)MessageCodeConstants.ResiduesWithSpace:
                //    message = "A space was found in the residues";
                //    break;
                //case (int)MessageCodeConstants.ResiduesWithTab:
                //    message = "A tab character was found in the residues";
                //    break;
                //case (int)MessageCodeConstants.ResiduesWithInvalidCharacters:
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
                case (int)MessageCodeConstants.ProteinEntriesNotFound:
                    message = "File does not contain any protein entries";
                    break;
                case (int)MessageCodeConstants.FinalProteinEntryMissingResidues:
                    message = "The last entry in the file is a protein header line, but there is no protein sequence line after it";
                    break;
                case (int)MessageCodeConstants.FileDoesNotEndWithLinefeed:
                    message = "File does not end in a blank line; this is a problem for SEQUEST";
                    break;
                case (int)MessageCodeConstants.DuplicateProteinName:
                    message = "Duplicate protein name found";
                    break;

                // Warning messages
                case (int)MessageCodeConstants.ProteinNameIsTooShort:
                    message = "Protein name is shorter than the minimum suggested length of " + mMinimumProteinNameLength.ToString() + " characters";
                    break;
                //case (int)MessageCodeConstants.ProteinNameContainsComma:
                //    message = "Protein name contains a comma";
                //    break;
                //case (int)MessageCodeConstants.ProteinNameContainsVerticalBars:
                //    message = "Protein name contains two or more vertical bars";
                //    break;
                //case (int)MessageCodeConstants.ProteinNameContainsWarningCharacters:
                //    message = "Protein name contains undesirable characters";
                //    break;
                //case (int)MessageCodeConstants.ProteinNameWithoutDescription:
                //    message = "Line contains a protein name, but not a description";
                //    break;
                case (int)MessageCodeConstants.BlankLineBeforeProteinName:
                    message = "Blank line found before the protein name; this is acceptable, but not preferred";
                    break;
                // case (int)MessageCodeConstants.ProteinNameAndDescriptionSeparatedByTab:
                //    message = "Protein name is separated from the protein description by a tab";
                //    break;
                case (int)MessageCodeConstants.ResiduesLineTooLong:
                    message = "Residues line is longer than the suggested maximum length of " + mMaximumResiduesPerLine.ToString() + " characters";
                    break;
                //case (int)MessageCodeConstants.ProteinDescriptionWithTab:
                //    message = "Protein description contains a tab character";
                //    break;
                //case (int)MessageCodeConstants.ProteinDescriptionWithQuotationMark:
                //    message = "Protein description contains a quotation mark";
                //    break;
                //case (int)MessageCodeConstants.ProteinDescriptionWithEscapedSlash:
                //    message = "Protein description contains escaped slash: \/";
                //    break;
                //case (int)MessageCodeConstants.ProteinDescriptionWithUndesirableCharacter:
                //    message = "Protein description contains undesirable characters";
                //    break;
                //case (int)MessageCodeConstants.ResiduesLineTooLong:
                //    message = "Residues line is longer than the suggested maximum length of " + mMaximumResiduesPerLine.ToString + " characters";
                //    break;
                //case (int)MessageCodeConstants.ResiduesLineContainsU:
                //    message = "Residues line contains U (selenocysteine); this residue is unsupported by SEQUEST";
                //    break;

                case (int)MessageCodeConstants.DuplicateProteinSequence:
                    message = "Duplicate protein sequences found";
                    break;
                case (int)MessageCodeConstants.RenamedProtein:
                    message = "Renamed protein because duplicate name";
                    break;
                case (int)MessageCodeConstants.ProteinRemovedSinceDuplicateSequence:
                    message = "Removed protein since duplicate sequence";
                    break;
                case (int)MessageCodeConstants.DuplicateProteinNameRetained:
                    message = "Duplicate protein retained in fixed file";
                    break;
                case (int)MessageCodeConstants.UnspecifiedError:
                    message = "Unspecified error";
                    break;
                default:
                    // Search the custom rules for the given code
                    var matchFound = SearchRulesForID(mHeaderLineRules, errorMessageCode, out message);

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

        private string LookupMessageType(MsgTypeConstants EntryType)
        {
            return EntryType switch
            {
                MsgTypeConstants.ErrorMsg => "Error",
                MsgTypeConstants.WarningMsg => "Warning",
                _ => "Status"
            };
        }

        /// <summary>
        /// Validate a single FASTA file
        /// </summary>
        /// <returns>True if success; false if a fatal error</returns>
        /// <remarks>
        /// Note that .ProcessFile returns True if a file is successfully processed (even if errors are found)
        /// Used by CustomValidateFastaFiles
        /// </remarks>
        // ReSharper disable once UnusedMember.Global
        protected bool SimpleProcessFile(string inputFilePath)
        {
            return ProcessFile(inputFilePath, null, null, false);
        }

        /// <summary>
        /// Main processing function
        /// </summary>
        /// <param name="inputFilePath"></param>
        /// <param name="outputDirectoryPath"></param>
        /// <param name="parameterFilePath"></param>
        /// <param name="resetErrorCode"></param>
        /// <returns>True if success, False if failure</returns>
        public override bool ProcessFile(
            string inputFilePath,
            string outputDirectoryPath,
            string parameterFilePath,
            bool resetErrorCode)
        {
            if (resetErrorCode)
            {
                SetLocalErrorCode(ValidateFastaFileErrorCodes.NoError);
            }

            if (!LoadParameterFileSettings(parameterFilePath))
            {
                var statusMessage = "Parameter file load error: " + parameterFilePath;
                OnWarningEvent(statusMessage);
                if (ErrorCode == ProcessFilesErrorCodes.NoError)
                {
                    SetBaseClassErrorCode(ProcessFilesErrorCodes.InvalidParameterFile);
                }

                return false;
            }

            try
            {
                if (string.IsNullOrEmpty(inputFilePath))
                {
                    ShowWarning("Input file name is empty");
                    SetBaseClassErrorCode(ProcessFilesErrorCodes.InvalidInputFilePath);
                    return false;
                }

                Console.WriteLine();
                ShowMessage("Parsing " + Path.GetFileName(inputFilePath));

                if (!CleanupFilePaths(ref inputFilePath, ref outputDirectoryPath))
                {
                    SetBaseClassErrorCode(ProcessFilesErrorCodes.FilePathError);
                    return false;
                }
                else
                {
                    // List of protein names to keep
                    // Keys are protein names, values are the number of entries written to the fixed FASTA file for the given protein name
                    NestedStringIntList preloadedProteinNamesToKeep = null;

                    if (!string.IsNullOrEmpty(ExistingProteinHashFile))
                    {
                        var loadSuccess = LoadExistingProteinHashFile(ExistingProteinHashFile, out preloadedProteinNamesToKeep);
                        if (!loadSuccess)
                        {
                            return false;
                        }
                    }

                    try
                    {
                        // Obtain the full path to the input file
                        var ioFile = new FileInfo(inputFilePath);
                        var inputFilePathFull = ioFile.FullName;

                        var success = AnalyzeFastaFile(inputFilePathFull, preloadedProteinNamesToKeep);

                        if (success)
                        {
                            ReportResults(outputDirectoryPath, mOutputToStatsFile);
                            DeleteTempFiles();
                            return true;
                        }
                        else
                        {
                            if (mOutputToStatsFile)
                            {
                                mStatsFilePath = ConstructStatsFilePath(outputDirectoryPath);

                                using var statsFileWriter = new StreamWriter(mStatsFilePath, true);

                                statsFileWriter.WriteLine(GetTimeStamp() + "\t" +
                                                          "Error parsing " +
                                                          Path.GetFileName(inputFilePath) + ": " + GetErrorMessage());
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
            catch (Exception ex)
            {
                OnErrorEvent("Error in ProcessFile", ex);
                return false;
            }
        }

        private readonly char[] extraCharsToTrim = { '|', ' ' };

        private void PrependExtraTextToProteinDescription(string extraProteinNameText, ref string proteinDescription)
        {
            if (!string.IsNullOrEmpty(extraProteinNameText))
            {
                // If extraProteinNameText ends in a vertical bar and/or space, them remove them
                extraProteinNameText = extraProteinNameText.TrimEnd(extraCharsToTrim);

                if (!string.IsNullOrEmpty(proteinDescription))
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
            StringBuilder currentResidues,
            NestedStringDictionary<int> proteinSequenceHashes,
            List<ProteinHashInfo> proteinSeqHashInfo,
            bool consolidateDupsIgnoreILDiff,
            TextWriter fixedFastaWriter,
            int currentValidResidueLineLengthMax,
            TextWriter sequenceHashWriter)
        {
            // Check for and remove any asterisks at the end of the residues
            while (currentResidues.Length > 0 && currentResidues[currentResidues.Length - 1] == '*')
            {
                currentResidues.Remove(currentResidues.Length - 1, 1);
            }

            if (currentResidues.Length > 0)
            {
                // Remove any spaces from the residues

                if (mCheckForDuplicateProteinSequences || mSaveBasicProteinHashInfoFile)
                {
                    // Process the previous protein entry to store a hash of the protein sequence
                    ProcessSequenceHashInfo(
                        proteinName, currentResidues,
                        proteinSequenceHashes,
                        proteinSeqHashInfo,
                        consolidateDupsIgnoreILDiff, sequenceHashWriter);
                }

                if (mGenerateFixedFastaFile && mFixedFastaOptions.WrapLongResidueLines)
                {
                    // Write out the residues
                    // Wrap the lines at currentValidResidueLineLengthMax characters (but do not allow to be longer than mMaximumResiduesPerLine residues)

                    var wrapLength = currentValidResidueLineLengthMax;
                    if (wrapLength <= 0 || wrapLength > mMaximumResiduesPerLine)
                    {
                        wrapLength = mMaximumResiduesPerLine;
                    }

                    if (wrapLength < 10)
                    {
                        // Do not allow wrapLength to be less than 10
                        wrapLength = 10;
                    }

                    var index = 0;
                    var proteinResidueCount = currentResidues.Length;
                    while (index < currentResidues.Length)
                    {
                        var length = Math.Min(wrapLength, proteinResidueCount - index);
                        fixedFastaWriter.WriteLine(currentResidues.ToString(index, length));
                        index += wrapLength;
                    }
                }

                currentResidues.Clear();
            }
        }

        private void ProcessSequenceHashInfo(
            string proteinName,
            StringBuilder residues,
            NestedStringDictionary<int> proteinSequenceHashes,
            List<ProteinHashInfo> proteinSeqHashInfo,
            bool consolidateDupsIgnoreILDiff,
            TextWriter sequenceHashWriter)
        {
            try
            {
                if (residues.Length > 0)
                {
                    // Compute the hash value for residues
                    var computedHash = ComputeProteinHash(residues, consolidateDupsIgnoreILDiff);

                    if (sequenceHashWriter != null)
                    {
                        var dataValues = new List<string>()
                        {
                            ProteinCount.ToString(),
                            proteinName,
                            residues.Length.ToString(),
                            computedHash
                        };

                        sequenceHashWriter.WriteLine(FlattenList(dataValues));
                    }

                    if (mCheckForDuplicateProteinSequences && proteinSequenceHashes != null)
                    {
                        // See if proteinSequenceHashes contains hash
                        if (proteinSequenceHashes.TryGetValue(computedHash, out var seqHashLookupPointer))
                        {
                            // Value exists; update the entry in proteinSeqHashInfo
                            CachedSequenceHashInfoUpdate(proteinSeqHashInfo[seqHashLookupPointer], proteinName);
                        }
                        else
                        {
                            // Value not yet present; add it
                            var index = CachedSequenceHashInfoUpdateAppend(
                                proteinSeqHashInfo,
                                computedHash, residues, proteinName);

                            proteinSequenceHashes.Add(computedHash, index);
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

        private void CachedSequenceHashInfoUpdate(ProteinHashInfo proteinSeqHashInfo, string proteinName)
        {
            if ((proteinSeqHashInfo.ProteinNameFirst ?? "") == (proteinName ?? ""))
            {
                proteinSeqHashInfo.DuplicateProteinNameCount++;
            }
            else
            {
                proteinSeqHashInfo.AddAdditionalProtein(proteinName);
            }
        }

        private int CachedSequenceHashInfoUpdateAppend(
            List<ProteinHashInfo> proteinSeqHashInfo,
            string computedHash,
            StringBuilder sbCurrentResidues,
            string proteinName)
        {
            if (proteinSeqHashInfo.Count >= proteinSeqHashInfo.Capacity)
            {
                // Need to reserve more space in proteinSeqHashInfo
                if (proteinSeqHashInfo.Capacity < 1000000)
                {
                    proteinSeqHashInfo.Capacity *= 2;
                }
                else
                {
                    proteinSeqHashInfo.Capacity = (int)Math.Round(proteinSeqHashInfo.Capacity * 1.2);
                }
            }

            var newProteinHashInfo = new ProteinHashInfo(computedHash, sbCurrentResidues, proteinName);
            var addLocation = proteinSeqHashInfo.Count;
            proteinSeqHashInfo.Add(newProteinHashInfo);

            return addLocation;
        }

        /// <summary>
        /// Read rules from an XML parameter file
        /// </summary>
        /// <param name="settingsFile"></param>
        /// <param name="sectionName"></param>
        /// <param name="rules"></param>
        /// <returns>True if the section named sectionName is present and if it contains an item with keyName = "RuleCount"</returns>
        /// <remarks>Even if RuleCount = 0, this function will return True</remarks>
        private bool ReadRulesFromParameterFile(
            XmlSettingsFileAccessor settingsFile,
            string sectionName,
            List<RuleDefinition> rules)
        {
            var success = false;

            var ruleCount = settingsFile.GetParam(sectionName, XML_OPTION_ENTRY_RULE_COUNT, -1);

            if (ruleCount >= 0)
            {
                ClearRulesDataStructure(rules);

                for (var ruleNumber = 1; ruleNumber <= ruleCount; ruleNumber++)
                {
                    var ruleBase = "Rule" + ruleNumber.ToString();

                    var newRule = new RuleDefinition(settingsFile.GetParam(sectionName, ruleBase + "MatchRegEx", string.Empty));

                    // Only read the rule settings if MatchRegEx contains 1 or more characters
                    if (newRule.MatchRegEx.Length > 0)
                    {
                        newRule.MatchIndicatesProblem = settingsFile.GetParam(sectionName, ruleBase + "MatchIndicatesProblem", true);
                        newRule.MessageWhenProblem = settingsFile.GetParam(sectionName, ruleBase + "MessageWhenProblem", "Error found with RegEx " + newRule.MatchRegEx);
                        newRule.Severity = settingsFile.GetParam(sectionName, ruleBase + "Severity", (short)3);
                        newRule.DisplayMatchAsExtraInfo = settingsFile.GetParam(sectionName, ruleBase + "DisplayMatchAsExtraInfo", false);

                        SetRule(rules, newRule.MatchRegEx, newRule.MatchIndicatesProblem, newRule.MessageWhenProblem, newRule.Severity, newRule.DisplayMatchAsExtraInfo);
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
                mFileErrors,
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
            RecordFastaFileProblemWork(
                mFileWarnings, lineNumber, charIndex, proteinName,
                warningMessageCode, extraInfo, context);
        }

        private void RecordFastaFileProblemWork(
            MsgInfosAndSummary items,
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
                if (!items.MessageCodeToErrorStats.TryGetValue(messageCode, out var errorStats))
                {
                    errorStats = new ErrorStats(messageCode);
                    items.MessageCodeToErrorStats.Add(messageCode, errorStats);
                }

                if (errorStats.CountSpecified >= mMaximumFileErrorsToTrack)
                {
                    errorStats.CountUnspecified++;
                }
                else
                {
                    items.Messages.Add(new MsgInfo(
                        lineNumber,
                        charIndex + 1,
                        proteinName ?? string.Empty,
                        messageCode,
                        extraInfo ?? string.Empty,
                        context ?? string.Empty));

                    errorStats.CountSpecified++;
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
            try
            {
                // Define the output file path
                var timeStamp = GetTimeStamp().Replace(" ", "_").Replace(":", "_").Replace("/", "_");

                var inputFile = new FileInfo(parameterFilePath);

                var outputFile = new FileInfo(parameterFilePath + "_" + timeStamp + ".fixed");

                // Open the input file and output file
                using (var reader = new StreamReader(inputFile.FullName))
                using (var writer = new StreamWriter(new FileStream(outputFile.FullName, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)))
                {
                    // Parse each line in the file
                    while (!reader.EndOfStream)
                    {
                        var lineIn = reader.ReadLine();

                        if (lineIn != null)
                        {
                            lineIn = lineIn.Replace("&gt;", ">").Replace("&lt;", "<");
                            writer.WriteLine(lineIn);
                        }
                    }
                }

                // Delete the input file
                inputFile.Delete();

                // Rename the output file to the input file
                outputFile.MoveTo(parameterFilePath);
            }
            catch (Exception ex)
            {
                OnErrorEvent("Error in ReplaceXMLCodesWithText", ex);
            }
        }

        private void ReportMemoryUsage()
        {
            // ReSharper disable once ConditionIsAlwaysTrueOrFalse
            if (REPORT_DETAILED_MEMORY_USAGE)
            {
#pragma warning disable CS0162 // Unreachable code detected
                Console.WriteLine(MEM_USAGE_PREFIX + mMemoryUsageLogger.GetMemoryUsageSummary());
#pragma warning restore CS0162 // Unreachable code detected
            }
            else
            {
                Console.WriteLine(MEM_USAGE_PREFIX + GetProcessMemoryUsageWithTimestamp());
            }
        }

        private void ReportMemoryUsage(
            NestedStringIntList preloadedProteinNamesToKeep,
            NestedStringDictionary<int> proteinSequenceHashes,
            ICollection<string> proteinNames,
            IEnumerable<ProteinHashInfo> proteinSeqHashInfo)
        {
            Console.WriteLine();
            ReportMemoryUsage();

            if (preloadedProteinNamesToKeep?.Count > 0)
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
            NestedStringDictionary<int> proteinNameFirst,
            NestedStringDictionary<int> proteinsWritten,
            NestedStringDictionary<string> duplicateProteinList)
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
            string outputDirectoryPath,
            bool outputToStatsFile)
        {
            try
            {
                var outputOptions = new OutputOptions(outputToStatsFile, "\t");

                try
                {
                    outputOptions.SourceFile = Path.GetFileName(mFastaFilePath);
                }
                catch (Exception)
                {
                    outputOptions.SourceFile = "Unknown_filename_due_to_error.fasta";
                }

                if (outputToStatsFile)
                {
                    mStatsFilePath = ConstructStatsFilePath(outputDirectoryPath);
                    var fileAlreadyExists = File.Exists(mStatsFilePath);

                    var success = false;
                    var retryCount = 0;

                    while (!success && retryCount < 5)
                    {
                        try
                        {
                            outputOptions.OutFile = new StreamWriter(new FileStream(mStatsFilePath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite));
                            success = true;
                        }
                        catch (Exception)
                        {
                            // Failed to open file, wait 1 second, then try again
                            retryCount++;
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
                        SetLocalErrorCode(ValidateFastaFileErrorCodes.ErrorCreatingStatsFile);
                    }
                }
                else
                {
                    outputOptions.SepChar = ", ";
                }

                ReportResultAddEntry(
                    outputOptions, MsgTypeConstants.StatusMsg,
                    "Full path to file", mFastaFilePath);

                ReportResultAddEntry(
                    outputOptions, MsgTypeConstants.StatusMsg,
                    "Protein count", ProteinCount.ToString("#,##0"));

                ReportResultAddEntry(
                    outputOptions, MsgTypeConstants.StatusMsg,
                    "Residue count", ResidueCount.ToString("#,##0"));

                string proteinName;
                if (mFileErrors.Messages.Count > 0)
                {
                    ReportResultAddEntry(
                        outputOptions, MsgTypeConstants.ErrorMsg,
                        "Error count", GetErrorWarningCounts(MsgTypeConstants.ErrorMsg, ErrorWarningCountTypes.Total).ToString());

                    if (mFileErrors.Messages.Count > 1)
                    {
                        mFileErrors.Messages.Sort();
                    }

                    for (var index = 0; index <= mFileErrors.Messages.Count - 1; index++)
                    {
                        var fileError = mFileErrors.Messages[index];
                        if (string.IsNullOrEmpty(fileError.ProteinName))
                        {
                            proteinName = "N/A";
                        }
                        else
                        {
                            proteinName = string.Copy(fileError.ProteinName);
                        }

                        var messageDescription = LookupMessageDescription(fileError.MessageCode, fileError.ExtraInfo);

                        ReportResultAddEntry(
                            outputOptions, MsgTypeConstants.ErrorMsg,
                            fileError.LineNumber,
                            fileError.ColNumber,
                            proteinName,
                            messageDescription,
                            fileError.Context);
                    }
                }

                if (mFileWarnings.Messages.Count > 0)
                {
                    ReportResultAddEntry(
                        outputOptions, MsgTypeConstants.WarningMsg,
                        "Warning count",
                        GetErrorWarningCounts(MsgTypeConstants.WarningMsg, ErrorWarningCountTypes.Total).ToString());

                    if (mFileWarnings.Messages.Count > 1)
                    {
                        mFileWarnings.Messages.Sort();
                    }

                    for (var index = 0; index <= mFileWarnings.Messages.Count - 1; index++)
                    {
                        var fileWarning = mFileWarnings.Messages[index];
                        if (string.IsNullOrEmpty(fileWarning.ProteinName))
                        {
                            proteinName = "N/A";
                        }
                        else
                        {
                            proteinName = string.Copy(fileWarning.ProteinName);
                        }

                        ReportResultAddEntry(
                            outputOptions, MsgTypeConstants.WarningMsg,
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
                    outputOptions, MsgTypeConstants.StatusMsg,
                    "Summary line",
                    ProteinCount.ToString() + " proteins, " + ResidueCount.ToString() + " residues, " + (fastaFile.Length / 1024.0).ToString("0") + " KB");

                if (outputToStatsFile)
                {
                    outputOptions.OutFile?.Close();
                }
            }
            catch (Exception ex)
            {
                OnErrorEvent("Error in ReportResults", ex);
            }
        }

        private void ReportResultAddEntry(
            OutputOptions outputOptions,
            MsgTypeConstants entryType,
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
            OutputOptions outputOptions,
            MsgTypeConstants entryType,
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

            if (!string.IsNullOrEmpty(context))
            {
                dataColumns.Add(context);
            }

            var message = string.Join(outputOptions.SepChar, dataColumns);

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

            LineCount = 0;
            ProteinCount = 0;
            ResidueCount = 0;

            mFixedFastaStats.Reset();
            mFileErrors.Reset();
            mFileWarnings.Reset();

            AbortProcessing = false;
        }

        private void SaveRulesToParameterFile(XmlSettingsFileAccessor settingsFile, string sectionName, IList<RuleDefinition> rules)
        {
            if (rules == null || rules.Count == 0)
            {
                settingsFile.SetParam(sectionName, XML_OPTION_ENTRY_RULE_COUNT, 0);
            }
            else
            {
                settingsFile.SetParam(sectionName, XML_OPTION_ENTRY_RULE_COUNT, rules.Count);

                for (var ruleNumber = 1; ruleNumber <= rules.Count; ruleNumber++)
                {
                    var ruleBase = "Rule" + ruleNumber.ToString();

                    var rule = rules[ruleNumber - 1];
                    settingsFile.SetParam(sectionName, ruleBase + "MatchRegEx", rule.MatchRegEx);
                    settingsFile.SetParam(sectionName, ruleBase + "MatchIndicatesProblem", rule.MatchIndicatesProblem);
                    settingsFile.SetParam(sectionName, ruleBase + "MessageWhenProblem", rule.MessageWhenProblem);
                    settingsFile.SetParam(sectionName, ruleBase + "Severity", rule.Severity);
                    settingsFile.SetParam(sectionName, ruleBase + "DisplayMatchAsExtraInfo", rule.DisplayMatchAsExtraInfo);
                }
            }
        }

        /// <summary>
        /// Save settings to a parameter file
        /// </summary>
        /// <param name="parameterFilePath"></param>
        public bool SaveSettingsToParameterFile(string parameterFilePath)
        {
            // Save a model parameter file

            var settingsFile = new XmlSettingsFileAccessor();

            try
            {
                if (string.IsNullOrEmpty(parameterFilePath))
                {
                    // No parameter file specified; do not save the settings
                    return true;
                }

                if (!File.Exists(parameterFilePath))
                {
                    // Need to generate a blank XML settings file

                    using var writer = new StreamWriter(parameterFilePath, false);

                    writer.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
                    writer.WriteLine("<sections>");
                    writer.WriteLine("  <section name=\"" + XML_SECTION_OPTIONS + "\">");
                    writer.WriteLine("  </section>");
                    writer.WriteLine("</sections>");
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
            IList<RuleDefinition> rules,
            int errorMessageCode,
            out string message)
        {
            if (rules != null)
            {
                for (var index = 0; index <= rules.Count - 1; index++)
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
            SetRule(RuleTypes.HeaderLine, @"^>[^ \t]+\xA0", true, "Non-breaking space after the protein name", DEFAULT_WARNING_SEVERITY);

            // Protein Name error characters
            var allowedChars = @"A-Za-z0-9.\-_:,\|/()\[\]\=\+#";

            if (mAllowAllSymbolsInProteinNames)
            {
                allowedChars += @"!@$%^&*<>?,\\";
            }

            var allowedCharsMatchSpec = "[^" + allowedChars + "]";

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

        private void SetLocalErrorCode(
            ValidateFastaFileErrorCodes newErrorCode,
            bool leaveExistingErrorCodeUnchanged = false)
        {
            if (leaveExistingErrorCodeUnchanged && mLocalErrorCode != ValidateFastaFileErrorCodes.NoError)
            {
                // An error code is already defined; do not change it
            }
            else
            {
                mLocalErrorCode = newErrorCode;

                if (newErrorCode == ValidateFastaFileErrorCodes.NoError)
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
            short severityLevel,
            bool displayMatchAsExtraInfo = false)
        {
            switch (ruleType)
            {
                case RuleTypes.HeaderLine:
                    SetRule(mHeaderLineRules, regexToMatch, doesMatchIndicateProblem, problemReturnMessage, severityLevel, displayMatchAsExtraInfo);
                    break;
                case RuleTypes.ProteinDescription:
                    SetRule(mProteinDescriptionRules, regexToMatch, doesMatchIndicateProblem, problemReturnMessage, severityLevel, displayMatchAsExtraInfo);
                    break;
                case RuleTypes.ProteinName:
                    SetRule(mProteinNameRules, regexToMatch, doesMatchIndicateProblem, problemReturnMessage, severityLevel, displayMatchAsExtraInfo);
                    break;
                case RuleTypes.ProteinSequence:
                    SetRule(mProteinSequenceRules, regexToMatch, doesMatchIndicateProblem, problemReturnMessage, severityLevel, displayMatchAsExtraInfo);
                    break;
            }
        }

        private void SetRule(
            List<RuleDefinition> rules,
            string matchRegEx,
            bool matchIndicatesProblem,
            string messageWhenProblem,
            short severity,
            bool displayMatchAsExtraInfo)
        {
            if (rules == null)
            {
                throw new ArgumentNullException(nameof(rules));
            }

            rules.Add(new RuleDefinition(matchRegEx)
            {
                MatchIndicatesProblem = matchIndicatesProblem,
                MessageWhenProblem = messageWhenProblem,
                Severity = severity,
                DisplayMatchAsExtraInfo = displayMatchAsExtraInfo,
                CustomRuleID = mMasterCustomRuleID,
            });

            mMasterCustomRuleID++;
        }

        private bool SortFile(FileInfo proteinHashFile, int sortColumnIndex, string sortedFilePath)
        {
            var sortUtility = new FlexibleFileSortUtility.TextFileSorter();

            mSortUtilityErrorMessage = string.Empty;

            if (proteinHashFile.Directory != null)
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
            sortUtility.ErrorEvent += SortUtility_ErrorEvent;

            var success = sortUtility.SortFile(proteinHashFile.FullName, sortedFilePath);

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
            var spaceIndex = GetBestSpaceIndex(headerLine);

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

        private bool VerifyLinefeedAtEOF(string inputFilePath, bool addCrLfIfMissing)
        {
            try
            {
                // Open the input file and validate that the final characters are CrLf, simply CR, or simply LF
                using var fsInFile = new FileStream(inputFilePath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);

                if (fsInFile.Length > 2)
                {
                    fsInFile.Seek(-1, SeekOrigin.End);

                    var lastByte = fsInFile.ReadByte();

                    if (lastByte == 10 || lastByte == 13)
                    {
                        // File ends in a linefeed or carriage return character; that's good
                        return true;
                    }
                }

                if (!addCrLfIfMissing)
                    return true;

                ShowMessage("Appending CrLf return to: " + Path.GetFileName(inputFilePath));
                fsInFile.WriteByte(13);

                fsInFile.WriteByte(10);

                return true;
            }
            catch (Exception ex)
            {
                ShowErrorMessage("Error in VerifyLinefeedAtEOF: " + ex.Message);
                SetLocalErrorCode(ValidateFastaFileErrorCodes.ErrorVerifyingLinefeedAtEOF);
                return false;
            }
        }

        private readonly Regex reAdditionalProtein = new(@"(.+)-[a-z]\d*", RegexOptions.Compiled);

        private void WriteCachedProtein(
            string cachedProteinName,
            string cachedProteinDescription,
            TextWriter consolidatedFastaWriter,
            IList<ProteinHashInfo> proteinSeqHashInfo,
            StringBuilder sbCachedProteinResidueLines,
            StringBuilder sbCachedProteinResidues,
            bool consolidateDuplicateProteinSeqsInFasta,
            bool consolidateDupsIgnoreILDiff,
            NestedStringDictionary<int> proteinNameFirst,
            NestedStringDictionary<string> duplicateProteinList,
            int lineCountRead,
            NestedStringDictionary<int> proteinsWritten)
        {
            var lineOut = string.Empty;

            bool keepProtein;

            var additionalProteinNames = new List<string>();

            if (proteinNameFirst.TryGetValue(cachedProteinName, out var seqIndex))
            {
                // cachedProteinName was found in proteinNameFirst

                if (proteinsWritten.TryGetValue(cachedProteinName, out _))
                {
                    if (mFixedFastaOptions.KeepDuplicateNamedProteinsUnlessMatchingSequence)
                    {
                        // Keep this protein if its sequence hash differs from the first protein with this name
                        var proteinHash = ComputeProteinHash(sbCachedProteinResidues, consolidateDupsIgnoreILDiff);
                        if ((proteinSeqHashInfo[seqIndex].SequenceHash ?? "") != (proteinHash ?? ""))
                        {
                            RecordFastaFileWarning(lineCountRead, 1, cachedProteinName, (int)MessageCodeConstants.DuplicateProteinNameRetained);
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
                    if (proteinSeqHashInfo[seqIndex].AdditionalProteins.Any())
                    {
                        // The protein has duplicate proteins
                        // Construct a list of the duplicate protein names

                        additionalProteinNames.Clear();
                        foreach (var additionalProtein in proteinSeqHashInfo[seqIndex].AdditionalProteins)
                        {
                            // Add the additional protein name if it is not of the form "BaseName-b", "BaseName-c", etc.
                            var skipDupProtein = false;

                            if (additionalProtein == null)
                            {
                                skipDupProtein = true;
                            }
                            else if (string.Equals(additionalProtein, cachedProteinName, StringComparison.OrdinalIgnoreCase))
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
                                var match = reAdditionalProtein.Match(additionalProtein);

                                if (match.Success)
                                {
                                    if (string.Equals(cachedProteinName, match.Groups[1].Value, StringComparison.OrdinalIgnoreCase))
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
                            var updatedDescription = cachedProteinDescription + "; Duplicate proteins: " + FlattenArray(additionalProteinNames, ',');
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
                mFixedFastaStats.DuplicateSequenceProteinsSkipped++;

                string masterProteinInfo;
                if (!duplicateProteinList.TryGetValue(cachedProteinName, out var masterProteinName))
                {
                    masterProteinInfo = "same as ??";
                }
                else
                {
                    masterProteinInfo = "same as " + masterProteinName;
                }

                RecordFastaFileWarning(lineCountRead, 0, cachedProteinName, (int)MessageCodeConstants.ProteinRemovedSinceDuplicateSequence, masterProteinInfo, string.Empty);
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

            var firstProtein = proteinNames.ElementAtOrDefault(0);
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

                for (var proteinIndex = 1; proteinIndex <= proteinNames.Count - 1; proteinIndex++)
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

        private void SortUtility_ErrorEvent(string message, Exception ex)
        {
            mSortUtilityErrorMessage = message;
        }

        #endregion
    }
}