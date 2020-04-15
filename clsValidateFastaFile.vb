Option Strict On

' This class will read a protein fasta file and validate its contents
'
' -------------------------------------------------------------------------------
' Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
' Program started March 21, 2005

' E-mail: matthew.monroe@pnnl.gov or proteomics@pnnl.gov
' Website: https://omics.pnl.gov/ or https://panomics.pnnl.gov/
' -------------------------------------------------------------------------------
'
' Licensed under the Apache License, Version 2.0; you may not use this file except
' in compliance with the License.  You may obtain a copy of the License at
' http://www.apache.org/licenses/LICENSE-2.0

Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports PRISM

Public Class clsValidateFastaFile
    Inherits FileProcessor.ProcessFilesBase

    ''' <summary>
    ''' Constructor
    ''' </summary>
    Public Sub New()
        mFileDate = "April 15, 2020"
        InitializeLocalVariables()
    End Sub

    ''' <summary>
    ''' Constructor that takes a parameter file
    ''' </summary>
    ''' <param name="parameterFilePath"></param>
    Public Sub New(parameterFilePath As String)
        Me.New()
        LoadParameterFileSettings(parameterFilePath)
    End Sub

#Region "Constants and Enums"
    Private Const DEFAULT_MINIMUM_PROTEIN_NAME_LENGTH As Integer = 3

    ''' <summary>
    ''' The maximum suggested value when using SEQUEST is 34 characters
    ''' In contrast, MS-GF+ supports long protein names
    ''' </summary>
    ''' <remarks></remarks>
    Public Const DEFAULT_MAXIMUM_PROTEIN_NAME_LENGTH As Integer = 60
    Private Const DEFAULT_MAXIMUM_RESIDUES_PER_LINE As Integer = 120

    Public Const DEFAULT_PROTEIN_LINE_START_CHAR As Char = ">"c
    Public Const DEFAULT_LONG_PROTEIN_NAME_SPLIT_CHAR As Char = "|"c
    Public Const DEFAULT_PROTEIN_NAME_FIRST_REF_SEP_CHARS As String = ":|"
    Public Const DEFAULT_PROTEIN_NAME_SUBSEQUENT_REF_SEP_CHARS As String = ":|;"

    Private Const INVALID_PROTEIN_NAME_CHAR_REPLACEMENT As Char = "_"c

    Private Const CUSTOM_RULE_ID_START As Integer = 1000
    Private Const DEFAULT_CONTEXT_LENGTH As Integer = 13

    Public Const MESSAGE_TEXT_PROTEIN_DESCRIPTION_MISSING As String = "Line contains a protein name, but not a description"
    Public Const MESSAGE_TEXT_PROTEIN_DESCRIPTION_TOO_LONG As String = "Protein description is over 900 characters long"
    Public Const MESSAGE_TEXT_ASTERISK_IN_RESIDUES As String = "An asterisk was found in the residues"
    Public Const MESSAGE_TEXT_DASH_IN_RESIDUES As String = "A dash was found in the residues"

    Public Const XML_SECTION_OPTIONS As String = "ValidateFastaFileOptions"
    Public Const XML_SECTION_FIXED_FASTA_FILE_OPTIONS As String = "ValidateFastaFixedFASTAFileOptions"

    Public Const XML_SECTION_FASTA_HEADER_LINE_RULES As String = "ValidateFastaHeaderLineRules"
    Public Const XML_SECTION_FASTA_PROTEIN_NAME_RULES As String = "ValidateFastaProteinNameRules"
    Public Const XML_SECTION_FASTA_PROTEIN_DESCRIPTION_RULES As String = "ValidateFastaProteinDescriptionRules"
    Public Const XML_SECTION_FASTA_PROTEIN_SEQUENCE_RULES As String = "ValidateFastaProteinSequenceRules"

    Public Const XML_OPTION_ENTRY_RULE_COUNT As String = "RuleCount"

    ' The value of 7995 is chosen because the maximum varchar() value in Sql Server is varchar(8000)
    ' and we want to prevent truncation errors when importing protein names and descriptions into Sql Server
    Public Const MAX_PROTEIN_DESCRIPTION_LENGTH As Integer = 7995

    Private Const MEM_USAGE_PREFIX = "MemUsage: "
    Private Const REPORT_DETAILED_MEMORY_USAGE = False

    Private Const PROTEIN_NAME_COLUMN = "Protein_Name"
    Private Const SEQUENCE_LENGTH_COLUMN = "Sequence_Length"
    Private Const SEQUENCE_HASH_COLUMN = "Sequence_Hash"
    Private Const PROTEIN_HASHES_FILENAME_SUFFIX = "_ProteinHashes.txt"

    Private Const DEFAULT_WARNING_SEVERITY = 3
    Private Const DEFAULT_ERROR_SEVERITY = 7

    ' Note: Custom rules start with message code CUSTOM_RULE_ID_START=1000, and therefore
    ' the values in enum eMessageCodeConstants should all be less than CUSTOM_RULE_ID_START
    Public Enum eMessageCodeConstants
        UnspecifiedError = 0

        ' Error messages
        ProteinNameIsTooLong = 1
        LineStartsWithSpace = 2
        '        RightArrowFollowedBySpace = 3
        '        RightArrowFollowedByTab = 4
        '        RightArrowButNoProteinName = 5
        BlankLineBetweenProteinNameAndResidues = 6
        BlankLineInMiddleOfResidues = 7
        ResiduesFoundWithoutProteinHeader = 8
        ProteinEntriesNotFound = 9
        FinalProteinEntryMissingResidues = 10
        FileDoesNotEndWithLinefeed = 11
        DuplicateProteinName = 12

        ' Warning messages
        ProteinNameIsTooShort = 13
        '        ProteinNameContainsVerticalBars = 14
        '        ProteinNameContainsWarningCharacters = 21
        '        ProteinNameWithoutDescription = 14
        BlankLineBeforeProteinName = 15
        '        ProteinNameAndDescriptionSeparatedByTab = 16
        '        ProteinDescriptionWithTab = 25
        '        ProteinDescriptionWithQuotationMark = 26
        '        ProteinDescriptionWithEscapedSlash = 27
        '        ProteinDescriptionWithUndesirableCharacter = 28
        ResiduesLineTooLong = 17
        '        ResiduesLineContainsU = 30
        DuplicateProteinSequence = 18
        RenamedProtein = 19
        ProteinRemovedSinceDuplicateSequence = 20
        DuplicateProteinNameRetained = 21
    End Enum

    Structure udtMsgInfoType
        Public LineNumber As Integer
        Public ColNumber As Integer
        Public ProteinName As String
        Public MessageCode As Integer
        Public ExtraInfo As String
        Public Context As String

        Public Overrides Function ToString() As String
            Return String.Format("Line {0}, protein {1}, code {2}: {3}", LineNumber, ProteinName, MessageCode, ExtraInfo)
        End Function
    End Structure

    Structure udtOutputOptionsType
        Public SourceFile As String
        Public OutputToStatsFile As Boolean
        Public OutFile As StreamWriter
        Public SepChar As String

        Public Overrides Function ToString() As String
            Return SourceFile
        End Function
    End Structure

    Enum RuleTypes
        HeaderLine
        ProteinName
        ProteinDescription
        ProteinSequence
    End Enum

    Enum SwitchOptions
        AddMissingLineFeedAtEOF
        AllowAsteriskInResidues
        CheckForDuplicateProteinNames
        GenerateFixedFASTAFile
        SplitOutMultipleRefsInProteinName
        OutputToStatsFile
        WarnBlankLinesBetweenProteins
        WarnLineStartsWithSpace
        NormalizeFileLineEndCharacters
        CheckForDuplicateProteinSequences
        FixedFastaRenameDuplicateNameProteins
        SaveProteinSequenceHashInfoFiles
        FixedFastaConsolidateDuplicateProteinSeqs
        FixedFastaConsolidateDupsIgnoreILDiff
        FixedFastaTruncateLongProteinNames
        FixedFastaSplitOutMultipleRefsForKnownAccession
        FixedFastaWrapLongResidueLines
        FixedFastaRemoveInvalidResidues
        SaveBasicProteinHashInfoFile
        AllowDashInResidues
        FixedFastaKeepDuplicateNamedProteins        ' Keep duplicate named proteins, unless the name and sequence match exactly, then they're removed
        AllowAllSymbolsInProteinNames
    End Enum

    Enum FixedFASTAFileValues
        DuplicateProteinNamesSkippedCount
        ProteinNamesInvalidCharsReplaced
        ProteinNamesMultipleRefsRemoved
        TruncatedProteinNameCount
        UpdatedResidueLines
        DuplicateProteinNamesRenamedCount
        DuplicateProteinSeqsSkippedCount
    End Enum

    Enum ErrorWarningCountTypes
        Specified
        Unspecified
        Total
    End Enum

    Enum eMsgTypeConstants
        ErrorMsg = 0
        WarningMsg = 1
        StatusMsg = 2
    End Enum

    Enum eValidateFastaFileErrorCodes
        NoError = 0
        OptionsSectionNotFound = 1
        ErrorReadingInputFile = 2
        ErrorCreatingStatsFile = 4
        ErrorVerifyingLinefeedAtEOF = 8
        UnspecifiedError = -1
    End Enum

    Enum eLineEndingCharacters
        CRLF  'Windows
        CR    'Old Style Mac
        LF    'Unix/Linux/OS X
        LFCR  'Oddball (Just for completeness!)
    End Enum
#End Region

#Region "Structures"

    Private Structure udtErrorStatsType
        Public MessageCode As Integer               ' Note: Custom rules start with message code CUSTOM_RULE_ID_START
        Public CountSpecified As Integer
        Public CountUnspecified As Integer

        Public Overrides Function ToString() As String
            Return MessageCode & ": " & CountSpecified & " specified, " & CountUnspecified & " unspecified"
        End Function
    End Structure

    Private Structure udtItemSummaryIndexedType
        Public ErrorStatsCount As Integer
        Public ErrorStats() As udtErrorStatsType        ' Note: This array ranges from 0 to .ErrorStatsCount since it is Dimmed with extra space
        Public MessageCodeToArrayIndex As Dictionary(Of Integer, Integer)
    End Structure

    Public Structure udtRuleDefinitionType
        Public MatchRegEx As String
        Public MatchIndicatesProblem As Boolean     ' True means text matching the RegEx means a problem; false means if text doesn't match the RegEx, then that means a problem
        Public MessageWhenProblem As String         ' Message to display if a problem is present
        Public Severity As Short                    ' 0 is lowest severity, 9 is highest severity; value >= 5 means error
        Public DisplayMatchAsExtraInfo As Boolean   ' If true, then the matching text is stored as the context info
        Public CustomRuleID As Integer              ' This value is auto-assigned

        Public Overrides Function ToString() As String
            Return CustomRuleID & ": " & MessageWhenProblem
        End Function
    End Structure

    Private Structure udtRuleDefinitionExtendedType
        Public RuleDefinition As udtRuleDefinitionType
        Public reRule As Regex

        ' ReSharper disable once NotAccessedField.Local
        Public Valid As Boolean

        Public Overrides Function ToString() As String
            Return RuleDefinition.CustomRuleID & ": " & RuleDefinition.MessageWhenProblem
        End Function
    End Structure

    Private Structure udtFixedFastaOptionsType
        Public SplitOutMultipleRefsInProteinName As Boolean
        Public SplitOutMultipleRefsForKnownAccession As Boolean
        Public LongProteinNameSplitChars As Char()
        Public ProteinNameInvalidCharsToRemove As Char()
        Public RenameProteinsWithDuplicateNames As Boolean
        Public KeepDuplicateNamedProteinsUnlessMatchingSequence As Boolean      ' Ignored if RenameProteinsWithDuplicateNames=true or ConsolidateProteinsWithDuplicateSeqs=true
        Public ConsolidateProteinsWithDuplicateSeqs As Boolean
        Public ConsolidateDupsIgnoreILDiff As Boolean
        Public TruncateLongProteinNames As Boolean
        Public WrapLongResidueLines As Boolean
        Public RemoveInvalidResidues As Boolean
    End Structure

    Private Structure udtFixedFastaStatsType
        Public TruncatedProteinNameCount As Integer
        Public UpdatedResidueLines As Integer
        Public ProteinNamesInvalidCharsReplaced As Integer
        Public ProteinNamesMultipleRefsRemoved As Integer
        Public DuplicateNameProteinsSkipped As Integer
        Public DuplicateNameProteinsRenamed As Integer
        Public DuplicateSequenceProteinsSkipped As Integer
    End Structure

    Private Structure udtProteinNameTruncationRegex
        Public reMatchIPI As Regex
        Public reMatchGI As Regex
        Public reMatchJGI As Regex
        Public reMatchJGIBaseAndID As Regex
        Public reMatchGeneric As Regex
        Public reMatchDoubleBarOrColonAndBar As Regex
    End Structure

#End Region

#Region "Classwide Variables"

    ''' <summary>
    ''' Fasta file path being examined
    ''' </summary>
    ''' <remarks>Used by clsCustomValidateFastaFiles</remarks>
    Protected mFastaFilePath As String

    Private mLineCount As Integer
    Private mProteinCount As Integer
    Private mResidueCount As Long

    Private mFixedFastaStats As udtFixedFastaStatsType

    Private mFileErrorCount As Integer
    Private mFileErrors() As udtMsgInfoType
    Private mFileErrorStats As udtItemSummaryIndexedType

    Private mFileWarningCount As Integer

    Private mFileWarnings() As udtMsgInfoType
    Private mFileWarningStats As udtItemSummaryIndexedType

    Private mHeaderLineRules() As udtRuleDefinitionType
    Private mProteinNameRules() As udtRuleDefinitionType
    Private mProteinDescriptionRules() As udtRuleDefinitionType
    Private mProteinSequenceRules() As udtRuleDefinitionType
    Private mMasterCustomRuleID As Integer = CUSTOM_RULE_ID_START

    Private mProteinNameFirstRefSepChars() As Char
    Private mProteinNameSubsequentRefSepChars() As Char

    Private mAddMissingLinefeedAtEOF As Boolean
    Private mCheckForDuplicateProteinNames As Boolean

    ' This will be set to True if
    '   mSaveProteinSequenceHashInfoFiles = True or mFixedFastaOptions.ConsolidateProteinsWithDuplicateSeqs = True
    Private mCheckForDuplicateProteinSequences As Boolean

    Private mMaximumFileErrorsToTrack As Integer        ' This is the maximum # of errors per type to track
    Private mMinimumProteinNameLength As Integer
    Private mMaximumProteinNameLength As Integer
    Private mMaximumResiduesPerLine As Integer

    Private mFixedFastaOptions As udtFixedFastaOptionsType     ' Used if mGenerateFixedFastaFile = True

    Private mOutputToStatsFile As Boolean
    Private mStatsFilePath As String

    Private mGenerateFixedFastaFile As Boolean
    Private mSaveProteinSequenceHashInfoFiles As Boolean

    ' When true, creates a text file that will contain the protein name and sequence hash for each protein;
    '  this option will not store protein names and/or hashes in memory, and is thus useful for processing
    '  huge .Fasta files to determine duplicate proteins
    Private mSaveBasicProteinHashInfoFile As Boolean

    Private mProteinLineStartChar As Char

    Private mAllowAsteriskInResidues As Boolean
    Private mAllowDashInResidues As Boolean
    Private mAllowAllSymbolsInProteinNames As Boolean

    Private mWarnBlankLinesBetweenProteins As Boolean
    Private mWarnLineStartsWithSpace As Boolean
    Private mNormalizeFileLineEndCharacters As Boolean

    ''' <summary>
    ''' The number of characters at the start of key strings to use when adding items to clsNestedStringDictionary instances
    ''' </summary>
    ''' <remarks>
    ''' If this value is too short, all of the items added to the clsNestedStringDictionary instance
    ''' will be tracked by the same dictionary, which could result in a dictionary surpassing the 2 GB boundary
    ''' </remarks>
    Private mProteinNameSpannerCharLength As Byte = 1

    Private mLocalErrorCode As eValidateFastaFileErrorCodes

    Private mMemoryUsageLogger As clsMemoryUsageLogger

    Private mProcessMemoryUsageMBAtStart As Single

    Private mSortUtilityErrorMessage As String

    Private mTempFilesToDelete As List(Of String)

#End Region

#Region "Properties"

    ''' <summary>
    ''' Gets or sets a processing option
    ''' </summary>
    ''' <param name="SwitchName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Be sure to call SetDefaultRules() after setting all of the options</remarks>
    Public Property OptionSwitch(switchName As SwitchOptions) As Boolean
        Get
            Return GetOptionSwitchValue(switchName)
        End Get
        Set
            SetOptionSwitch(switchName, Value)
        End Set
    End Property

    ''' <summary>
    ''' Set a processing option
    ''' </summary>
    ''' <param name="SwitchName"></param>
    ''' <param name="State"></param>
    ''' <remarks>Be sure to call SetDefaultRules() after setting all of the options</remarks>
    Public Sub SetOptionSwitch(switchName As SwitchOptions, state As Boolean)

        Select Case switchName
            Case SwitchOptions.AddMissingLinefeedatEOF
                mAddMissingLinefeedAtEOF = state
            Case SwitchOptions.AllowAsteriskInResidues
                mAllowAsteriskInResidues = state
            Case SwitchOptions.CheckForDuplicateProteinNames
                mCheckForDuplicateProteinNames = state
            Case SwitchOptions.GenerateFixedFASTAFile
                mGenerateFixedFastaFile = state
            Case SwitchOptions.OutputToStatsFile
                mOutputToStatsFile = state
            Case SwitchOptions.SplitOutMultipleRefsInProteinName
                mFixedFastaOptions.SplitOutMultipleRefsInProteinName = state
            Case SwitchOptions.WarnBlankLinesBetweenProteins
                mWarnBlankLinesBetweenProteins = state
            Case SwitchOptions.WarnLineStartsWithSpace
                mWarnLineStartsWithSpace = state
            Case SwitchOptions.NormalizeFileLineEndCharacters
                mNormalizeFileLineEndCharacters = state
            Case SwitchOptions.CheckForDuplicateProteinSequences
                mCheckForDuplicateProteinSequences = state
            Case SwitchOptions.FixedFastaRenameDuplicateNameProteins
                mFixedFastaOptions.RenameProteinsWithDuplicateNames = state
            Case SwitchOptions.FixedFastaKeepDuplicateNamedProteins
                mFixedFastaOptions.KeepDuplicateNamedProteinsUnlessMatchingSequence = state
            Case SwitchOptions.SaveProteinSequenceHashInfoFiles
                mSaveProteinSequenceHashInfoFiles = state
            Case SwitchOptions.FixedFastaConsolidateDuplicateProteinSeqs
                mFixedFastaOptions.ConsolidateProteinsWithDuplicateSeqs = state
            Case SwitchOptions.FixedFastaConsolidateDupsIgnoreILDiff
                mFixedFastaOptions.ConsolidateDupsIgnoreILDiff = state
            Case SwitchOptions.FixedFastaTruncateLongProteinNames
                mFixedFastaOptions.TruncateLongProteinNames = state
            Case SwitchOptions.FixedFastaSplitOutMultipleRefsForKnownAccession
                mFixedFastaOptions.SplitOutMultipleRefsForKnownAccession = state
            Case SwitchOptions.FixedFastaWrapLongResidueLines
                mFixedFastaOptions.WrapLongResidueLines = state
            Case SwitchOptions.FixedFastaRemoveInvalidResidues
                mFixedFastaOptions.RemoveInvalidResidues = state
            Case SwitchOptions.SaveBasicProteinHashInfoFile
                mSaveBasicProteinHashInfoFile = state
            Case SwitchOptions.AllowDashInResidues
                mAllowDashInResidues = state
            Case SwitchOptions.AllowAllSymbolsInProteinNames
                mAllowAllSymbolsInProteinNames = state
        End Select

    End Sub

    Public Function GetOptionSwitchValue(SwitchName As SwitchOptions) As Boolean

        Select Case SwitchName
            Case SwitchOptions.AddMissingLinefeedatEOF
                Return mAddMissingLinefeedAtEOF
            Case SwitchOptions.AllowAsteriskInResidues
                Return mAllowAsteriskInResidues
            Case SwitchOptions.CheckForDuplicateProteinNames
                Return mCheckForDuplicateProteinNames
            Case SwitchOptions.GenerateFixedFASTAFile
                Return mGenerateFixedFastaFile
            Case SwitchOptions.OutputToStatsFile
                Return mOutputToStatsFile
            Case SwitchOptions.SplitOutMultipleRefsInProteinName
                Return mFixedFastaOptions.SplitOutMultipleRefsInProteinName
            Case SwitchOptions.WarnBlankLinesBetweenProteins
                Return mWarnBlankLinesBetweenProteins
            Case SwitchOptions.WarnLineStartsWithSpace
                Return mWarnLineStartsWithSpace
            Case SwitchOptions.NormalizeFileLineEndCharacters
                Return mNormalizeFileLineEndCharacters
            Case SwitchOptions.CheckForDuplicateProteinSequences
                Return mCheckForDuplicateProteinSequences
            Case SwitchOptions.FixedFastaRenameDuplicateNameProteins
                Return mFixedFastaOptions.RenameProteinsWithDuplicateNames
            Case SwitchOptions.FixedFastaKeepDuplicateNamedProteins
                Return mFixedFastaOptions.KeepDuplicateNamedProteinsUnlessMatchingSequence
            Case SwitchOptions.SaveProteinSequenceHashInfoFiles
                Return mSaveProteinSequenceHashInfoFiles
            Case SwitchOptions.FixedFastaConsolidateDuplicateProteinSeqs
                Return mFixedFastaOptions.ConsolidateProteinsWithDuplicateSeqs
            Case SwitchOptions.FixedFastaConsolidateDupsIgnoreILDiff
                Return mFixedFastaOptions.ConsolidateDupsIgnoreILDiff
            Case SwitchOptions.FixedFastaTruncateLongProteinNames
                Return mFixedFastaOptions.TruncateLongProteinNames
            Case SwitchOptions.FixedFastaSplitOutMultipleRefsForKnownAccession
                Return mFixedFastaOptions.SplitOutMultipleRefsForKnownAccession
            Case SwitchOptions.FixedFastaWrapLongResidueLines
                Return mFixedFastaOptions.WrapLongResidueLines
            Case SwitchOptions.FixedFastaRemoveInvalidResidues
                Return mFixedFastaOptions.RemoveInvalidResidues
            Case SwitchOptions.SaveBasicProteinHashInfoFile
                Return mSaveBasicProteinHashInfoFile
            Case SwitchOptions.AllowDashInResidues
                Return mAllowDashInResidues
            Case SwitchOptions.AllowAllSymbolsInProteinNames
                Return mAllowAllSymbolsInProteinNames
        End Select

        Return False

    End Function

    Public ReadOnly Property ErrorWarningCounts(
      messageType As eMsgTypeConstants,
      CountType As ErrorWarningCountTypes) As Integer

        Get
            Dim tmpValue As Integer
            Select Case CountType
                Case ErrorWarningCountTypes.Total
                    Select Case messageType
                        Case eMsgTypeConstants.ErrorMsg
                            tmpValue = mFileErrorCount + ComputeTotalUnspecifiedCount(mFileErrorStats)
                        Case eMsgTypeConstants.WarningMsg
                            tmpValue = mFileWarningCount + ComputeTotalUnspecifiedCount(mFileWarningStats)
                        Case eMsgTypeConstants.StatusMsg
                            tmpValue = 0
                    End Select
                Case ErrorWarningCountTypes.Unspecified
                    Select Case messageType
                        Case eMsgTypeConstants.ErrorMsg
                            tmpValue = ComputeTotalUnspecifiedCount(mFileErrorStats)
                        Case eMsgTypeConstants.WarningMsg
                            tmpValue = ComputeTotalSpecifiedCount(mFileWarningStats)
                        Case eMsgTypeConstants.StatusMsg
                            tmpValue = 0
                    End Select
                Case ErrorWarningCountTypes.Specified
                    Select Case messageType
                        Case eMsgTypeConstants.ErrorMsg
                            tmpValue = mFileErrorCount
                        Case eMsgTypeConstants.WarningMsg
                            tmpValue = mFileWarningCount
                        Case eMsgTypeConstants.StatusMsg
                            tmpValue = 0
                    End Select
            End Select

            Return tmpValue
        End Get
    End Property

    Public ReadOnly Property FixedFASTAFileStats(valueType As FixedFASTAFileValues) As Integer
        Get
            Dim tmpValue As Integer
            Select Case valueType
                Case FixedFASTAFileValues.DuplicateProteinNamesSkippedCount
                    tmpValue = mFixedFastaStats.DuplicateNameProteinsSkipped
                Case FixedFASTAFileValues.ProteinNamesInvalidCharsReplaced
                    tmpValue = mFixedFastaStats.ProteinNamesInvalidCharsReplaced
                Case FixedFASTAFileValues.ProteinNamesMultipleRefsRemoved
                    tmpValue = mFixedFastaStats.ProteinNamesMultipleRefsRemoved
                Case FixedFASTAFileValues.TruncatedProteinNameCount
                    tmpValue = mFixedFastaStats.TruncatedProteinNameCount
                Case FixedFASTAFileValues.UpdatedResidueLines
                    tmpValue = mFixedFastaStats.UpdatedResidueLines
                Case FixedFASTAFileValues.DuplicateProteinNamesRenamedCount
                    tmpValue = mFixedFastaStats.DuplicateNameProteinsRenamed
                Case FixedFASTAFileValues.DuplicateProteinSeqsSkippedCount
                    tmpValue = mFixedFastaStats.DuplicateSequenceProteinsSkipped
            End Select
            Return tmpValue

        End Get
    End Property

    Public ReadOnly Property ProteinCount As Integer
        Get
            Return mProteinCount
        End Get
    End Property

    Public ReadOnly Property LineCount As Integer
        Get
            Return mLineCount
        End Get
    End Property

    Public ReadOnly Property LocalErrorCode As eValidateFastaFileErrorCodes
        Get
            Return mLocalErrorCode
        End Get
    End Property

    Public ReadOnly Property ResidueCount As Long
        Get
            Return mResidueCount
        End Get
    End Property

    Public ReadOnly Property FastaFilePath As String
        Get
            Return mFastaFilePath
        End Get
    End Property

    Public ReadOnly Property ErrorMessageTextByIndex(
      index As Integer,
      valueSeparator As String) As String

        Get
            Return GetFileErrorTextByIndex(index, valueSeparator)
        End Get
    End Property

    Public ReadOnly Property WarningMessageTextByIndex(
      index As Integer,
      valueSeparator As String) As String

        Get
            Return GetFileWarningTextByIndex(index, valueSeparator)
        End Get
    End Property

    Public ReadOnly Property ErrorsByIndex(errorIndex As Integer) As udtMsgInfoType
        Get
            Return (GetFileErrorByIndex(errorIndex))
        End Get
    End Property

    Public ReadOnly Property WarningsByIndex(warningIndex As Integer) As udtMsgInfoType
        Get
            Return GetFileWarningByIndex(warningIndex)
        End Get
    End Property

    ''' <summary>
    ''' Existing protein hash file to load into memory instead of computing new hash values while reading the fasta file
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ExistingProteinHashFile As String

    Public Property MaximumFileErrorsToTrack As Integer
        Get
            Return mMaximumFileErrorsToTrack
        End Get
        Set
            If Value < 1 Then Value = 1
            mMaximumFileErrorsToTrack = Value
        End Set
    End Property

    Public Property MaximumProteinNameLength As Integer
        Get
            Return mMaximumProteinNameLength
        End Get
        Set
            If Value < 8 Then
                ' Do not allow maximum lengths less than 8; use the default
                Value = DEFAULT_MAXIMUM_PROTEIN_NAME_LENGTH
            End If
            mMaximumProteinNameLength = Value
        End Set
    End Property

    Public Property MinimumProteinNameLength As Integer
        Get
            Return mMinimumProteinNameLength
        End Get
        Set
            If Value < 1 Then Value = DEFAULT_MINIMUM_PROTEIN_NAME_LENGTH
            mMinimumProteinNameLength = Value
        End Set
    End Property

    Public Property MaximumResiduesPerLine As Integer
        Get
            Return mMaximumResiduesPerLine
        End Get
        Set
            If Value = 0 Then
                Value = DEFAULT_MAXIMUM_RESIDUES_PER_LINE
            ElseIf Value < 40 Then
                Value = 40
            End If

            mMaximumResiduesPerLine = Value
        End Set
    End Property

    Public Property ProteinLineStartChar As Char
        Get
            Return mProteinLineStartChar
        End Get
        Set
            mProteinLineStartChar = Value
        End Set
    End Property

    Public ReadOnly Property StatsFilePath As String
        Get
            If mStatsFilePath Is Nothing Then
                Return String.Empty
            Else
                Return mStatsFilePath
            End If
        End Get
    End Property

    Public Property ProteinNameInvalidCharsToRemove As String
        Get
            Return CharArrayToString(mFixedFastaOptions.ProteinNameInvalidCharsToRemove)
        End Get
        Set
            If Value Is Nothing Then
                Value = String.Empty
            End If

            ' Check for and remove any spaces from Value, since
            ' a space does not make sense for an invalid protein name character
            Value = Value.Replace(" "c, String.Empty)
            If Value.Length > 0 Then
                mFixedFastaOptions.ProteinNameInvalidCharsToRemove = Value.ToCharArray
            Else
                mFixedFastaOptions.ProteinNameInvalidCharsToRemove = New Char() {}      ' Default to an empty character array if Value is empty
            End If
        End Set
    End Property

    Public Property ProteinNameFirstRefSepChars As String
        Get
            Return CharArrayToString(mProteinNameFirstRefSepChars)
        End Get
        Set
            If Value Is Nothing Then
                Value = String.Empty
            End If

            ' Check for and remove any spaces from Value, since
            ' a space does not make sense for a separation character
            Value = Value.Replace(" "c, String.Empty)
            If Value.Length > 0 Then
                mProteinNameFirstRefSepChars = Value.ToCharArray
            Else
                mProteinNameFirstRefSepChars = DEFAULT_PROTEIN_NAME_FIRST_REF_SEP_CHARS.ToCharArray     ' Use the default if Value is empty
            End If
        End Set
    End Property

    Public Property ProteinNameSubsequentRefSepChars As String
        Get
            Return CharArrayToString(mProteinNameSubsequentRefSepChars)
        End Get
        Set
            If Value Is Nothing Then
                Value = String.Empty
            End If

            ' Check for and remove any spaces from Value, since
            ' a space does not make sense for a separation character
            Value = Value.Replace(" "c, String.Empty)
            If Value.Length > 0 Then
                mProteinNameSubsequentRefSepChars = Value.ToCharArray
            Else
                mProteinNameSubsequentRefSepChars = DEFAULT_PROTEIN_NAME_SUBSEQUENT_REF_SEP_CHARS.ToCharArray     ' Use the default if Value is empty
            End If
        End Set
    End Property

    Public Property LongProteinNameSplitChars As String
        Get
            Return CharArrayToString(mFixedFastaOptions.LongProteinNameSplitChars)
        End Get
        Set
            If Not Value Is Nothing Then
                ' Check for and remove any spaces from Value, since
                ' a space does not make sense for a protein name split char
                Value = Value.Replace(" "c, String.Empty)
                If Value.Length > 0 Then
                    mFixedFastaOptions.LongProteinNameSplitChars = Value.ToCharArray
                End If
            End If
        End Set
    End Property

    Public ReadOnly Property FileWarningList As List(Of udtMsgInfoType)
        Get
            Return GetFileWarnings()
        End Get
    End Property

    Public ReadOnly Property FileErrorList As List(Of udtMsgInfoType)
        Get
            Return GetFileErrors()
        End Get
    End Property

#End Region

    Public Event ProgressCompleted()

    Public Event WroteLineEndNormalizedFASTA(newFilePath As String)

    Private Sub OnProgressComplete() Handles MyBase.ProgressComplete
        RaiseEvent ProgressCompleted()
        OperationComplete()
    End Sub

    Private Sub OnWroteLineEndNormalizedFASTA(newFilePath As String)
        RaiseEvent WroteLineEndNormalizedFASTA(newFilePath)
    End Sub

    ''' <summary>
    ''' Examine the given fasta file to look for problems.
    ''' Optionally create a new, fixed fasta file
    ''' Optionally also consolidate proteins with duplicate sequences
    ''' </summary>
    ''' <param name="fastaFilePathToCheck"></param>
    ''' <param name="preloadedProteinNamesToKeep">
    ''' Preloaded list of protein names to include in the fixed fasta file
    ''' Keys are protein names, values are the number of entries written to the fixed fasta file for the given protein name
    ''' </param>
    ''' <returns>True if the file was successfully analyzed (even if errors were found)</returns>
    ''' <remarks>Assumes fastaFilePathToCheck exists</remarks>
    Private Function AnalyzeFastaFile(fastaFilePathToCheck As String, preloadedProteinNamesToKeep As clsNestedStringIntList) As Boolean

        Dim fixedFastaWriter As StreamWriter = Nothing
        Dim sequenceHashWriter As StreamWriter = Nothing

        Dim fastaFilePathOut = "UndefinedFilePath.xyz"

        Dim success As Boolean
        Dim exceptionCaught = False

        Dim consolidateDuplicateProteinSeqsInFasta = False
        Dim keepDuplicateNamedProteinsUnlessMatchingSequence = False
        Dim consolidateDupsIgnoreILDiff = False

        ' This array tracks protein hash details
        Dim proteinSequenceHashCount As Integer
        Dim proteinSeqHashInfo() As clsProteinHashInfo

        Dim headerLineRuleDetails() As udtRuleDefinitionExtendedType
        Dim proteinNameRuleDetails() As udtRuleDefinitionExtendedType
        Dim proteinDescriptionRuleDetails() As udtRuleDefinitionExtendedType
        Dim proteinSequenceRuleDetails() As udtRuleDefinitionExtendedType

        Try
            ' Reset the data structures and variables
            ResetStructures()
            ReDim proteinSeqHashInfo(0)

            ReDim headerLineRuleDetails(1)
            ReDim proteinNameRuleDetails(1)
            ReDim proteinDescriptionRuleDetails(1)
            ReDim proteinSequenceRuleDetails(1)

            ' This is a dictionary of dictionaries, with one dictionary for each letter or number that a SHA-1 hash could start with
            ' This dictionary of dictionaries provides a quick lookup for existing protein hashes
            ' This dictionary is not used if preloadedProteinNamesToKeep contains data
            Const SPANNER_CHAR_LENGTH = 1
            Dim proteinSequenceHashes = New clsNestedStringDictionary(Of Integer)(False, SPANNER_CHAR_LENGTH)
            Dim usingPreloadedProteinNames = False

            If Not preloadedProteinNamesToKeep Is Nothing AndAlso preloadedProteinNamesToKeep.Count > 0 Then
                ' Auto enable/disable some options
                mSaveBasicProteinHashInfoFile = False
                mCheckForDuplicateProteinSequences = False

                ' Auto-enable creating a fixed fasta file
                mGenerateFixedFastaFile = True
                mFixedFastaOptions.ConsolidateProteinsWithDuplicateSeqs = False
                mFixedFastaOptions.RenameProteinsWithDuplicateNames = False

                ' Note: do not change .ConsolidateDupsIgnoreILDiff
                ' If .ConsolidateDupsIgnoreILDiff was enabled when the hash file was made with /B
                ' it should also be enabled when using /HashFile

                usingPreloadedProteinNames = True
            End If

            If mNormalizeFileLineEndCharacters Then
                mFastaFilePath = NormalizeFileLineEndings(
                    fastaFilePathToCheck,
                 "CRLF_" & Path.GetFileName(fastaFilePathToCheck),
                 eLineEndingCharacters.CRLF)

                If mFastaFilePath <> fastaFilePathToCheck Then
                    fastaFilePathToCheck = String.Copy(mFastaFilePath)
                    OnWroteLineEndNormalizedFASTA(fastaFilePathToCheck)
                End If
            Else
                mFastaFilePath = String.Copy(fastaFilePathToCheck)
            End If

            OnProgressUpdate("Parsing " & Path.GetFileName(mFastaFilePath), 0)

            Dim proteinHeaderFound = False
            Dim processingResidueBlock = False
            Dim blankLineProcessed = False

            Dim proteinName = String.Empty
            Dim sbCurrentResidues = New StringBuilder

            ' Initialize the RegEx objects

            Dim reProteinNameTruncation = New udtProteinNameTruncationRegex
            With reProteinNameTruncation
                ' Note that each of these RegEx tests contain two groups with captured text:

                ' The following will extract IPI:IPI00048500.11 from IPI:IPI00048500.11|ref|23848934
                .reMatchIPI =
                 New Regex("^(IPI:IPI[\w.]{2,})\|(.+)",
                  RegexOptions.Singleline Or RegexOptions.Compiled)

                ' The following will extract gi|169602219 from gi|169602219|ref|XP_001794531.1|
                .reMatchGI =
                 New Regex("^(gi\|\d+)\|(.+)",
                  RegexOptions.Singleline Or RegexOptions.Compiled)

                ' The following will extract jgi|Batde5|906240 from jgi|Batde5|90624|GP3.061830
                .reMatchJGI =
                 New Regex("^(jgi\|[^|]+\|[^|]+)\|(.+)",
                  RegexOptions.Singleline Or RegexOptions.Compiled)

                ' The following will extract bob|234384 from  bob|234384|ref|483293
                '                         or bob|845832 from  bob|845832;ref|384923
                .reMatchGeneric =
                 New Regex("^(\w{2,}[" &
                 CharArrayToString(mProteinNameFirstRefSepChars) & "][\w\d._]{2,})[" &
                 CharArrayToString(mProteinNameSubsequentRefSepChars) & "](.+)",
                 RegexOptions.Singleline Or RegexOptions.Compiled)
            End With

            With reProteinNameTruncation
                ' The following matches jgi|Batde5|23435 ; it requires that there be a number after the second bar
                .reMatchJGIBaseAndID =
                 New Regex("^jgi\|[^|]+\|\d+",
                   RegexOptions.Singleline Or RegexOptions.Compiled)

                ' Note that this RegEx contains a group with captured text:
                .reMatchDoubleBarOrColonAndBar =
                 New Regex("[" &
                  CharArrayToString(mProteinNameFirstRefSepChars) & "][^" &
                  CharArrayToString(mProteinNameSubsequentRefSepChars) & "]*([" &
                  CharArrayToString(mProteinNameSubsequentRefSepChars) & "])",
                  RegexOptions.Singleline Or RegexOptions.Compiled)
            End With

            ' Non-letter characters in residues
            Dim allowedResidueChars = "A-Z"
            If mAllowAsteriskInResidues Then allowedResidueChars &= "*"
            If mAllowDashInResidues Then allowedResidueChars &= "-"

            Dim reNonLetterResidues =
              New Regex("[^" & allowedResidueChars & "]",
              RegexOptions.Singleline Or RegexOptions.Compiled)

            ' Make sure mFixedFastaOptions.LongProteinNameSplitChars contains at least one character
            If mFixedFastaOptions.LongProteinNameSplitChars Is Nothing OrElse mFixedFastaOptions.LongProteinNameSplitChars.Length = 0 Then
                mFixedFastaOptions.LongProteinNameSplitChars = New Char() {DEFAULT_LONG_PROTEIN_NAME_SPLIT_CHAR}
            End If

            ' Initialize the rule details UDTs, which contain a RegEx object for each rule
            InitializeRuleDetails(mHeaderLineRules, headerLineRuleDetails)
            InitializeRuleDetails(mProteinNameRules, proteinNameRuleDetails)
            InitializeRuleDetails(mProteinDescriptionRules, proteinDescriptionRuleDetails)
            InitializeRuleDetails(mProteinSequenceRules, proteinSequenceRuleDetails)

            ' Open the file and read, at most, the first 100,000 characters to see if it contains CrLf or just Lf
            Dim terminatorSize = DetermineLineTerminatorSize(fastaFilePathToCheck)

            ' Pre-scan a portion of the fasta file to determine the appropriate value for mProteinNameSpannerCharLength
            AutoDetermineFastaProteinNameSpannerCharLength(mFastaFilePath, terminatorSize)

            ' Open the input file

            Using fastaReader = New StreamReader(New FileStream(fastaFilePathToCheck, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))

                ' Optionally, open the output fasta file
                If mGenerateFixedFastaFile Then

                    Try
                        fastaFilePathOut =
                         Path.Combine(Path.GetDirectoryName(fastaFilePathToCheck),
                         Path.GetFileNameWithoutExtension(fastaFilePathToCheck) & "_new.fasta")
                        fixedFastaWriter = New StreamWriter(New FileStream(fastaFilePathOut, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
                    Catch ex As Exception
                        ' Error opening output file
                        RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError,
                         "Error creating output file " & fastaFilePathOut & ": " & ex.Message, String.Empty)
                        OnErrorEvent("Error creating output file (Create _new.fasta)", ex)
                        Return False
                    End Try
                End If

                ' Optionally, open the Sequence Hash file
                If mSaveBasicProteinHashInfoFile Then
                    Dim basicProteinHashInfoFilePath = "<undefined>"

                    Try
                        basicProteinHashInfoFilePath =
                         Path.Combine(Path.GetDirectoryName(fastaFilePathToCheck),
                         Path.GetFileNameWithoutExtension(fastaFilePathToCheck) & PROTEIN_HASHES_FILENAME_SUFFIX)
                        sequenceHashWriter = New StreamWriter(New FileStream(basicProteinHashInfoFilePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))

                        Dim headerNames = New List(Of String) From {
                            "Protein_ID",
                            PROTEIN_NAME_COLUMN,
                            SEQUENCE_LENGTH_COLUMN,
                            SEQUENCE_HASH_COLUMN}

                        sequenceHashWriter.WriteLine(FlattenList(headerNames))

                    Catch ex As Exception
                        ' Error opening output file
                        RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError,
                         "Error creating output file " & basicProteinHashInfoFilePath & ": " & ex.Message, String.Empty)
                        OnErrorEvent("Error creating output file (Create " & PROTEIN_HASHES_FILENAME_SUFFIX & ")", ex)
                    End Try

                End If

                If mGenerateFixedFastaFile And (mFixedFastaOptions.RenameProteinsWithDuplicateNames OrElse mFixedFastaOptions.KeepDuplicateNamedProteinsUnlessMatchingSequence) Then
                    ' Make sure mCheckForDuplicateProteinNames is enabled
                    mCheckForDuplicateProteinNames = True
                End If

                ' Initialize proteinNames
                Dim proteinNames = New SortedSet(Of String)(StringComparer.CurrentCultureIgnoreCase)

                ' Optionally, initialize the protein sequence hash objects
                If mSaveProteinSequenceHashInfoFiles Then
                    mCheckForDuplicateProteinSequences = True
                End If

                If mGenerateFixedFastaFile And mFixedFastaOptions.ConsolidateProteinsWithDuplicateSeqs Then
                    mCheckForDuplicateProteinSequences = True
                    mSaveProteinSequenceHashInfoFiles = Not usingPreloadedProteinNames
                    consolidateDuplicateProteinSeqsInFasta = True
                    keepDuplicateNamedProteinsUnlessMatchingSequence = mFixedFastaOptions.KeepDuplicateNamedProteinsUnlessMatchingSequence
                    consolidateDupsIgnoreILDiff = mFixedFastaOptions.ConsolidateDupsIgnoreILDiff
                ElseIf mGenerateFixedFastaFile And mFixedFastaOptions.KeepDuplicateNamedProteinsUnlessMatchingSequence Then
                    mCheckForDuplicateProteinSequences = True
                    mSaveProteinSequenceHashInfoFiles = Not usingPreloadedProteinNames
                    consolidateDuplicateProteinSeqsInFasta = False
                    keepDuplicateNamedProteinsUnlessMatchingSequence = mFixedFastaOptions.KeepDuplicateNamedProteinsUnlessMatchingSequence
                    consolidateDupsIgnoreILDiff = mFixedFastaOptions.ConsolidateDupsIgnoreILDiff
                End If

                If mCheckForDuplicateProteinSequences Then
                    proteinSequenceHashes.Clear()
                    proteinSequenceHashCount = 0
                    ReDim proteinSeqHashInfo(99)
                End If

                ' Parse each line in the file
                Dim bytesRead As Int64 = 0
                Dim lastMemoryUsageReport = DateTime.UtcNow

                ' Note: This value is updated only if the line length is < mMaximumResiduesPerLine
                Dim currentValidResidueLineLengthMax = 0
                Dim processingDuplicateOrInvalidProtein = False

                Dim lastProgressReport = DateTime.UtcNow

                Do While Not fastaReader.EndOfStream

                    Dim lineIn As String

                    Try
                        lineIn = fastaReader.ReadLine()
                    Catch ex As OutOfMemoryException
                        OnErrorEvent(String.Format(
                            "Error in AnalyzeFastaFile reading line {0}; " &
                            "it is most likely millions of characters long, " &
                            "indicating a corrupt fasta file", mLineCount + 1), ex)
                        exceptionCaught = True
                        Exit Do
                    Catch ex As Exception
                        OnErrorEvent(String.Format("Error in AnalyzeFastaFile reading line {0}", mLineCount + 1), ex)
                        exceptionCaught = True
                        Exit Do
                    End Try

                    bytesRead += lineIn.Length + terminatorSize

                    If mLineCount Mod 250 = 0 Then
                        If MyBase.AbortProcessing Then Exit Do

                        If DateTime.UtcNow.Subtract(lastProgressReport).TotalSeconds >= 0.5 Then
                            lastProgressReport = DateTime.UtcNow
                            Dim percentComplete = CType(bytesRead / CType(fastaReader.BaseStream.Length, Single) * 100.0, Single)
                            If consolidateDuplicateProteinSeqsInFasta OrElse keepDuplicateNamedProteinsUnlessMatchingSequence Then
                                ' Bump the % complete down so that 100% complete in this routine will equate to 75% complete
                                ' The remaining 25% will occur in ConsolidateDuplicateProteinSeqsInFasta
                                percentComplete = percentComplete * 3 / 4
                            End If

                            MyBase.UpdateProgress("Validating FASTA File (" & Math.Round(percentComplete, 0) & "% Done)", percentComplete)

                            If DateTime.UtcNow.Subtract(lastMemoryUsageReport).TotalMinutes >= 1 Then
                                lastMemoryUsageReport = DateTime.UtcNow
                                ReportMemoryUsage(preloadedProteinNamesToKeep, proteinSequenceHashes, proteinNames, proteinSeqHashInfo)
                            End If
                        End If
                    End If

                    mLineCount += 1

                    If (lineIn.Length > 10000000) Then
                        RecordFastaFileError(mLineCount, 0, proteinName, eMessageCodeConstants.ResiduesLineTooLong, "Line is over 10 million residues long; skipping", String.Empty)
                        Continue Do
                    ElseIf (lineIn.Length > 1000000) Then
                        RecordFastaFileWarning(mLineCount, 0, proteinName, eMessageCodeConstants.ResiduesLineTooLong, "Line is over 1 million residues long; this is very suspicious", String.Empty)
                    ElseIf (lineIn.Length > 100000) Then
                        RecordFastaFileWarning(mLineCount, 0, proteinName, eMessageCodeConstants.ResiduesLineTooLong, "Line is over 1 million residues long; this could indicate a problem", String.Empty)
                    End If

                    If lineIn Is Nothing Then Continue Do

                    If lineIn.Trim.Length = 0 Then
                        ' We typically only want blank lines at the end of the fasta file or between two protein entries
                        blankLineProcessed = True
                        Continue Do
                    End If

                    If lineIn.Chars(0) = " "c Then
                        If mWarnLineStartsWithSpace Then
                            RecordFastaFileError(mLineCount, 0, String.Empty,
                                                 eMessageCodeConstants.LineStartsWithSpace, String.Empty, ExtractContext(lineIn, 0))
                        End If
                    End If

                    ' Note: Only trim the start of the line; do not trim the end of the line since Sequest incorrectly notates the peptide terminal state if a residue has a space after it
                    lineIn = lineIn.TrimStart

                    If lineIn.Chars(0) = mProteinLineStartChar Then
                        ' Protein entry

                        If sbCurrentResidues.Length > 0 Then
                            ProcessResiduesForPreviousProtein(
                                proteinName, sbCurrentResidues,
                                proteinSequenceHashes,
                                proteinSequenceHashCount, proteinSeqHashInfo,
                                consolidateDupsIgnoreILDiff,
                                fixedFastaWriter,
                                currentValidResidueLineLengthMax,
                                sequenceHashWriter)

                            currentValidResidueLineLengthMax = 0
                        End If

                        ' Now process this protein entry
                        mProteinCount += 1
                        proteinHeaderFound = True
                        processingResidueBlock = False
                        processingDuplicateOrInvalidProtein = False

                        proteinName = String.Empty

                        AnalyzeFastaProcessProteinHeader(
                            fixedFastaWriter,
                            lineIn,
                            proteinName,
                            processingDuplicateOrInvalidProtein,
                            preloadedProteinNamesToKeep,
                            proteinNames,
                            headerLineRuleDetails,
                            proteinNameRuleDetails,
                            proteinDescriptionRuleDetails,
                            reProteinNameTruncation)

                        If blankLineProcessed Then
                            ' The previous line was blank; raise a warning
                            If mWarnBlankLinesBetweenProteins Then
                                RecordFastaFileWarning(mLineCount, 0, proteinName, eMessageCodeConstants.BlankLineBeforeProteinName)
                            End If
                        End If

                    Else
                        ' Protein residues

                        If Not processingResidueBlock Then
                            If proteinHeaderFound Then
                                proteinHeaderFound = False

                                If blankLineProcessed Then
                                    RecordFastaFileError(mLineCount, 0, proteinName, eMessageCodeConstants.BlankLineBetweenProteinNameAndResidues)
                                End If
                            Else
                                RecordFastaFileError(mLineCount, 0, String.Empty, eMessageCodeConstants.ResiduesFoundWithoutProteinHeader)
                            End If

                            processingResidueBlock = True
                        Else
                            If blankLineProcessed Then
                                RecordFastaFileError(mLineCount, 0, proteinName, eMessageCodeConstants.BlankLineInMiddleOfResidues)
                            End If
                        End If

                        Dim newResidueCount = lineIn.Length
                        mResidueCount += newResidueCount

                        ' Check the line length; raise a warning if longer than suggested
                        If newResidueCount > mMaximumResiduesPerLine Then
                            RecordFastaFileWarning(mLineCount, 0, proteinName, eMessageCodeConstants.ResiduesLineTooLong, newResidueCount.ToString, String.Empty)
                        End If

                        ' Test the protein sequence rules
                        EvaluateRules(proteinSequenceRuleDetails, proteinName, lineIn, 0, lineIn, 5)

                        If mGenerateFixedFastaFile OrElse mCheckForDuplicateProteinSequences OrElse mSaveBasicProteinHashInfoFile Then
                            Dim residuesClean As String

                            If mFixedFastaOptions.RemoveInvalidResidues Then
                                ' Auto-fix residues to remove any non-letter characters (spaces, asterisks, etc.)
                                residuesClean = reNonLetterResidues.Replace(lineIn, String.Empty)
                            Else
                                ' Do not remove non-letter characters, but do remove leading or trailing whitespace
                                residuesClean = String.Copy(lineIn.Trim())
                            End If

                            If Not fixedFastaWriter Is Nothing AndAlso Not processingDuplicateOrInvalidProtein Then
                                If residuesClean <> lineIn Then
                                    mFixedFastaStats.UpdatedResidueLines += 1
                                End If

                                If Not mFixedFastaOptions.WrapLongResidueLines Then
                                    ' Only write out this line if not auto-wrapping long residue lines
                                    ' If we are auto-wrapping, then the residues will be written out by the call to ProcessResiduesForPreviousProtein
                                    fixedFastaWriter.WriteLine(residuesClean)
                                End If
                            End If

                            If mCheckForDuplicateProteinSequences OrElse mFixedFastaOptions.WrapLongResidueLines Then
                                ' Only add the residues if this is not a duplicate/invalid protein
                                If Not processingDuplicateOrInvalidProtein Then
                                    sbCurrentResidues.Append(residuesClean)
                                    If residuesClean.Length > currentValidResidueLineLengthMax Then
                                        currentValidResidueLineLengthMax = residuesClean.Length
                                    End If
                                End If
                            End If
                        End If

                        ' Reset the blank line tracking variable
                        blankLineProcessed = False

                    End If

                Loop

                If sbCurrentResidues.Length > 0 Then
                    ProcessResiduesForPreviousProtein(
                       proteinName, sbCurrentResidues,
                       proteinSequenceHashes,
                       proteinSequenceHashCount, proteinSeqHashInfo,
                       consolidateDupsIgnoreILDiff,
                       fixedFastaWriter, currentValidResidueLineLengthMax,
                       sequenceHashWriter)
                End If

                If mCheckForDuplicateProteinSequences Then
                    ' Step through proteinSeqHashInfo and look for duplicate sequences
                    For index = 0 To proteinSequenceHashCount - 1
                        If proteinSeqHashInfo(index).AdditionalProteins.Count > 0 Then
                            With proteinSeqHashInfo(index)
                                RecordFastaFileWarning(mLineCount, 0, .ProteinNameFirst, eMessageCodeConstants.DuplicateProteinSequence,
                                  .ProteinNameFirst & ", " & FlattenArray(.AdditionalProteins, ","c), .SequenceStart)
                            End With
                        End If
                    Next index
                End If

                Dim memoryUsageMB = clsMemoryUsageLogger.GetProcessMemoryUsageMB
                If memoryUsageMB > mProcessMemoryUsageMBAtStart * 4 OrElse
                   memoryUsageMB - mProcessMemoryUsageMBAtStart > 50 Then
                    ReportMemoryUsage(preloadedProteinNamesToKeep, proteinSequenceHashes, proteinNames, proteinSeqHashInfo)
                End If

            End Using

            ' Close the output files
            If Not fixedFastaWriter Is Nothing Then
                fixedFastaWriter.Close()
            End If

            If Not sequenceHashWriter Is Nothing Then
                sequenceHashWriter.Close()
            End If

            If mProteinCount = 0 Then
                RecordFastaFileError(mLineCount, 0, String.Empty, eMessageCodeConstants.ProteinEntriesNotFound)
            ElseIf proteinHeaderFound Then
                RecordFastaFileError(mLineCount, 0, proteinName, eMessageCodeConstants.FinalProteinEntryMissingResidues)
            ElseIf Not blankLineProcessed Then
                ' File does not end in multiple blank lines; need to re-open it using a binary reader and check the last two characters to make sure they're valid
                Threading.Thread.Sleep(100)

                If Not VerifyLinefeedAtEOF(fastaFilePathToCheck, mAddMissingLinefeedAtEOF) Then
                    RecordFastaFileError(mLineCount, 0, String.Empty, eMessageCodeConstants.FileDoesNotEndWithLinefeed)
                End If
            End If

            If usingPreloadedProteinNames Then
                ' Report stats on the number of proteins read, the number written, and any that had duplicate protein names in the original fasta file
                Dim nameCountNotFound = 0
                Dim duplicateProteinNameCount = 0
                Dim proteinCountWritten = 0
                Dim preloadedProteinNameCount = 0

                For Each spanningKey In preloadedProteinNamesToKeep.GetSpanningKeys
                    Dim proteinsForKey = preloadedProteinNamesToKeep.GetListForSpanningKey(spanningKey)
                    preloadedProteinNameCount += proteinsForKey.Count

                    For Each proteinEntry In proteinsForKey
                        If proteinEntry.Value = 0 Then
                            nameCountNotFound += 1
                        Else
                            proteinCountWritten += 1
                            If proteinEntry.Value > 1 Then
                                duplicateProteinNameCount += 1
                            End If
                        End If
                    Next
                Next

                Console.WriteLine()
                If proteinCountWritten = preloadedProteinNameCount Then
                    ShowMessage("Fixed Fasta has all " & proteinCountWritten.ToString("#,##0") & " proteins determined from the pre-existing protein hash file")
                Else
                    ShowMessage("Fixed Fasta has " & proteinCountWritten.ToString("#,##0") & " of the " & preloadedProteinNameCount.ToString("#,##0") & " proteins determined from the pre-existing protein hash file")
                End If

                If nameCountNotFound > 0 Then
                    ShowMessage("WARNING: " & nameCountNotFound.ToString("#,##0") & " protein names were in the protein name list to keep, but were not found in the fasta file")
                End If

                If duplicateProteinNameCount > 0 Then
                    ShowMessage("WARNING: " & duplicateProteinNameCount.ToString("#,##0") & " protein names were present multiple times in the fasta file; duplicate entries were skipped")
                End If

                success = Not exceptionCaught

            ElseIf mSaveProteinSequenceHashInfoFiles Then
                Dim percentComplete = 98.0!
                If consolidateDuplicateProteinSeqsInFasta OrElse keepDuplicateNamedProteinsUnlessMatchingSequence Then
                    percentComplete = percentComplete * 3 / 4
                End If
                MyBase.UpdateProgress("Validating FASTA File (" & Math.Round(percentComplete, 0) & "% Done)", percentComplete)

                Dim hashInfoSuccess = AnalyzeFastaSaveHashInfo(
                    fastaFilePathToCheck,
                    proteinSequenceHashCount,
                    proteinSeqHashInfo,
                    consolidateDuplicateProteinSeqsInFasta,
                    consolidateDupsIgnoreILDiff,
                    keepDuplicateNamedProteinsUnlessMatchingSequence,
                    fastaFilePathOut)

                success = hashInfoSuccess And Not exceptionCaught
            Else
                success = Not exceptionCaught
            End If

            If MyBase.AbortProcessing Then
                MyBase.UpdateProgress("Parsing aborted")
            Else
                MyBase.UpdateProgress("Parsing complete", 100)
            End If

        Catch ex As Exception
            OnErrorEvent(String.Format("Error in AnalyzeFastaFile reading line {0}", mLineCount), ex)
            success = False
        Finally
            ' These close statements will typically be redundant,
            ' However, if an exception occurs, then they will be needed to close the files

            If Not fixedFastaWriter Is Nothing Then
                fixedFastaWriter.Close()
            End If

            If Not sequenceHashWriter Is Nothing Then
                sequenceHashWriter.Close()
            End If

        End Try

        Return success

    End Function

    Private Sub AnalyzeFastaProcessProteinHeader(
      fixedFastaWriter As TextWriter,
      lineIn As String,
      <Out> ByRef proteinName As String,
      <Out> ByRef processingDuplicateOrInvalidProtein As Boolean,
      preloadedProteinNamesToKeep As clsNestedStringIntList,
      proteinNames As ISet(Of String),
      headerLineRuleDetails As IList(Of udtRuleDefinitionExtendedType),
      proteinNameRuleDetails As IList(Of udtRuleDefinitionExtendedType),
      proteinDescriptionRuleDetails As IList(Of udtRuleDefinitionExtendedType),
      reProteinNameTruncation As udtProteinNameTruncationRegex)

        Dim descriptionStartIndex As Integer

        Dim proteinDescription As String = String.Empty

        Dim skipDuplicateProtein = False

        proteinName = String.Empty
        processingDuplicateOrInvalidProtein = True

        Try
            SplitFastaProteinHeaderLine(lineIn, proteinName, proteinDescription, descriptionStartIndex)

            If proteinName.Length = 0 Then
                processingDuplicateOrInvalidProtein = True
            Else
                processingDuplicateOrInvalidProtein = False
            End If

            ' Test the header line rules
            EvaluateRules(headerLineRuleDetails, proteinName, lineIn, 0, lineIn, DEFAULT_CONTEXT_LENGTH)

            If proteinDescription.Length > 0 Then
                ' Test the protein description rules

                EvaluateRules(
                    proteinDescriptionRuleDetails, proteinName, proteinDescription,
                    descriptionStartIndex, lineIn, DEFAULT_CONTEXT_LENGTH)
            End If

            If proteinName.Length > 0 Then

                ' Check for protein names that are too long or too short
                If proteinName.Length < mMinimumProteinNameLength Then
                    RecordFastaFileWarning(mLineCount, 1, proteinName,
                     eMessageCodeConstants.ProteinNameIsTooShort, proteinName.Length.ToString, String.Empty)
                ElseIf proteinName.Length > mMaximumProteinNameLength Then
                    RecordFastaFileError(mLineCount, 1, proteinName,
                     eMessageCodeConstants.ProteinNameIsTooLong, proteinName.Length.ToString, String.Empty)
                End If

                ' Test the protein name rules
                EvaluateRules(proteinNameRuleDetails, proteinName, proteinName, 1, lineIn, DEFAULT_CONTEXT_LENGTH)

                If Not preloadedProteinNamesToKeep Is Nothing AndAlso preloadedProteinNamesToKeep.Count > 0 Then
                    ' See if preloadedProteinNamesToKeep contains proteinName
                    Dim matchCount As Integer = preloadedProteinNamesToKeep.GetValueForItem(proteinName, -1)

                    If matchCount >= 0 Then
                        ' Name is known; increment the value for this protein

                        If matchCount = 0 Then
                            skipDuplicateProtein = False
                        Else
                            ' An entry with this protein name has already been written
                            ' Do not include the duplicate
                            skipDuplicateProtein = True
                        End If

                        If Not preloadedProteinNamesToKeep.SetValueForItem(proteinName, matchCount + 1) Then
                            ShowMessage("WARNING: protein " & proteinName & " not found in preloadedProteinNamesToKeep")
                        End If

                    Else
                        ' Unknown protein name; do not keep this protein
                        skipDuplicateProtein = True
                        processingDuplicateOrInvalidProtein = True
                    End If

                    If mGenerateFixedFastaFile Then
                        ' Make sure proteinDescription doesn't start with a | or space
                        If proteinDescription.Length > 0 Then
                            proteinDescription = proteinDescription.TrimStart(New Char() {"|"c, " "c})
                        End If
                    End If

                Else

                    If mGenerateFixedFastaFile Then
                        proteinName = AutoFixProteinNameAndDescription(proteinName, proteinDescription, reProteinNameTruncation)
                    End If

                    ' Optionally, check for duplicate protein names
                    If mCheckForDuplicateProteinNames Then
                        proteinName = ExamineProteinName(proteinName, proteinNames, skipDuplicateProtein, processingDuplicateOrInvalidProtein)

                        If skipDuplicateProtein Then
                            processingDuplicateOrInvalidProtein = True
                        End If
                    End If

                End If

                If Not fixedFastaWriter Is Nothing AndAlso Not skipDuplicateProtein Then
                    fixedFastaWriter.WriteLine(ConstructFastaHeaderLine(proteinName.Trim, proteinDescription.Trim))
                End If
            End If


        Catch ex As Exception
            RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError,
             "Error parsing protein header line '" & lineIn & "': " & ex.Message, String.Empty)
            OnErrorEvent("Error parsing protein header line", ex)
        End Try


    End Sub

    Private Function AnalyzeFastaSaveHashInfo(
      fastaFilePathToCheck As String,
      proteinSequenceHashCount As Integer,
      proteinSeqHashInfo As IList(Of clsProteinHashInfo),
      consolidateDuplicateProteinSeqsInFasta As Boolean,
      consolidateDupsIgnoreILDiff As Boolean,
      keepDuplicateNamedProteinsUnlessMatchingSequence As Boolean,
      fastaFilePathOut As String) As Boolean

        Dim swUniqueProteinSeqsOut As StreamWriter
        Dim swDuplicateProteinMapping As StreamWriter = Nothing

        Dim uniqueProteinSeqsFileOut As String = String.Empty
        Dim duplicateProteinMappingFileOut As String = String.Empty

        Dim index As Integer
        Dim duplicateIndex As Integer

        Dim duplicateProteinSeqsFound As Boolean
        Dim success As Boolean

        duplicateProteinSeqsFound = False

        Try
            uniqueProteinSeqsFileOut =
             Path.Combine(Path.GetDirectoryName(fastaFilePathToCheck),
             Path.GetFileNameWithoutExtension(fastaFilePathToCheck) & "_UniqueProteinSeqs.txt")

            ' Create swUniqueProteinSeqsOut
            swUniqueProteinSeqsOut = New StreamWriter(New FileStream(uniqueProteinSeqsFileOut, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
        Catch ex As Exception
            ' Error opening output file
            RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError,
             "Error creating output file " & uniqueProteinSeqsFileOut & ": " & ex.Message, String.Empty)
            OnErrorEvent("Error creating output file (SaveHashInfo to _UniqueProteinSeqs.txt)", ex)
            Return False
        End Try

        Try
            ' Define the path to the protein mapping file, but don't create it yet; just delete it if it exists
            ' We'll only create it if two or more proteins have the same protein sequence
            duplicateProteinMappingFileOut =
              Path.Combine(Path.GetDirectoryName(fastaFilePathToCheck),
              Path.GetFileNameWithoutExtension(fastaFilePathToCheck) & "_UniqueProteinSeqDuplicates.txt")                       ' Look for duplicateProteinMappingFileOut and erase it if it exists

            If File.Exists(duplicateProteinMappingFileOut) Then
                File.Delete(duplicateProteinMappingFileOut)
            End If
        Catch ex As Exception
            ' Error deleting output file
            RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError,
             "Error deleting output file " & duplicateProteinMappingFileOut & ": " & ex.Message, String.Empty)
            OnErrorEvent("Error deleting output file (SaveHashInfo to _UniqueProteinSeqDuplicates.txt)", ex)
            Return False
        End Try

        Try

            Dim headerColumns = New List(Of String) From {
                "Sequence_Index",
                "Protein_Name_First",
                SEQUENCE_LENGTH_COLUMN,
                SEQUENCE_HASH_COLUMN,
                "Protein_Count",
                "Duplicate_Proteins"}

            swUniqueProteinSeqsOut.WriteLine(FlattenList(headerColumns))

            For index = 0 To proteinSequenceHashCount - 1
                With proteinSeqHashInfo(index)

                    Dim dataValues = New List(Of String) From {
                        (index + 1).ToString,
                        .ProteinNameFirst,
                        .SequenceLength.ToString(),
                        .SequenceHash,
                        (.AdditionalProteins.Count + 1).ToString(),
                        FlattenArray(.AdditionalProteins, ","c)}

                    swUniqueProteinSeqsOut.WriteLine(FlattenList(dataValues))

                    If .AdditionalProteins.Count > 0 Then
                        duplicateProteinSeqsFound = True

                        If swDuplicateProteinMapping Is Nothing Then
                            ' Need to create swDuplicateProteinMapping
                            swDuplicateProteinMapping = New StreamWriter(New FileStream(duplicateProteinMappingFileOut, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))

                            Dim proteinHeaderColumns = New List(Of String) From {
                                "Sequence_Index",
                                "Protein_Name_First",
                                SEQUENCE_LENGTH_COLUMN,
                                "Duplicate_Protein"}

                            swDuplicateProteinMapping.WriteLine(FlattenList(proteinHeaderColumns))
                        End If

                        For Each additionalProtein As String In .AdditionalProteins
                            If Not .AdditionalProteins(duplicateIndex) Is Nothing Then
                                If additionalProtein.Trim.Length > 0 Then
                                    Dim proteinDataValues = New List(Of String) From {
                                        (index + 1).ToString,
                                        .ProteinNameFirst,
                                        .SequenceLength.ToString(),
                                        additionalProtein}

                                    swDuplicateProteinMapping.WriteLine(FlattenList(proteinDataValues))
                                End If
                            End If
                        Next

                    ElseIf .DuplicateProteinNameCount > 0 Then
                        duplicateProteinSeqsFound = True
                    End If
                End With

            Next index

            swUniqueProteinSeqsOut.Close()
            If Not swDuplicateProteinMapping Is Nothing Then swDuplicateProteinMapping.Close()

            success = True

        Catch ex As Exception
            ' Error writing results
            RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError,
             "Error writing results to " & uniqueProteinSeqsFileOut & " or " & duplicateProteinMappingFileOut & ": " & ex.Message, String.Empty)
            OnErrorEvent("Error writing results to " & uniqueProteinSeqsFileOut & " or " & duplicateProteinMappingFileOut, ex)
            success = False
        End Try

        If success And proteinSequenceHashCount > 0 And duplicateProteinSeqsFound Then
            If consolidateDuplicateProteinSeqsInFasta OrElse keepDuplicateNamedProteinsUnlessMatchingSequence Then
                success = CorrectForDuplicateProteinSeqsInFasta(consolidateDuplicateProteinSeqsInFasta, consolidateDupsIgnoreILDiff, fastaFilePathOut, proteinSequenceHashCount, proteinSeqHashInfo)
            End If
        End If

        Return success

    End Function

    ''' <summary>
    ''' Pre-scan a portion of the Fasta file to determine the appropriate value for mProteinNameSpannerCharLength
    ''' </summary>
    ''' <param name="fastaFilePathToTest">Fasta file to examine</param>
    ''' <param name="terminatorSize">Linefeed length (1 for LF or 2 for CRLF)</param>
    ''' <remarks>
    ''' Reads 50 MB chunks from 10 sections of the Fasta file (or the entire Fasta file if under 500 MB in size)
    ''' Keeps track of the portion of protein names in common between adjacent proteins
    ''' Uses this information to determine an appropriate value for mProteinNameSpannerCharLength
    ''' </remarks>
    Private Sub AutoDetermineFastaProteinNameSpannerCharLength(fastaFilePathToTest As String, terminatorSize As Integer)

        Const PARTS_TO_SAMPLE = 10
        Const KILOBYTES_PER_SAMPLE = 51200

        Dim proteinStartLetters = New Dictionary(Of String, Integer)
        Dim startTime = DateTime.UtcNow
        Dim showStats = False

        Dim fastaFile = New FileInfo(fastaFilePathToTest)
        If Not fastaFile.Exists Then Return

        Dim fullScanLengthBytes = 1024L * PARTS_TO_SAMPLE * KILOBYTES_PER_SAMPLE
        Dim linesReadTotal As Int64

        If fastaFile.Length < fullScanLengthBytes Then
            fullScanLengthBytes = fastaFile.Length
            linesReadTotal = AutoDetermineFastaProteinNameSpannerCharLength(fastaFile, terminatorSize, proteinStartLetters, 0, fastaFile.Length)
        Else

            Dim stepSizeBytes = CLng(Math.Round(fastaFile.Length / PARTS_TO_SAMPLE, 0))

            For byteOffsetStart As Int64 = 0 To fastaFile.Length Step stepSizeBytes
                Dim linesRead = AutoDetermineFastaProteinNameSpannerCharLength(fastaFile, terminatorSize, proteinStartLetters, byteOffsetStart, KILOBYTES_PER_SAMPLE * 1024)

                If linesRead < 0 Then
                    ' This indicates an error, probably from a corrupt file; do not read further
                    Exit For
                End If

                linesReadTotal += linesRead

                If Not showStats AndAlso DateTime.UtcNow.Subtract(startTime).TotalMilliseconds > 500 Then
                    showStats = True
                    ShowMessage("Pre-scanning the file to look for common base protein names")
                End If
            Next
        End If

        If proteinStartLetters.Count = 0 Then
            mProteinNameSpannerCharLength = 1
        Else
            Dim preScanProteinCount = (From item In proteinStartLetters Select item.Value).Sum()

            If showStats Then
                Dim percentFileProcessed = fullScanLengthBytes / fastaFile.Length * 100
                ShowMessage(String.Format(
                    "  parsed {0:0}% of the file, reading {1:#,##0} lines and finding {2:#,##0} proteins",
                    percentFileProcessed, linesReadTotal, preScanProteinCount))
            End If

            ' Determine the appropriate spanner length given the observation counts of the base names
            mProteinNameSpannerCharLength = clsNestedStringIntList.DetermineSpannerLengthUsingStartLetterStats(proteinStartLetters)
        End If

        ShowMessage("Using ProteinNameSpannerCharLength = " & mProteinNameSpannerCharLength)
        Console.WriteLine()

    End Sub

    ''' <summary>
    ''' Read a portion of the Fasta file, comparing adjacent protein names and keeping track of the name portions in common
    ''' </summary>
    ''' <param name="fastaFile"></param>
    ''' <param name="terminatorSize"></param>
    ''' <param name="proteinStartLetters"></param>
    ''' <param name="startOffset"></param>
    ''' <param name="bytesToRead"></param>
    ''' <returns>The number of lines read</returns>
    ''' <remarks></remarks>
    Private Function AutoDetermineFastaProteinNameSpannerCharLength(
      fastaFile As FileInfo,
      terminatorSize As Integer,
      proteinStartLetters As IDictionary(Of String, Integer),
      startOffset As Int64,
      bytesToRead As Int64) As Int64

        Dim linesRead = 0L

        Try
            Dim previousProteinLength = 0
            Dim previousProteinName = String.Empty

            If startOffset >= fastaFile.Length Then
                ShowMessage("Ignoring byte offset of " & startOffset &
                            " in AutoDetermineProteinNameSpannerCharLength since past the end of the file " &
                            "(" & fastaFile.Length & " bytes")
                Return 0
            End If

            Dim bytesRead As Int64 = 0

            Dim firstLineDiscarded As Boolean
            If startOffset = 0 Then
                firstLineDiscarded = True
            End If

            Using inStream = New FileStream(fastaFile.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                inStream.Position = startOffset

                Using reader = New StreamReader(inStream)

                    Do While Not reader.EndOfStream

                        Dim lineIn = reader.ReadLine()
                        bytesRead += terminatorSize
                        linesRead += 1

                        If String.IsNullOrEmpty(lineIn) Then
                            Continue Do
                        End If

                        bytesRead += lineIn.Length
                        If Not firstLineDiscarded Then
                            ' We can't trust that this was a full line of text; skip it
                            firstLineDiscarded = True
                            Continue Do
                        End If

                        If bytesRead > bytesToRead Then
                            Exit Do
                        End If

                        If Not lineIn.Chars(0) = mProteinLineStartChar Then
                            Continue Do
                        End If

                        ' Make sure the protein name and description are valid
                        ' Find the first space and/or tab
                        Dim spaceIndex = GetBestSpaceIndex(lineIn)
                        Dim proteinName As String

                        If spaceIndex > 1 Then
                            proteinName = lineIn.Substring(1, spaceIndex - 1)
                        Else
                            ' Line does not contain a description
                            If spaceIndex <= 0 Then
                                If lineIn.Trim.Length <= 1 Then
                                    Continue Do
                                Else
                                    ' The line contains a protein name, but not a description
                                    proteinName = lineIn.Substring(1)
                                End If
                            Else
                                ' Space or tab found directly after the > symbol
                                Continue Do
                            End If
                        End If

                        If previousProteinLength = 0 Then
                            previousProteinName = String.Copy(proteinName)
                            previousProteinLength = previousProteinName.Length
                            Continue Do
                        End If

                        Dim currentNameLength = proteinName.Length
                        Dim charIndex = 0

                        While charIndex < previousProteinLength
                            If charIndex >= currentNameLength Then
                                Exit While
                            End If

                            If previousProteinName(charIndex) <> proteinName.Chars(charIndex) Then
                                ' Difference found; add/update the dictionary
                                Exit While
                            End If

                            charIndex += 1
                        End While

                        Dim charsInCommon = charIndex
                        If charsInCommon > 0 Then
                            Dim baseName As String = previousProteinName.Substring(0, charsInCommon)
                            Dim matchCount = 0

                            If proteinStartLetters.TryGetValue(baseName, matchCount) Then
                                proteinStartLetters(baseName) = matchCount + 1
                            Else
                                proteinStartLetters.Add(baseName, 1)
                            End If
                        End If

                        previousProteinName = String.Copy(proteinName)
                        previousProteinLength = previousProteinName.Length
                    Loop

                End Using

            End Using

        Catch ex As OutOfMemoryException
            OnErrorEvent("Out of memory exception in AutoDetermineProteinNameSpannerCharLength", ex)

            ' Example message: Insufficient memory to continue the execution of the program
            ' This can happen with a corrupt .fasta file with a line that has millions of characters
            Return -1

        Catch ex As Exception
            OnErrorEvent("Error in AutoDetermineProteinNameSpannerCharLength", ex)
        End Try

        Return linesRead

    End Function

    Private Function AutoFixProteinNameAndDescription(
      ByRef proteinName As String,
      ByRef proteinDescription As String,
      reProteinNameTruncation As udtProteinNameTruncationRegex) As String

        Dim proteinNameTooLong As Boolean
        Dim reMatch As Match
        Dim newProteinName As String
        Dim charIndex As Integer
        Dim minCharIndex As Integer
        Dim extraProteinNameText As String
        Dim invalidChar As Char

        Dim multipleRefsSplitOutFromKnownAccession = False

        ' Auto-fix potential errors in the protein name

        ' Possibly truncate to mMaximumProteinNameLength characters
        If proteinName.Length > mMaximumProteinNameLength Then
            proteinNameTooLong = True
        Else
            proteinNameTooLong = False
        End If

        If mFixedFastaOptions.SplitOutMultipleRefsForKnownAccession OrElse
           (mFixedFastaOptions.TruncateLongProteinNames And proteinNameTooLong) Then

            ' First see if the name fits the pattern IPI:IPI00048500.11|
            ' Next see if the name fits the pattern gi|7110699|
            ' Next see if the name fits the pattern jgi
            ' Next see if the name fits the generic pattern defined by reProteinNameTruncation.reMatchGeneric
            ' Otherwise, use mFixedFastaOptions.LongProteinNameSplitChars to define where to truncate

            newProteinName = String.Copy(proteinName)
            extraProteinNameText = String.Empty

            reMatch = reProteinNameTruncation.reMatchIPI.Match(proteinName)
            If reMatch.Success Then
                multipleRefsSplitOutFromKnownAccession = True
            Else
                ' IPI didn't match; try gi
                reMatch = reProteinNameTruncation.reMatchGI.Match(proteinName)
            End If

            If reMatch.Success Then
                multipleRefsSplitOutFromKnownAccession = True
            Else
                ' GI didn't match; try jgi
                reMatch = reProteinNameTruncation.reMatchJGI.Match(proteinName)
            End If

            If reMatch.Success Then
                multipleRefsSplitOutFromKnownAccession = True
            Else
                ' jgi didn't match; try generic (text separated by a series of colons or bars),
                '  but only if the name is too long
                If mFixedFastaOptions.TruncateLongProteinNames And proteinNameTooLong Then
                    reMatch = reProteinNameTruncation.reMatchGeneric.Match(proteinName)
                End If
            End If

            If reMatch.Success Then
                ' Truncate the protein name, but move the truncated portion into the next group
                newProteinName = reMatch.Groups(1).Value
                extraProteinNameText = reMatch.Groups(2).Value

            ElseIf mFixedFastaOptions.TruncateLongProteinNames And proteinNameTooLong Then

                ' Name is too long, but it didn't match the known patterns
                ' Find the last occurrence of mFixedFastaOptions.LongProteinNameSplitChars (default is vertical bar)
                '   and truncate the text following the match
                ' Repeat the process until the protein name length >= mMaximumProteinNameLength

                ' See if any of the characters in proteinNameSplitChars is present after
                ' character 6 but less than character mMaximumProteinNameLength
                minCharIndex = 6

                Do
                    charIndex = newProteinName.LastIndexOfAny(mFixedFastaOptions.LongProteinNameSplitChars)
                    If charIndex >= minCharIndex Then
                        If extraProteinNameText.Length > 0 Then
                            extraProteinNameText = "|" & extraProteinNameText
                        End If
                        extraProteinNameText = newProteinName.Substring(charIndex + 1) & extraProteinNameText
                        newProteinName = newProteinName.Substring(0, charIndex)
                    Else
                        charIndex = -1
                    End If

                Loop While charIndex > 0 And newProteinName.Length > mMaximumProteinNameLength

            End If

            If extraProteinNameText.Length > 0 Then
                If proteinNameTooLong Then
                    mFixedFastaStats.TruncatedProteinNameCount += 1
                Else
                    mFixedFastaStats.ProteinNamesMultipleRefsRemoved += 1
                End If

                proteinName = String.Copy(newProteinName)

                PrependExtraTextToProteinDescription(extraProteinNameText, proteinDescription)
            End If

        End If

        If mFixedFastaOptions.ProteinNameInvalidCharsToRemove.Length > 0 Then
            newProteinName = String.Copy(proteinName)

            ' First remove invalid characters from the beginning or end of the protein name
            newProteinName = newProteinName.Trim(mFixedFastaOptions.ProteinNameInvalidCharsToRemove)

            If newProteinName.Length >= 1 Then
                For Each invalidChar In mFixedFastaOptions.ProteinNameInvalidCharsToRemove
                    ' Next, replace any remaining instances of the character with an underscore
                    newProteinName = newProteinName.Replace(invalidChar, INVALID_PROTEIN_NAME_CHAR_REPLACEMENT)
                Next

                If proteinName <> newProteinName Then
                    If newProteinName.Length >= 3 Then
                        proteinName = String.Copy(newProteinName)
                        mFixedFastaStats.ProteinNamesInvalidCharsReplaced += 1
                    End If
                End If
            End If
        End If

        If mFixedFastaOptions.SplitOutMultipleRefsInProteinName AndAlso Not multipleRefsSplitOutFromKnownAccession Then
            ' Look for multiple refs in the protein name, but only if we didn't already split out multiple refs above

            reMatch = reProteinNameTruncation.reMatchDoubleBarOrColonAndBar.Match(proteinName)
            If reMatch.Success Then
                ' Protein name contains 2 or more vertical bars, or a colon and a bar
                ' Split out the multiple refs and place them in the description
                ' However, jgi names are supposed to have two vertical bars, so we need to treat that data differently

                extraProteinNameText = String.Empty

                reMatch = reProteinNameTruncation.reMatchJGIBaseAndID.Match(proteinName)
                If reMatch.Success Then
                    ' ProteinName is similar to jgi|Organism|00000
                    ' Check whether there is any text following the match
                    If reMatch.Length < proteinName.Length Then
                        ' Extra text exists; populate extraProteinNameText
                        extraProteinNameText = proteinName.Substring(reMatch.Length + 1)
                        proteinName = reMatch.ToString
                    End If
                Else
                    ' Find the first vertical bar or colon
                    charIndex = proteinName.IndexOfAny(mProteinNameFirstRefSepChars)

                    If charIndex > 0 Then
                        ' Find the second vertical bar, colon, or semicolon
                        charIndex = proteinName.IndexOfAny(mProteinNameSubsequentRefSepChars, charIndex + 1)

                        If charIndex > 0 Then
                            ' Split the protein name
                            extraProteinNameText = proteinName.Substring(charIndex + 1)
                            proteinName = proteinName.Substring(0, charIndex)
                        End If
                    End If

                End If

                If extraProteinNameText.Length > 0 Then
                    PrependExtraTextToProteinDescription(extraProteinNameText, proteinDescription)
                    mFixedFastaStats.ProteinNamesMultipleRefsRemoved += 1
                End If

            End If
        End If

        ' Make sure proteinDescription doesn't start with a | or space
        If proteinDescription.Length > 0 Then
            proteinDescription = proteinDescription.TrimStart(New Char() {"|"c, " "c})
        End If

        Return proteinName

    End Function

    Private Function BoolToStringInt(value As Boolean) As String
        If value Then
            Return "1"
        Else
            Return "0"
        End If
    End Function

    Private Function CharArrayToString(charArray As IEnumerable(Of Char)) As String
        Return String.Join("", charArray)
    End Function


    Private Sub ClearAllRules()
        Me.ClearRules(RuleTypes.HeaderLine)
        Me.ClearRules(RuleTypes.ProteinDescription)
        Me.ClearRules(RuleTypes.ProteinName)
        Me.ClearRules(RuleTypes.ProteinSequence)

        mMasterCustomRuleID = CUSTOM_RULE_ID_START
    End Sub

    Private Sub ClearRules(ruleType As RuleTypes)
        Select Case ruleType
            Case RuleTypes.HeaderLine
                Me.ClearRulesDataStructure(mHeaderLineRules)
            Case RuleTypes.ProteinDescription
                Me.ClearRulesDataStructure(mProteinDescriptionRules)
            Case RuleTypes.ProteinName
                Me.ClearRulesDataStructure(mProteinNameRules)
            Case RuleTypes.ProteinSequence
                Me.ClearRulesDataStructure(mProteinSequenceRules)
        End Select
    End Sub

    Private Sub ClearRulesDataStructure(ByRef rules() As udtRuleDefinitionType)
        ReDim rules(-1)
    End Sub

    Public Function ComputeProteinHash(sbResidues As StringBuilder, consolidateDupsIgnoreILDiff As Boolean) As String

        If sbResidues.Length > 0 Then
            ' Compute the hash value for sbCurrentResidues
            If consolidateDupsIgnoreILDiff Then
                Return HashUtilities.ComputeStringHashSha1(sbResidues.ToString().Replace("L"c, "I"c)).ToUpper()
            Else
                Return HashUtilities.ComputeStringHashSha1(sbResidues.ToString()).ToUpper()
            End If
        Else
            Return String.Empty
        End If

    End Function

    Private Function ComputeTotalSpecifiedCount(errorStats As udtItemSummaryIndexedType) As Integer
        Dim total As Integer
        Dim index As Integer

        total = 0
        For index = 0 To errorStats.ErrorStatsCount - 1
            total += errorStats.ErrorStats(index).CountSpecified
        Next index

        Return total

    End Function

    Private Function ComputeTotalUnspecifiedCount(errorStats As udtItemSummaryIndexedType) As Integer
        Dim total As Integer
        Dim index As Integer

        total = 0
        For index = 0 To errorStats.ErrorStatsCount - 1
            total += errorStats.ErrorStats(index).CountUnspecified
        Next index

        Return total

    End Function

    ''' <summary>
    ''' Looks for duplicate proteins in the Fasta file
    ''' Creates a new fasta file that has exact duplicates removed
    ''' Will consolidate proteins with the same sequence if consolidateDuplicateProteinSeqsInFasta=True
    ''' </summary>
    ''' <param name="consolidateDuplicateProteinSeqsInFasta"></param>
    ''' <param name="fixedFastaFilePath"></param>
    ''' <param name="proteinSequenceHashCount"></param>
    ''' <param name="proteinSeqHashInfo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CorrectForDuplicateProteinSeqsInFasta(
      consolidateDuplicateProteinSeqsInFasta As Boolean,
      consolidateDupsIgnoreILDiff As Boolean,
      fixedFastaFilePath As String,
      proteinSequenceHashCount As Integer,
      proteinSeqHashInfo As IList(Of clsProteinHashInfo)) As Boolean

        Dim fsInFile As Stream
        Dim consolidatedFastaWriter As StreamWriter = Nothing

        Dim bytesRead As Int64
        Dim terminatorSize As Integer
        Dim percentComplete As Single
        Dim lineCountRead As Integer

        Dim fixedFastaFilePathTemp As String = String.Empty
        Dim lineIn As String

        Dim cachedProteinName As String = String.Empty
        Dim cachedProteinDescription As String = String.Empty
        Dim sbCachedProteinResidueLines = New StringBuilder(250)
        Dim sbCachedProteinResidues = New StringBuilder(250)

        ' This list contains the protein names that we will keep; values are the index values pointing into proteinSeqHashInfo
        ' If consolidateDuplicateProteinSeqsInFasta=False, this will contain all protein names
        ' If consolidateDuplicateProteinSeqsInFasta=True, we only keep the first name found for a given sequence
        Dim proteinNameFirst As clsNestedStringDictionary(Of Integer)

        ' This list keeps track of the protein names that have been written out to the new fasta file
        ' Keys are the protein names; values are the index of the entry in proteinSeqHashInfo()
        Dim proteinsWritten As clsNestedStringDictionary(Of Integer)

        ' This list contains the names of duplicate proteins; the hash values are the protein names of the master protein that has the same sequence
        Dim duplicateProteinList As clsNestedStringDictionary(Of String)

        Dim descriptionStartIndex As Integer

        Dim success As Boolean

        If proteinSequenceHashCount <= 0 Then
            Return True
        End If

        ''''''''''''''''''''''
        ' Processing Steps
        ''''''''''''''''''''''
        '
        ' Open fixedFastaFilePath with the fasta file reader
        ' Create a new fasta file with a writer

        ' For each protein, check whether it has duplicates
        ' If not, just write it out to the new fasta file

        ' If it does have duplicates and it is the master, then append the duplicate protein names to the end of the description for the protein
        '  and write out the name, new description, and sequence to the new fasta file

        ' Otherwise, check if it is a duplicate of a master protein
        ' If it is, then do not write the name, description, or sequence to the new fasta file

        Try
            fixedFastaFilePathTemp = fixedFastaFilePath & ".TempFixed"

            If File.Exists(fixedFastaFilePathTemp) Then
                File.Delete(fixedFastaFilePathTemp)
            End If

            File.Move(fixedFastaFilePath, fixedFastaFilePathTemp)
        Catch ex As Exception
            RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError,
             "Error renaming " & fixedFastaFilePath & " to " & fixedFastaFilePathTemp & ": " & ex.Message, String.Empty)
            OnErrorEvent("Error renaming fixed fasta to .tempfixed", ex)
            Return False
        End Try

        Dim fastaReader As StreamReader

        Try
            ' Open the file and read, at most, the first 100,000 characters to see if it contains CrLf or just Lf
            terminatorSize = DetermineLineTerminatorSize(fixedFastaFilePathTemp)

            ' Open the Fixed fasta file
            fsInFile = New FileStream(
               fixedFastaFilePathTemp,
               FileMode.Open,
               FileAccess.Read,
               FileShare.ReadWrite)

            fastaReader = New StreamReader(fsInFile)

        Catch ex As Exception
            RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError,
             "Error opening " & fixedFastaFilePathTemp & ": " & ex.Message, String.Empty)
            OnErrorEvent("Error opening fixedFastaFilePathTemp", ex)
            Return False
        End Try

        Try
            ' Create the new fasta file
            consolidatedFastaWriter = New StreamWriter(New FileStream(fixedFastaFilePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
        Catch ex As Exception
            RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError,
             "Error creating consolidated fasta output file " & fixedFastaFilePath & ": " & ex.Message, String.Empty)
            OnErrorEvent("Error creating consolidated fasta output file", ex)
        End Try

        Try
            ' Populate proteinNameFirst with the protein names in proteinSeqHashInfo().ProteinNameFirst
            proteinNameFirst = New clsNestedStringDictionary(Of Integer)(True, mProteinNameSpannerCharLength)

            ' Populate htDuplicateProteinList with the protein names in proteinSeqHashInfo().AdditionalProteins
            duplicateProteinList = New clsNestedStringDictionary(Of String)(True, mProteinNameSpannerCharLength)

            For index = 0 To proteinSequenceHashCount - 1
                With proteinSeqHashInfo(index)

                    If Not proteinNameFirst.ContainsKey(.ProteinNameFirst) Then
                        proteinNameFirst.Add(.ProteinNameFirst, index)
                    Else
                        ' .ProteinNameFirst is already present in proteinNameFirst
                        ' The fixed fasta file will only actually contain the first occurrence of .ProteinNameFirst, so we can effectively ignore this entry
                        ' but we should increment the DuplicateNameSkipCount

                    End If

                    If .AdditionalProteins.Count > 0 Then
                        For Each additionalProtein As String In .AdditionalProteins

                            If consolidateDuplicateProteinSeqsInFasta Then
                                ' Update the duplicate protein name list
                                If Not duplicateProteinList.ContainsKey(additionalProtein) Then
                                    duplicateProteinList.Add(additionalProtein, .ProteinNameFirst)
                                End If

                            Else
                                ' We are not consolidating proteins with the same sequence but different protein names
                                ' Append this entry to proteinNameFirst

                                If Not proteinNameFirst.ContainsKey(additionalProtein) Then
                                    proteinNameFirst.Add(additionalProtein, index)
                                Else
                                    ' .AdditionalProteins(dupIndex) is already present in proteinNameFirst
                                    ' Increment the DuplicateNameSkipCount
                                End If

                            End If

                        Next
                    End If
                End With
            Next index

            proteinsWritten = New clsNestedStringDictionary(Of Integer)(False, mProteinNameSpannerCharLength)

            Dim lastMemoryUsageReport = DateTime.UtcNow

            ' Parse each line in the file
            lineCountRead = 0
            bytesRead = 0
            mFixedFastaStats.DuplicateSequenceProteinsSkipped = 0

            Do While Not fastaReader.EndOfStream
                lineIn = fastaReader.ReadLine()
                bytesRead += lineIn.Length + terminatorSize

                If lineCountRead Mod 50 = 0 Then
                    If MyBase.AbortProcessing Then Exit Do

                    percentComplete = 75 + CType(bytesRead / CType(fastaReader.BaseStream.Length, Single) * 100.0, Single) / 4
                    MyBase.UpdateProgress("Consolidating duplicate proteins to create a new FASTA File (" & Math.Round(percentComplete, 0) & "% Done)", percentComplete)

                    If DateTime.UtcNow.Subtract(lastMemoryUsageReport).TotalMinutes >= 1 Then
                        lastMemoryUsageReport = DateTime.UtcNow
                        ReportMemoryUsage(proteinNameFirst, proteinsWritten, duplicateProteinList)
                    End If
                End If

                lineCountRead += 1

                If Not lineIn Is Nothing Then
                    If lineIn.Trim.Length > 0 Then
                        ' Note: Trim the start of the line (however, since this is a fixed fasta file it should not start with a space)
                        lineIn = lineIn.TrimStart

                        If lineIn.Chars(0) = mProteinLineStartChar Then
                            ' Protein entry line

                            If Not String.IsNullOrEmpty(cachedProteinName) Then
                                ' Write out the cached protein and it's residues

                                WriteCachedProtein(
                                 cachedProteinName, cachedProteinDescription,
                                 consolidatedFastaWriter, proteinSeqHashInfo,
                                 sbCachedProteinResidueLines, sbCachedProteinResidues,
                                 consolidateDuplicateProteinSeqsInFasta, consolidateDupsIgnoreILDiff,
                                 proteinNameFirst, duplicateProteinList,
                                 lineCountRead, proteinsWritten)

                                cachedProteinName = String.Empty
                                sbCachedProteinResidueLines.Length = 0
                                sbCachedProteinResidues.Length = 0
                            End If

                            ' Extract the protein name and description
                            SplitFastaProteinHeaderLine(lineIn, cachedProteinName, cachedProteinDescription, descriptionStartIndex)

                        Else
                            ' Protein residues
                            sbCachedProteinResidueLines.AppendLine(lineIn)
                            sbCachedProteinResidues.Append(lineIn.Trim())
                        End If
                    End If
                End If

            Loop

            If Not String.IsNullOrEmpty(cachedProteinName) Then
                ' Write out the cached protein and it's residues
                WriteCachedProtein(
                 cachedProteinName, cachedProteinDescription,
                 consolidatedFastaWriter, proteinSeqHashInfo,
                 sbCachedProteinResidueLines, sbCachedProteinResidues,
                 consolidateDuplicateProteinSeqsInFasta, consolidateDupsIgnoreILDiff,
                 proteinNameFirst, duplicateProteinList,
                 lineCountRead, proteinsWritten)
            End If

            ReportMemoryUsage(proteinNameFirst, proteinsWritten, duplicateProteinList)

            success = True

        Catch ex As Exception
            RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError,
             "Error writing to consolidated fasta file " & fixedFastaFilePath & ": " & ex.Message, String.Empty)
            OnErrorEvent("Error writing to consolidated fasta file", ex)
            Return False
        Finally
            Try
                If Not fastaReader Is Nothing Then fastaReader.Close()
                If Not consolidatedFastaWriter Is Nothing Then consolidatedFastaWriter.Close()

                Threading.Thread.Sleep(100)

                File.Delete(fixedFastaFilePathTemp)
            Catch ex As Exception
                ' Ignore errors here
                OnWarningEvent("Error closing file handles in CorrectForDuplicateProteinSeqsInFasta: " & ex.Message)
            End Try
        End Try

        Return success

    End Function

    Private Function ConstructFastaHeaderLine(ByRef proteinName As String, proteinDescription As String) As String

        If proteinName Is Nothing Then proteinName = "????"

        If String.IsNullOrWhiteSpace(proteinDescription) Then
            Return mProteinLineStartChar & proteinName
        Else
            Return mProteinLineStartChar & proteinName & " " & proteinDescription
        End If

    End Function

    Private Function ConstructStatsFilePath(outputFolderPath As String) As String

        Dim outFilePath As String = String.Empty

        Try
            ' Record the current time in now
            outFilePath = "FastaFileStats_" & DateTime.Now.ToString("yyyy-MM-dd") & ".txt"

            If Not outputFolderPath Is Nothing AndAlso outputFolderPath.Length > 0 Then
                outFilePath = Path.Combine(outputFolderPath, outFilePath)
            End If
        Catch ex As Exception
            If String.IsNullOrWhiteSpace(outFilePath) Then
                outFilePath = "FastaFileStats.txt"
            End If
        End Try

        Return outFilePath

    End Function

    Private Sub DeleteTempFiles()
        If Not mTempFilesToDelete Is Nothing AndAlso mTempFilesToDelete.Count > 0 Then
            For Each filePath In mTempFilesToDelete
                Try
                    If File.Exists(filePath) Then
                        File.Delete(filePath)
                    End If
                Catch ex As Exception
                    ' Ignore errors
                End Try
            Next
        End If
    End Sub

    Private Function DetermineLineTerminatorSize(inputFilePath As String) As Integer

        Dim endCharType As eLineEndingCharacters = Me.DetermineLineTerminatorType(inputFilePath)

        Select Case endCharType
            Case eLineEndingCharacters.CR
                Return 1
            Case eLineEndingCharacters.LF
                Return 1
            Case eLineEndingCharacters.CRLF
                Return 2
            Case eLineEndingCharacters.LFCR
                Return 2
        End Select

        Return 2

    End Function

    Private Function DetermineLineTerminatorType(inputFilePath As String) As eLineEndingCharacters
        Dim oneByte As Integer

        Dim endCharacterType As eLineEndingCharacters

        Try
            ' Open the input file and look for the first carriage return or line feed
            Using fsInFile = New FileStream(inputFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)

                Do While fsInFile.Position < fsInFile.Length AndAlso fsInFile.Position < 100000

                    oneByte = fsInFile.ReadByte()

                    If oneByte = 10 Then
                        ' Found linefeed
                        If fsInFile.Position < fsInFile.Length Then
                            oneByte = fsInFile.ReadByte()
                            If oneByte = 13 Then
                                ' LfCr
                                endCharacterType = eLineEndingCharacters.LFCR
                            Else
                                ' Lf only
                                endCharacterType = eLineEndingCharacters.LF
                            End If
                        Else
                            endCharacterType = eLineEndingCharacters.LF
                        End If
                        Exit Do
                    ElseIf oneByte = 13 Then
                        ' Found carriage return
                        If fsInFile.Position < fsInFile.Length Then
                            oneByte = fsInFile.ReadByte()
                            If oneByte = 10 Then
                                ' CrLf
                                endCharacterType = eLineEndingCharacters.CRLF
                            Else
                                ' Cr only
                                endCharacterType = eLineEndingCharacters.CR
                            End If
                        Else
                            endCharacterType = eLineEndingCharacters.CR
                        End If
                        Exit Do
                    End If

                Loop

            End Using

        Catch ex As Exception
            SetLocalErrorCode(eValidateFastaFileErrorCodes.ErrorVerifyingLinefeedAtEOF)
        End Try

        Return endCharacterType

    End Function

    ' Unused function
    ''Private Function ExtractListItem(list As String, item As Integer) As String
    ''    Dim items() As String
    ''    Dim item As String

    ''    item = String.Empty
    ''    If item >= 1 And Not list Is Nothing Then
    ''        items = list.Split(","c)
    ''        If items.Length >= item Then
    ''            item = items(item - 1)
    ''        End If
    ''    End If

    ''    Return item

    ''End Function

    Private Function NormalizeFileLineEndings(
      pathOfFileToFix As String,
      newFileName As String,
      desiredLineEndCharacterType As eLineEndingCharacters) As String

        Dim newEndChar As String = ControlChars.CrLf

        Dim endCharType As eLineEndingCharacters = Me.DetermineLineTerminatorType(pathOfFileToFix)

        Dim origEndCharCount As Integer

        If endCharType <> desiredLineEndCharacterType Then
            Select Case desiredLineEndCharacterType
                Case eLineEndingCharacters.CRLF
                    newEndChar = ControlChars.CrLf
                Case eLineEndingCharacters.CR
                    newEndChar = ControlChars.Cr
                Case eLineEndingCharacters.LF
                    newEndChar = ControlChars.Lf
                Case eLineEndingCharacters.LFCR
                    newEndChar = ControlChars.CrLf
            End Select

            Select Case endCharType
                Case eLineEndingCharacters.CR
                    origEndCharCount = 2
                Case eLineEndingCharacters.CRLF
                    origEndCharCount = 1
                Case eLineEndingCharacters.LF
                    origEndCharCount = 1
                Case eLineEndingCharacters.LFCR
                    origEndCharCount = 2
            End Select

            If Not Path.IsPathRooted(newFileName) Then
                newFileName = Path.Combine(Path.GetDirectoryName(pathOfFileToFix), Path.GetFileName(newFileName))
            End If

            Dim targetFile = New FileInfo(pathOfFileToFix)
            Dim fileSizeBytes = targetFile.Length

            Dim reader = targetFile.OpenText()

            Using writer = New StreamWriter(New FileStream(newFileName, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))

                Me.OnProgressUpdate("Normalizing Line Endings...", 0.0)

                Dim dataLine = reader.ReadLine
                Dim linesRead As Long = 0
                Do While Not dataLine Is Nothing
                    writer.Write(dataLine)
                    writer.Write(newEndChar)

                    Dim currentFilePos = dataLine.Length + origEndCharCount
                    linesRead += 1

                    If linesRead Mod 1000 = 0 Then
                        Me.OnProgressUpdate("Normalizing Line Endings (" &
                         Math.Round(CDbl(currentFilePos / fileSizeBytes * 100), 1).ToString &
                         " % complete", CSng(currentFilePos / fileSizeBytes * 100))
                    End If

                    dataLine = reader.ReadLine()
                Loop
                reader.Close()
            End Using

            Return newFileName
        Else
            Return pathOfFileToFix
        End If
    End Function

    Private Sub EvaluateRules(
      ruleDetails As IList(Of udtRuleDefinitionExtendedType),
      proteinName As String,
      textToTest As String,
      testTextOffsetInLine As Integer,
      entireLine As String,
      contextLength As Integer)

        Dim index As Integer
        Dim reMatch As Match
        Dim extraInfo As String
        Dim charIndexOfMatch As Integer

        For index = 0 To ruleDetails.Count - 1
            With ruleDetails(index)

                reMatch = .reRule.Match(textToTest)

                If (.RuleDefinition.MatchIndicatesProblem And reMatch.Success) OrElse
                 Not (.RuleDefinition.MatchIndicatesProblem And Not reMatch.Success) Then

                    If .RuleDefinition.DisplayMatchAsExtraInfo Then
                        extraInfo = reMatch.ToString
                    Else
                        extraInfo = String.Empty
                    End If

                    charIndexOfMatch = testTextOffsetInLine + reMatch.Index
                    If .RuleDefinition.Severity >= 5 Then
                        RecordFastaFileError(mLineCount, charIndexOfMatch, proteinName,
                         .RuleDefinition.CustomRuleID, extraInfo,
                         ExtractContext(entireLine, charIndexOfMatch, contextLength))
                    Else
                        RecordFastaFileWarning(mLineCount, charIndexOfMatch, proteinName,
                         .RuleDefinition.CustomRuleID, extraInfo,
                         ExtractContext(entireLine, charIndexOfMatch, contextLength))
                    End If

                End If

            End With
        Next index

    End Sub

    Private Function ExamineProteinName(
      ByRef proteinName As String,
      proteinNames As ISet(Of String),
      <Out> ByRef skipDuplicateProtein As Boolean,
      ByRef processingDuplicateOrInvalidProtein As Boolean) As String

        Dim duplicateName = proteinNames.Contains(proteinName)
        skipDuplicateProtein = False

        If duplicateName AndAlso mGenerateFixedFastaFile Then
            If mFixedFastaOptions.RenameProteinsWithDuplicateNames Then

                Dim letterToAppend = "b"c
                Dim numberToAppend = 0
                Dim newProteinName As String

                Do
                    newProteinName = proteinName & "-"c & letterToAppend
                    If numberToAppend > 0 Then
                        newProteinName &= numberToAppend.ToString
                    End If

                    duplicateName = proteinNames.Contains(newProteinName)

                    If duplicateName Then
                        ' Increment letterToAppend to the next letter and then try again to rename the protein
                        If letterToAppend = "z"c Then
                            ' We've reached "z"
                            ' Change back to "a" but increment numberToAppend
                            letterToAppend = "a"c
                            numberToAppend += 1
                        Else
                            ' letterToAppend = Chr(Asc(letterToAppend) + 1)
                            letterToAppend = Convert.ToChar(Convert.ToInt32(letterToAppend) + 1)
                        End If
                    End If
                Loop While duplicateName

                RecordFastaFileWarning(mLineCount, 1, proteinName, eMessageCodeConstants.RenamedProtein, "--> " & newProteinName, String.Empty)

                proteinName = String.Copy(newProteinName)
                mFixedFastaStats.DuplicateNameProteinsRenamed += 1
                skipDuplicateProtein = False

            ElseIf mFixedFastaOptions.KeepDuplicateNamedProteinsUnlessMatchingSequence Then
                skipDuplicateProtein = False
            Else
                skipDuplicateProtein = True
            End If

        End If

        If duplicateName Then
            If skipDuplicateProtein Or Not mGenerateFixedFastaFile Then
                RecordFastaFileError(mLineCount, 1, proteinName, eMessageCodeConstants.DuplicateProteinName)
                If mSaveBasicProteinHashInfoFile Then
                    processingDuplicateOrInvalidProtein = False
                Else
                    processingDuplicateOrInvalidProtein = True
                    mFixedFastaStats.DuplicateNameProteinsSkipped += 1
                End If
            Else
                RecordFastaFileWarning(mLineCount, 1, proteinName, eMessageCodeConstants.DuplicateProteinName)
                processingDuplicateOrInvalidProtein = False
            End If
        Else
            processingDuplicateOrInvalidProtein = False
        End If

        If Not proteinNames.Contains(proteinName) Then
            proteinNames.Add(proteinName)
        End If

        Return proteinName

    End Function

    Private Function ExtractContext(text As String, startIndex As Integer) As String
        Return ExtractContext(text, startIndex, DEFAULT_CONTEXT_LENGTH)
    End Function

    Private Function ExtractContext(text As String, startIndex As Integer, contextLength As Integer) As String
        ' Note that contextLength should be an odd number; if it isn't, we'll add 1 to it

        Dim contextStartIndex As Integer
        Dim contextEndIndex As Integer

        If contextLength Mod 2 = 0 Then
            contextLength += 1
        ElseIf contextLength < 1 Then
            contextLength = 1
        End If

        If text Is Nothing Then
            Return String.Empty
        ElseIf text.Length <= 1 Then
            Return text
        Else
            ' Define the start index for extracting the context from text
            contextStartIndex = startIndex - CInt((contextLength - 1) / 2)
            If contextStartIndex < 0 Then contextStartIndex = 0

            ' Define the end index for extracting the context from text
            contextEndIndex = Math.Max(startIndex + CInt((contextLength - 1) / 2), contextStartIndex + contextLength - 1)
            If contextEndIndex >= text.Length Then
                contextEndIndex = text.Length - 1
            End If

            ' Return the context portion of text
            Return text.Substring(contextStartIndex, contextEndIndex - contextStartIndex + 1)
        End If

    End Function

    Private Function FlattenArray(items As IEnumerable(Of String), sepChar As Char) As String
        If items Is Nothing Then
            Return String.Empty
        Else
            Return FlattenArray(items, items.Count, sepChar)
        End If
    End Function

    Private Function FlattenArray(items As IEnumerable(Of String), dataCount As Integer, sepChar As Char) As String
        Dim index As Integer
        Dim result As String

        If items Is Nothing Then
            Return String.Empty
        ElseIf items.Count = 0 OrElse dataCount <= 0 Then
            Return String.Empty
        Else
            If dataCount > items.Count Then
                dataCount = items.Count
            End If

            result = items(0)
            If result Is Nothing Then result = String.Empty

            For index = 1 To dataCount - 1
                If items(index) Is Nothing Then
                    result &= sepChar
                Else
                    result &= sepChar & items(index)
                End If
            Next index
            Return result
        End If

    End Function

    ''' <summary>
    ''' Convert a list of strings to a tab-delimited string
    ''' </summary>
    ''' <param name="dataValues"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function FlattenList(dataValues As IEnumerable(Of String)) As String
        Return FlattenArray(dataValues, ControlChars.Tab)
    End Function

    ''' <summary>
    ''' Find the first space (or first tab) in the protein header line
    ''' </summary>
    ''' <param name="headerLine"></param>
    ''' <returns></returns>
    ''' <remarks>Used for determining protein name</remarks>
    Private Function GetBestSpaceIndex(headerLine As String) As Integer

        Dim spaceIndex = headerLine.IndexOf(" "c)
        Dim tabIndex = headerLine.IndexOf(ControlChars.Tab)

        If spaceIndex = 1 Then
            ' Space found directly after the > symbol
        ElseIf tabIndex > 0 Then
            If tabIndex = 1 Then
                ' Tab character found directly after the > symbol
                spaceIndex = tabIndex
            Else
                ' Tab character found; does it separate the protein name and description?
                If spaceIndex <= 0 OrElse (spaceIndex > 0 AndAlso tabIndex < spaceIndex) Then
                    spaceIndex = tabIndex
                End If
            End If
        End If

        Return spaceIndex

    End Function

    Public Overrides Function GetDefaultExtensionsToParse() As IList(Of String)
        Dim extensionsToParse = New List(Of String) From {
            ".fasta"
        }

        Return extensionsToParse

    End Function

    Public Overrides Function GetErrorMessage() As String
        ' Returns "" if no error

        Dim errorMessage As String

        If MyBase.ErrorCode = ProcessFilesErrorCodes.LocalizedError Or
           MyBase.ErrorCode = ProcessFilesErrorCodes.NoError Then
            Select Case mLocalErrorCode
                Case eValidateFastaFileErrorCodes.NoError
                    errorMessage = ""
                Case eValidateFastaFileErrorCodes.OptionsSectionNotFound
                    errorMessage = "The section " & XML_SECTION_OPTIONS & " was not found in the parameter file"
                Case eValidateFastaFileErrorCodes.ErrorReadingInputFile
                    errorMessage = "Error reading input file"
                Case eValidateFastaFileErrorCodes.ErrorCreatingStatsFile
                    errorMessage = "Error creating stats output file"
                Case eValidateFastaFileErrorCodes.ErrorVerifyingLinefeedAtEOF
                    errorMessage = "Error verifying linefeed at end of file"
                Case eValidateFastaFileErrorCodes.UnspecifiedError
                    errorMessage = "Unspecified localized error"
                Case Else
                    ' This shouldn't happen
                    errorMessage = "Unknown error state"
            End Select
        Else
            errorMessage = MyBase.GetBaseClassErrorMessage()
        End If

        Return errorMessage

    End Function

    Private Function GetFileErrorTextByIndex(fileErrorIndex As Integer, sepChar As String) As String

        Dim proteinName As String

        If mFileErrorCount <= 0 Or fileErrorIndex < 0 Or fileErrorIndex >= mFileErrorCount Then
            Return String.Empty
        Else
            With mFileErrors(fileErrorIndex)
                If .ProteinName Is Nothing OrElse .ProteinName.Length = 0 Then
                    proteinName = "N/A"
                Else
                    proteinName = String.Copy(.ProteinName)
                End If

                Return LookupMessageType(eMsgTypeConstants.ErrorMsg) & sepChar &
                 "Line " & .LineNumber.ToString & sepChar &
                 "Col " & .ColNumber.ToString & sepChar &
                 proteinName & sepChar &
                 LookupMessageDescription(.MessageCode, .ExtraInfo) & sepChar & .Context

            End With
        End If

    End Function

    Private Function GetFileErrorByIndex(fileErrorIndex As Integer) As udtMsgInfoType

        If mFileErrorCount <= 0 Or fileErrorIndex < 0 Or fileErrorIndex >= mFileErrorCount Then
            Return New udtMsgInfoType
        Else
            Return mFileErrors(fileErrorIndex)
        End If

    End Function

    ''' <summary>
    ''' Retrieve the errors reported by the validator
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>Used by clsCustomValidateFastaFiles</remarks>
    Protected Function GetFileErrors() As List(Of udtMsgInfoType)

        Dim fileErrors = New List(Of udtMsgInfoType)

        For i = 0 To mFileErrorCount - 1
            fileErrors.Add(mFileErrors(i))
        Next

        Return fileErrors

    End Function


    Private Function GetFileWarningTextByIndex(fileWarningIndex As Integer, sepChar As String) As String
        Dim proteinName As String

        If mFileWarningCount <= 0 Or fileWarningIndex < 0 Or fileWarningIndex >= mFileWarningCount Then
            Return String.Empty
        Else
            With mFileWarnings(fileWarningIndex)
                If .ProteinName Is Nothing OrElse .ProteinName.Length = 0 Then
                    proteinName = "N/A"
                Else
                    proteinName = String.Copy(.ProteinName)
                End If

                Return LookupMessageType(eMsgTypeConstants.WarningMsg) & sepChar &
                 "Line " & .LineNumber.ToString & sepChar &
                 "Col " & .ColNumber.ToString & sepChar &
                 proteinName & sepChar &
                 LookupMessageDescription(.MessageCode, .ExtraInfo) &
                 sepChar & .Context

            End With
        End If

    End Function

    Private Function GetFileWarningByIndex(fileWarningIndex As Integer) As udtMsgInfoType

        If mFileWarningCount <= 0 Or fileWarningIndex < 0 Or fileWarningIndex >= mFileWarningCount Then
            Return New udtMsgInfoType
        Else
            Return mFileWarnings(fileWarningIndex)
        End If

    End Function

    ''' <summary>
    ''' Retrieve the warnings reported by the validator
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>Used by clsCustomValidateFastaFiles</remarks>
    Protected Function GetFileWarnings() As List(Of udtMsgInfoType)

        Dim fileWarnings = New List(Of udtMsgInfoType)

        For i = 0 To mFileWarningCount - 1
            fileWarnings.Add(mFileWarnings(i))
        Next

        Return fileWarnings

    End Function

    Private Function GetProcessMemoryUsageWithTimestamp() As String
        Return GetTimeStamp() & ControlChars.Tab & clsMemoryUsageLogger.GetProcessMemoryUsageMB.ToString("0.0") & " MB in use"
    End Function

    Private Function GetTimeStamp() As String
        ' Record the current time
        Return DateTime.Now.ToShortDateString & " " & DateTime.Now.ToLongTimeString
    End Function

    Private Sub InitializeLocalVariables()

        mLocalErrorCode = eValidateFastaFileErrorCodes.NoError

        Me.OptionSwitch(SwitchOptions.AddMissingLinefeedatEOF) = False

        Me.MaximumFileErrorsToTrack = 5

        Me.MinimumProteinNameLength = DEFAULT_MINIMUM_PROTEIN_NAME_LENGTH
        Me.MaximumProteinNameLength = DEFAULT_MAXIMUM_PROTEIN_NAME_LENGTH
        Me.MaximumResiduesPerLine = DEFAULT_MAXIMUM_RESIDUES_PER_LINE
        Me.ProteinLineStartChar = DEFAULT_PROTEIN_LINE_START_CHAR

        Me.OptionSwitch(SwitchOptions.AllowAsteriskInResidues) = False
        Me.OptionSwitch(SwitchOptions.AllowDashInResidues) = False
        Me.OptionSwitch(SwitchOptions.WarnBlankLinesBetweenProteins) = False
        Me.OptionSwitch(SwitchOptions.WarnLineStartsWithSpace) = True

        Me.OptionSwitch(SwitchOptions.CheckForDuplicateProteinNames) = True
        Me.OptionSwitch(SwitchOptions.CheckForDuplicateProteinSequences) = True

        Me.OptionSwitch(SwitchOptions.GenerateFixedFASTAFile) = False

        Me.OptionSwitch(SwitchOptions.FixedFastaSplitOutMultipleRefsForKnownAccession) = True
        Me.OptionSwitch(SwitchOptions.SplitOutMultipleRefsInProteinName) = False

        Me.OptionSwitch(SwitchOptions.FixedFastaRenameDuplicateNameProteins) = False
        Me.OptionSwitch(SwitchOptions.FixedFastaKeepDuplicateNamedProteins) = False

        Me.OptionSwitch(SwitchOptions.FixedFastaConsolidateDuplicateProteinSeqs) = False
        Me.OptionSwitch(SwitchOptions.FixedFastaConsolidateDupsIgnoreILDiff) = False

        Me.OptionSwitch(SwitchOptions.FixedFastaTruncateLongProteinNames) = True
        Me.OptionSwitch(SwitchOptions.FixedFastaWrapLongResidueLines) = True
        Me.OptionSwitch(SwitchOptions.FixedFastaRemoveInvalidResidues) = False

        Me.OptionSwitch(SwitchOptions.SaveProteinSequenceHashInfoFiles) = False

        Me.OptionSwitch(SwitchOptions.SaveBasicProteinHashInfoFile) = False


        mProteinNameFirstRefSepChars = DEFAULT_PROTEIN_NAME_FIRST_REF_SEP_CHARS.ToCharArray
        mProteinNameSubsequentRefSepChars = DEFAULT_PROTEIN_NAME_SUBSEQUENT_REF_SEP_CHARS.ToCharArray

        mFixedFastaOptions.LongProteinNameSplitChars = New Char() {DEFAULT_LONG_PROTEIN_NAME_SPLIT_CHAR}
        mFixedFastaOptions.ProteinNameInvalidCharsToRemove = New Char() {}          ' Default to an empty character array

        SetDefaultRules()

        ResetStructures()
        mFastaFilePath = String.Empty

        mMemoryUsageLogger = New clsMemoryUsageLogger(String.Empty)
        mProcessMemoryUsageMBAtStart = clsMemoryUsageLogger.GetProcessMemoryUsageMB()

        ' ReSharper disable once VbUnreachableCode
        If REPORT_DETAILED_MEMORY_USAGE Then
            ' mMemoryUsageMBAtStart = mMemoryUsageLogger.GetFreeMemoryMB()
            Console.WriteLine(MEM_USAGE_PREFIX & mMemoryUsageLogger.GetMemoryUsageHeader())
            Console.WriteLine(MEM_USAGE_PREFIX & mMemoryUsageLogger.GetMemoryUsageSummary())
        End If

        mTempFilesToDelete = New List(Of String)

    End Sub

    Private Sub InitializeRuleDetails(
      ByRef ruleDefinitions() As udtRuleDefinitionType,
      ByRef ruleDetails() As udtRuleDefinitionExtendedType)
        Dim index As Integer

        If ruleDefinitions Is Nothing OrElse ruleDefinitions.Length = 0 Then
            ReDim ruleDetails(-1)
        Else
            ReDim ruleDetails(ruleDefinitions.Length - 1)

            For index = 0 To ruleDefinitions.Length - 1
                Try
                    With ruleDetails(index)
                        .RuleDefinition = ruleDefinitions(index)
                        .reRule = New Regex(
                         .RuleDefinition.MatchRegEx,
                         RegexOptions.Singleline Or
                         RegexOptions.Compiled)
                        .Valid = True
                    End With
                Catch ex As Exception
                    ' Ignore the error, but mark .Valid = false
                    ruleDetails(index).Valid = False
                End Try
            Next index
        End If

    End Sub

    Private Function LoadExistingProteinHashFile(
      proteinHashFilePath As String,
      <Out> ByRef preloadedProteinNamesToKeep As clsNestedStringIntList) As Boolean

        ' List of protein names to keep
        ' Keys are protein names, values are the number of entries written to the fixed fasta file for the given protein name
        preloadedProteinNamesToKeep = Nothing

        Try

            Dim proteinHashFile = New FileInfo(proteinHashFilePath)

            If Not proteinHashFile.Exists Then
                ShowErrorMessage("Protein hash file not found: " & proteinHashFilePath)
                Return False
            End If

            ' Sort the protein has file on the Sequence_Hash column
            ' First cache the column names from the header line

            Dim headerInfo = New Dictionary(Of String, Integer)(StringComparer.InvariantCultureIgnoreCase)
            Dim proteinHashFileLines As Long = 0
            Dim cachedHeaderLine As String = String.Empty

            ShowMessage("Examining pre-existing protein hash file to count the number of entries: " & Path.GetFileName(proteinHashFilePath))
            Dim lastStatus = DateTime.UtcNow
            Dim progressDotShown = False

            Using hashFileReader = New StreamReader(New FileStream(proteinHashFile.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                If Not hashFileReader.EndOfStream Then
                    cachedHeaderLine = hashFileReader.ReadLine()

                    Dim headerNames = cachedHeaderLine.Split(ControlChars.Tab)

                    For colIndex = 0 To headerNames.Count - 1
                        headerInfo.Add(headerNames(colIndex), colIndex)
                    Next
                    proteinHashFileLines = 1
                End If

                While Not hashFileReader.EndOfStream
                    hashFileReader.ReadLine()
                    proteinHashFileLines += 1

                    If proteinHashFileLines Mod 10000 = 0 Then
                        If DateTime.UtcNow.Subtract(lastStatus).TotalSeconds >= 10 Then
                            Console.Write(".")
                            progressDotShown = True
                            lastStatus = DateTime.UtcNow
                        End If
                    End If
                End While
            End Using

            If progressDotShown Then Console.WriteLine()

            If headerInfo.Count = 0 Then
                ShowErrorMessage("Protein hash file is empty: " + proteinHashFilePath)
                Return False
            End If

            Dim sequenceHashColumnIndex As Integer
            If Not headerInfo.TryGetValue(SEQUENCE_HASH_COLUMN, sequenceHashColumnIndex) Then
                ShowErrorMessage("Protein hash file is missing the " & SEQUENCE_HASH_COLUMN & " column: " & proteinHashFilePath)
                Return False
            End If

            Dim proteinNameColumnIndex As Integer
            If Not headerInfo.TryGetValue(PROTEIN_NAME_COLUMN, proteinNameColumnIndex) Then
                ShowErrorMessage("Protein hash file is missing the " & PROTEIN_NAME_COLUMN & " column: " & proteinHashFilePath)
                Return False
            End If

            Dim sequenceLengthColumnIndex As Integer
            If Not headerInfo.TryGetValue(SEQUENCE_LENGTH_COLUMN, sequenceLengthColumnIndex) Then
                ShowErrorMessage("Protein hash file is missing the " & SEQUENCE_LENGTH_COLUMN & " column: " & proteinHashFilePath)
                Return False
            End If

            Dim baseHashFileName = Path.GetFileNameWithoutExtension(proteinHashFile.Name)
            Dim sortedProteinHashSuffix As String
            Dim proteinHashFilenameSuffixNoExtension = Path.GetFileNameWithoutExtension(PROTEIN_HASHES_FILENAME_SUFFIX)

            If baseHashFileName.EndsWith(proteinHashFilenameSuffixNoExtension) Then
                baseHashFileName = baseHashFileName.Substring(0, baseHashFileName.Length - proteinHashFilenameSuffixNoExtension.Length)
                sortedProteinHashSuffix = proteinHashFilenameSuffixNoExtension
            Else
                sortedProteinHashSuffix = String.Empty
            End If

            Dim baseDataFileName = Path.Combine(proteinHashFile.Directory.FullName, baseHashFileName)

            ' Note: do not add sortedProteinHashFilePath to mTempFilesToDelete
            Dim sortedProteinHashFilePath = baseDataFileName & sortedProteinHashSuffix & "_Sorted.tmp"

            Dim sortedProteinHashFile = New FileInfo(sortedProteinHashFilePath)
            Dim sortedHashFileLines As Long = 0
            Dim sortRequired = True

            If sortedProteinHashFile.Exists Then
                lastStatus = DateTime.UtcNow
                progressDotShown = False
                ShowMessage("Validating existing sorted protein hash file: " + sortedProteinHashFile.Name)

                ' The sorted file exists; if it has the same number of lines as the sort file, assume that it is complete
                Using sortedHashFileReader = New StreamReader(New FileStream(sortedProteinHashFile.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    If Not sortedHashFileReader.EndOfStream Then
                        Dim headerLine = sortedHashFileReader.ReadLine()
                        If Not String.Equals(headerLine, cachedHeaderLine) Then
                            sortedHashFileLines = -1
                        Else
                            sortedHashFileLines = 1
                        End If

                    End If

                    If sortedHashFileLines > 0 Then
                        While Not sortedHashFileReader.EndOfStream
                            sortedHashFileReader.ReadLine()
                            sortedHashFileLines += 1

                            If sortedHashFileLines Mod 10000 = 0 Then
                                If DateTime.UtcNow.Subtract(lastStatus).TotalSeconds >= 10 Then
                                    Console.Write((sortedHashFileLines / proteinHashFileLines * 100.0).ToString("0") & "% ")
                                    progressDotShown = True
                                    lastStatus = DateTime.UtcNow
                                End If
                            End If
                        End While
                    End If

                End Using

                If progressDotShown Then Console.WriteLine()

                If sortedHashFileLines = proteinHashFileLines Then
                    sortRequired = False
                Else
                    If sortedHashFileLines < 0 Then
                        ShowMessage("Existing sorted hash file has an incorrect header; re-creating it")
                    Else
                        ShowMessage(String.Format("Existing sorted hash file has fewer lines ({0}) than the original ({1}); re-creating it",
                                                  sortedHashFileLines, proteinHashFileLines))
                    End If

                    sortedProteinHashFile.Delete()
                    Threading.Thread.Sleep(50)
                End If
            End If

            ' Create the sorted protein sequence hash file if necessary
            If sortRequired Then
                Console.WriteLine()
                ShowMessage("Sorting the existing protein hash file to create " + Path.GetFileName(sortedProteinHashFilePath))
                Dim sortHashSuccess = SortFile(proteinHashFile, sequenceHashColumnIndex, sortedProteinHashFilePath)
                If Not sortHashSuccess Then
                    Return False
                End If
            End If

            ShowMessage("Determining the best spanner length for protein names")

            ' Examine the protein names in the sequence hash file to determine the appropriate spanner length for the names
            Dim spannerCharLength = clsNestedStringIntList.AutoDetermineSpannerCharLength(proteinHashFile, proteinNameColumnIndex, True)
            Const RAISE_EXCEPTION_IF_ADDED_DATA_NOT_SORTED = True

            ' List of protein names to keep
            preloadedProteinNamesToKeep = New clsNestedStringIntList(spannerCharLength, RAISE_EXCEPTION_IF_ADDED_DATA_NOT_SORTED)

            lastStatus = DateTime.UtcNow
            progressDotShown = False
            Console.WriteLine()
            ShowMessage("Finding the name of the first protein for each protein hash")

            Dim currentHash = String.Empty
            Dim currentHashSeqLength = String.Empty
            Dim proteinNamesCurrentHash = New SortedSet(Of String)

            Dim linesRead = 0
            Dim proteinNamesUnsortedCount = 0

            Dim proteinNamesUnsortedFilePath = baseDataFileName & "_ProteinNamesUnsorted.tmp"
            Dim proteinNamesToKeepFilePath = baseDataFileName & "_ProteinNamesToKeep.tmp"
            Dim uniqueProteinSeqsFilePath = baseDataFileName & "_UniqueProteinSeqs.txt"
            Dim uniqueProteinSeqDuplicateFilePath = baseDataFileName & "_UniqueProteinSeqDuplicates.txt"

            mTempFilesToDelete.Add(proteinNamesUnsortedFilePath)
            mTempFilesToDelete.Add(proteinNamesToKeepFilePath)

            ' Write the first protein name for each sequence hash to a new file
            ' Also create the _UniqueProteinSeqs.txt and _UniqueProteinSeqDuplicates.txt files

            Using sortedHashFileReader = New StreamReader(New FileStream(sortedProteinHashFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)),
                proteinNamesUnsortedWriter = New StreamWriter(New FileStream(proteinNamesUnsortedFilePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)),
                proteinNamesToKeepWriter = New StreamWriter(New FileStream(proteinNamesToKeepFilePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)),
                uniqueProteinSeqsWriter = New StreamWriter(New FileStream(uniqueProteinSeqsFilePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)),
                uniqueProteinSeqDuplicateWriter = New StreamWriter(New FileStream(uniqueProteinSeqDuplicateFilePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))

                proteinNamesUnsortedWriter.WriteLine(FlattenList(New List(Of String) From {PROTEIN_NAME_COLUMN, SEQUENCE_HASH_COLUMN}))

                proteinNamesToKeepWriter.WriteLine(PROTEIN_NAME_COLUMN)

                Dim headerColumnsProteinSeqs = New List(Of String) From {
                    "Sequence_Index",
                    "Protein_Name_First",
                    SEQUENCE_LENGTH_COLUMN,
                    "Sequence_Hash",
                    "Protein_Count",
                    "Duplicate_Proteins"}

                uniqueProteinSeqsWriter.WriteLine(FlattenList(headerColumnsProteinSeqs))

                Dim headerColumnsSeqDups = New List(Of String) From {
                    "Sequence_Index",
                    "Protein_Name_First",
                    SEQUENCE_LENGTH_COLUMN,
                    "Duplicate_Protein"}

                uniqueProteinSeqDuplicateWriter.WriteLine(FlattenList(headerColumnsSeqDups))

                If Not sortedHashFileReader.EndOfStream Then
                    ' Read the header line
                    sortedHashFileReader.ReadLine()
                End If

                Dim currentSequenceIndex = 0

                While Not sortedHashFileReader.EndOfStream
                    Dim dataLine = sortedHashFileReader.ReadLine()
                    linesRead += 1

                    Dim dataValues = dataLine.Split(ControlChars.Tab)

                    Dim proteinName = dataValues(proteinNameColumnIndex)
                    Dim proteinHash = dataValues(sequenceHashColumnIndex)

                    proteinNamesUnsortedWriter.WriteLine(FlattenList(New List(Of String) From {proteinName, proteinHash}))
                    proteinNamesUnsortedCount += 1

                    If String.Equals(currentHash, proteinHash) Then
                        ' Existing sequence hash
                        If Not proteinNamesCurrentHash.Contains(proteinName) Then
                            proteinNamesCurrentHash.Add(proteinName)
                        End If
                    Else
                        ' New sequence hash found

                        ' First write out the data for the last hash
                        If Not String.IsNullOrEmpty(currentHash) Then
                            WriteCachedProteinHashMetadata(
                                proteinNamesToKeepWriter,
                                uniqueProteinSeqsWriter,
                                uniqueProteinSeqDuplicateWriter,
                                currentSequenceIndex,
                                currentHash,
                                currentHashSeqLength,
                                proteinNamesCurrentHash)
                        End If

                        ' Update the currentHash values
                        currentSequenceIndex += 1
                        currentHash = String.Copy(proteinHash)
                        currentHashSeqLength = dataValues(sequenceLengthColumnIndex)

                        proteinNamesCurrentHash.Clear()
                        proteinNamesCurrentHash.Add(proteinName)

                    End If

                    If linesRead Mod 10000 = 0 Then
                        If DateTime.UtcNow.Subtract(lastStatus).TotalSeconds >= 10 Then
                            Console.Write((linesRead / proteinHashFileLines * 100.0).ToString("0") & "% ")
                            progressDotShown = True
                            lastStatus = DateTime.UtcNow
                        End If
                    End If

                End While

                ' Write out the data for the last hash
                If Not String.IsNullOrEmpty(currentHash) Then
                    WriteCachedProteinHashMetadata(
                        proteinNamesToKeepWriter,
                        uniqueProteinSeqsWriter,
                        uniqueProteinSeqDuplicateWriter,
                        currentSequenceIndex,
                        currentHash,
                        currentHashSeqLength,
                        proteinNamesCurrentHash)
                End If

            End Using

            If progressDotShown Then Console.WriteLine()
            Console.WriteLine()

            ' Sort the protein names to keep file
            Dim sortedProteinNamesToKeepFilePath = baseDataFileName & "_ProteinsToKeepSorted.tmp"
            mTempFilesToDelete.Add(sortedProteinNamesToKeepFilePath)

            ShowMessage("Sorting the protein names to keep file to create " + Path.GetFileName(sortedProteinNamesToKeepFilePath))
            Dim sortProteinNamesToKeepSuccess = SortFile(New FileInfo(proteinNamesToKeepFilePath), 0, sortedProteinNamesToKeepFilePath)

            If Not sortProteinNamesToKeepSuccess Then
                Return False
            End If

            Dim lastProteinAdded = String.Empty

            ' Read the sorted protein names to keep file and cache the protein names in memory
            Using sortedNamesFileReader = New StreamReader(New FileStream(sortedProteinNamesToKeepFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                ' Skip the header line
                sortedNamesFileReader.ReadLine()

                While Not sortedNamesFileReader.EndOfStream
                    Dim proteinName = sortedNamesFileReader.ReadLine()

                    If String.Equals(lastProteinAdded, proteinName) Then
                        Continue While
                    End If

                    ' Store the protein name, plus a 0
                    ' The stored value will be incremented when this protein name is encountered by the validator
                    preloadedProteinNamesToKeep.Add(proteinName, 0)

                    lastProteinAdded = String.Copy(proteinName)
                End While
            End Using

            ' Confirm that the data is sorted
            preloadedProteinNamesToKeep.Sort()

            ShowMessage("Cached " & preloadedProteinNamesToKeep.Count.ToString("#,##0") & " protein names into memory")
            ShowMessage("The fixed FASTA file will only contain entries for these proteins")

            ' Sort the protein names file so that we can check for duplicate protein names
            Dim sortedProteinNamesFilePath = baseDataFileName & "_ProteinNamesSorted.tmp"
            mTempFilesToDelete.Add(sortedProteinNamesFilePath)

            Console.WriteLine()
            ShowMessage("Sorting the protein names file to create " + Path.GetFileName(sortedProteinNamesFilePath))
            Dim sortProteinNamesSuccess = SortFile(New FileInfo(proteinNamesUnsortedFilePath), 0, sortedProteinNamesFilePath)

            If Not sortProteinNamesSuccess Then
                Return False
            End If

            Dim proteinNameSummaryFilePath = baseDataFileName & "_ProteinNameSummary.txt"

            ' We can now safely delete some files to free up disk space
            Try
                Threading.Thread.Sleep(100)
                File.Delete(proteinNamesUnsortedFilePath)
            Catch ex As Exception
                ' Ignore errors here
            End Try

            ' Look for duplicate protein names
            ' In addition, create a new file with all protein names plus also two new columns: First_Protein_For_Hash and Duplicate_Name

            Using sortedProteinNamesReader = New StreamReader(New FileStream(sortedProteinNamesFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)),
                  proteinNameSummaryWriter = New StreamWriter(New FileStream(proteinNameSummaryFilePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))

                ' Skip the header line
                sortedProteinNamesReader.ReadLine()

                Dim proteinNameHeaderColumns = New List(Of String) From {
                    PROTEIN_NAME_COLUMN,
                    SEQUENCE_HASH_COLUMN,
                    "First_Protein_For_Hash",
                    "Duplicate_Name"}

                proteinNameSummaryWriter.WriteLine(FlattenList(proteinNameHeaderColumns))

                Dim lastProtein = String.Empty
                Dim warningShown = False
                Dim duplicateCount = 0

                linesRead = 0
                lastStatus = DateTime.UtcNow
                progressDotShown = False

                While Not sortedProteinNamesReader.EndOfStream
                    Dim dataLine = sortedProteinNamesReader.ReadLine()
                    Dim dataValues = dataLine.Split(ControlChars.Tab)

                    Dim currentProtein As String = dataValues(0)
                    Dim sequenceHash As String = dataValues(1)

                    Dim firstProteinForHash = preloadedProteinNamesToKeep.Contains(currentProtein)
                    Dim duplicateProtein = False

                    If String.Equals(lastProtein, currentProtein) Then
                        duplicateProtein = True

                        If Not warningShown Then
                            ShowMessage("WARNING: the protein hash file has duplicate protein names: " & Path.GetFileName(proteinHashFilePath))
                            warningShown = True
                        End If
                        duplicateCount += 1

                        If duplicateCount < 10 Then
                            ShowMessage("  ... duplicate protein name: " + currentProtein)
                        ElseIf duplicateCount = 10 Then
                            ShowMessage("  ... additional duplicates will not be shown ...")
                        End If
                    End If

                    lastProtein = String.Copy(currentProtein)

                    Dim dataToWrite = New List(Of String) From {
                        currentProtein,
                        sequenceHash,
                        BoolToStringInt(firstProteinForHash),
                        BoolToStringInt(duplicateProtein)}

                    proteinNameSummaryWriter.WriteLine(FlattenList(dataToWrite))

                    If linesRead Mod 10000 = 0 Then
                        If DateTime.UtcNow.Subtract(lastStatus).TotalSeconds >= 10 Then
                            Console.Write((linesRead / proteinNamesUnsortedCount * 100.0).ToString("0") & "% ")
                            progressDotShown = True
                            lastStatus = DateTime.UtcNow
                        End If
                    End If

                End While

                If duplicateCount > 0 Then
                    ShowMessage("WARNING: Found " & duplicateCount.ToString("#,##0") & " duplicate protein names in the protein hash file")
                End If
            End Using

            If progressDotShown Then
                Console.WriteLine()
            End If

            Console.WriteLine()

            Return True
        Catch ex As Exception
            OnErrorEvent("Error in LoadExistingProteinHashFile", ex)
            Return False

        End Try

    End Function

    Public Function LoadParameterFileSettings(parameterFilePath As String) As Boolean

        Dim settingsFile As New XmlSettingsFileAccessor

        Dim customRulesLoaded As Boolean
        Dim success As Boolean

        Dim characterList As String

        Try

            If parameterFilePath Is Nothing OrElse parameterFilePath.Length = 0 Then
                ' No parameter file specified; nothing to load
                Return True
            End If

            If Not File.Exists(parameterFilePath) Then
                ' See if parameterFilePath points to a file in the same directory as the application
                parameterFilePath = Path.Combine(
                 Path.GetDirectoryName(Reflection.Assembly.GetExecutingAssembly().Location),
                  Path.GetFileName(parameterFilePath))
                If Not File.Exists(parameterFilePath) Then
                    MyBase.SetBaseClassErrorCode(ProcessFilesErrorCodes.ParameterFileNotFound)
                    Return False
                End If
            End If

            If settingsFile.LoadSettings(parameterFilePath) Then
                If Not settingsFile.SectionPresent(XML_SECTION_OPTIONS) Then
                    OnWarningEvent("The node '<section name=""" & XML_SECTION_OPTIONS & """> was not found in the parameter file: " & parameterFilePath)
                    MyBase.SetBaseClassErrorCode(ProcessFilesErrorCodes.InvalidParameterFile)
                    Return False
                Else
                    ' Read customized settings

                    Me.OptionSwitch(SwitchOptions.AddMissingLinefeedatEOF) =
                     settingsFile.GetParam(XML_SECTION_OPTIONS, "AddMissingLinefeedAtEOF",
                     Me.OptionSwitch(SwitchOptions.AddMissingLinefeedatEOF))
                    Me.OptionSwitch(SwitchOptions.AllowAsteriskInResidues) =
                     settingsFile.GetParam(XML_SECTION_OPTIONS, "AllowAsteriskInResidues",
                     Me.OptionSwitch(SwitchOptions.AllowAsteriskInResidues))
                    Me.OptionSwitch(SwitchOptions.AllowDashInResidues) =
                     settingsFile.GetParam(XML_SECTION_OPTIONS, "AllowDashInResidues",
                     Me.OptionSwitch(SwitchOptions.AllowDashInResidues))
                    Me.OptionSwitch(SwitchOptions.CheckForDuplicateProteinNames) =
                     settingsFile.GetParam(XML_SECTION_OPTIONS, "CheckForDuplicateProteinNames",
                     Me.OptionSwitch(SwitchOptions.CheckForDuplicateProteinNames))
                    Me.OptionSwitch(SwitchOptions.CheckForDuplicateProteinSequences) =
                     settingsFile.GetParam(XML_SECTION_OPTIONS, "CheckForDuplicateProteinSequences",
                     Me.OptionSwitch(SwitchOptions.CheckForDuplicateProteinSequences))

                    Me.OptionSwitch(SwitchOptions.SaveProteinSequenceHashInfoFiles) =
                     settingsFile.GetParam(XML_SECTION_OPTIONS, "SaveProteinSequenceHashInfoFiles",
                     Me.OptionSwitch(SwitchOptions.SaveProteinSequenceHashInfoFiles))

                    Me.OptionSwitch(SwitchOptions.SaveBasicProteinHashInfoFile) =
                     settingsFile.GetParam(XML_SECTION_OPTIONS, "SaveBasicProteinHashInfoFile",
                     Me.OptionSwitch(SwitchOptions.SaveBasicProteinHashInfoFile))

                    Me.MaximumFileErrorsToTrack = settingsFile.GetParam(XML_SECTION_OPTIONS,
                     "MaximumFileErrorsToTrack", Me.MaximumFileErrorsToTrack)
                    Me.MinimumProteinNameLength = settingsFile.GetParam(XML_SECTION_OPTIONS,
                     "MinimumProteinNameLength", Me.MinimumProteinNameLength)
                    Me.MaximumProteinNameLength = settingsFile.GetParam(XML_SECTION_OPTIONS,
                     "MaximumProteinNameLength", Me.MaximumProteinNameLength)
                    Me.MaximumResiduesPerLine = settingsFile.GetParam(XML_SECTION_OPTIONS,
                     "MaximumResiduesPerLine", Me.MaximumResiduesPerLine)

                    Me.OptionSwitch(SwitchOptions.WarnBlankLinesBetweenProteins) =
                     settingsFile.GetParam(XML_SECTION_OPTIONS, "WarnBlankLinesBetweenProteins",
                     Me.OptionSwitch(SwitchOptions.WarnBlankLinesBetweenProteins))
                    Me.OptionSwitch(SwitchOptions.WarnLineStartsWithSpace) =
                     settingsFile.GetParam(XML_SECTION_OPTIONS, "WarnLineStartsWithSpace",
                     Me.OptionSwitch(SwitchOptions.WarnLineStartsWithSpace))

                    Me.OptionSwitch(SwitchOptions.OutputToStatsFile) =
                     settingsFile.GetParam(XML_SECTION_OPTIONS, "OutputToStatsFile",
                     Me.OptionSwitch(SwitchOptions.OutputToStatsFile))

                    Me.OptionSwitch(SwitchOptions.NormalizeFileLineEndCharacters) =
                     settingsFile.GetParam(XML_SECTION_OPTIONS, "NormalizeFileLineEndCharacters",
                     Me.OptionSwitch(SwitchOptions.NormalizeFileLineEndCharacters))


                    If Not settingsFile.SectionPresent(XML_SECTION_FIXED_FASTA_FILE_OPTIONS) Then
                        ' "ValidateFastaFixedFASTAFileOptions" section not present
                        ' Only read the settings for GenerateFixedFASTAFile and SplitOutMultipleRefsInProteinName

                        Me.OptionSwitch(SwitchOptions.GenerateFixedFASTAFile) =
                         settingsFile.GetParam(XML_SECTION_OPTIONS, "GenerateFixedFASTAFile",
                         Me.OptionSwitch(SwitchOptions.GenerateFixedFASTAFile))

                        Me.OptionSwitch(SwitchOptions.SplitOutMultipleRefsInProteinName) =
                         settingsFile.GetParam(XML_SECTION_OPTIONS, "SplitOutMultipleRefsInProteinName",
                         Me.OptionSwitch(SwitchOptions.SplitOutMultipleRefsInProteinName))

                    Else
                        ' "ValidateFastaFixedFASTAFileOptions" section is present

                        Me.OptionSwitch(SwitchOptions.GenerateFixedFASTAFile) =
                         settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "GenerateFixedFASTAFile",
                         Me.OptionSwitch(SwitchOptions.GenerateFixedFASTAFile))

                        Me.OptionSwitch(SwitchOptions.SplitOutMultipleRefsInProteinName) =
                         settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "SplitOutMultipleRefsInProteinName",
                         Me.OptionSwitch(SwitchOptions.SplitOutMultipleRefsInProteinName))

                        Me.OptionSwitch(SwitchOptions.FixedFastaRenameDuplicateNameProteins) =
                         settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "RenameDuplicateNameProteins",
                         Me.OptionSwitch(SwitchOptions.FixedFastaRenameDuplicateNameProteins))

                        Me.OptionSwitch(SwitchOptions.FixedFastaKeepDuplicateNamedProteins) =
                         settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "KeepDuplicateNamedProteins",
                         Me.OptionSwitch(SwitchOptions.FixedFastaKeepDuplicateNamedProteins))

                        Me.OptionSwitch(SwitchOptions.FixedFastaConsolidateDuplicateProteinSeqs) =
                         settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ConsolidateDuplicateProteinSeqs",
                         Me.OptionSwitch(SwitchOptions.FixedFastaConsolidateDuplicateProteinSeqs))

                        Me.OptionSwitch(SwitchOptions.FixedFastaConsolidateDupsIgnoreILDiff) =
                         settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ConsolidateDupsIgnoreILDiff",
                         Me.OptionSwitch(SwitchOptions.FixedFastaConsolidateDupsIgnoreILDiff))

                        Me.OptionSwitch(SwitchOptions.FixedFastaTruncateLongProteinNames) =
                         settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "TruncateLongProteinNames",
                         Me.OptionSwitch(SwitchOptions.FixedFastaTruncateLongProteinNames))

                        Me.OptionSwitch(SwitchOptions.FixedFastaSplitOutMultipleRefsForKnownAccession) =
                         settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "SplitOutMultipleRefsForKnownAccession",
                         Me.OptionSwitch(SwitchOptions.FixedFastaSplitOutMultipleRefsForKnownAccession))

                        Me.OptionSwitch(SwitchOptions.FixedFastaWrapLongResidueLines) =
                         settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "WrapLongResidueLines",
                         Me.OptionSwitch(SwitchOptions.FixedFastaWrapLongResidueLines))

                        Me.OptionSwitch(SwitchOptions.FixedFastaRemoveInvalidResidues) =
                         settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "RemoveInvalidResidues",
                         Me.OptionSwitch(SwitchOptions.FixedFastaRemoveInvalidResidues))

                        ' Look for the special character lists
                        ' If defined, then update the default values
                        characterList = settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "LongProteinNameSplitChars", String.Empty)
                        If Not characterList Is Nothing AndAlso characterList.Length > 0 Then
                            ' Update mFixedFastaOptions.LongProteinNameSplitChars with characterList
                            Me.LongProteinNameSplitChars = characterList
                        End If

                        characterList = settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ProteinNameInvalidCharsToRemove", String.Empty)
                        If Not characterList Is Nothing AndAlso characterList.Length > 0 Then
                            ' Update mFixedFastaOptions.ProteinNameInvalidCharsToRemove with characterList
                            Me.ProteinNameInvalidCharsToRemove = characterList
                        End If

                        characterList = settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ProteinNameFirstRefSepChars", String.Empty)
                        If Not characterList Is Nothing AndAlso characterList.Length > 0 Then
                            ' Update mProteinNameFirstRefSepChars
                            Me.ProteinNameFirstRefSepChars = characterList.ToCharArray
                        End If

                        characterList = settingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ProteinNameSubsequentRefSepChars", String.Empty)
                        If Not characterList Is Nothing AndAlso characterList.Length > 0 Then
                            ' Update mProteinNameSubsequentRefSepChars
                            Me.ProteinNameSubsequentRefSepChars = characterList.ToCharArray
                        End If
                    End If

                    ' Read the custom rules
                    ' If all of the sections are missing, then use the default rules
                    customRulesLoaded = False

                    success = ReadRulesFromParameterFile(settingsFile, XML_SECTION_FASTA_HEADER_LINE_RULES, mHeaderLineRules)
                    customRulesLoaded = customRulesLoaded Or success

                    success = ReadRulesFromParameterFile(settingsFile, XML_SECTION_FASTA_PROTEIN_NAME_RULES, mProteinNameRules)
                    customRulesLoaded = customRulesLoaded Or success

                    success = ReadRulesFromParameterFile(settingsFile, XML_SECTION_FASTA_PROTEIN_DESCRIPTION_RULES, mProteinDescriptionRules)
                    customRulesLoaded = customRulesLoaded Or success

                    success = ReadRulesFromParameterFile(settingsFile, XML_SECTION_FASTA_PROTEIN_SEQUENCE_RULES, mProteinSequenceRules)
                    customRulesLoaded = customRulesLoaded Or success
                End If
            End If
        Catch ex As Exception
            OnErrorEvent("Error in LoadParameterFileSettings", ex)
            Return False
        End Try

        If Not customRulesLoaded Then
            SetDefaultRules()
        End If

        Return True

    End Function

    Public Function LookupMessageDescription(errorMessageCode As Integer) As String
        Return Me.LookupMessageDescription(errorMessageCode, Nothing)
    End Function

    Public Function LookupMessageDescription(errorMessageCode As Integer, extraInfo As String) As String


        Dim message As String
        Dim matchFound As Boolean

        Select Case errorMessageCode
            ' Error messages
            Case eMessageCodeConstants.ProteinNameIsTooLong
                message = "Protein name is longer than the maximum allowed length of " & mMaximumProteinNameLength.ToString & " characters"
                'Case eMessageCodeConstants.ProteinNameContainsInvalidCharacters
                '    message = "Protein name contains invalid characters"
                '    If Not specifiedInvalidProteinNameChars Then
                '        message &= " (should contain letters, numbers, period, dash, underscore, colon, comma, or vertical bar)"
                '        specifiedInvalidProteinNameChars = True
                '    End If
            Case eMessageCodeConstants.LineStartsWithSpace
                message = "Found a line starting with a space"
                'Case eMessageCodeConstants.RightArrowFollowedBySpace
                '    message = "Space found directly after the > symbol"
                'Case eMessageCodeConstants.RightArrowFollowedByTab
                '    message = "Tab character found directly after the > symbol"
                'Case eMessageCodeConstants.RightArrowButNoProteinName
                '    message = "Line starts with > but does not contain a protein name"
            Case eMessageCodeConstants.BlankLineBetweenProteinNameAndResidues
                message = "A blank line was found between the protein name and its residues"
            Case eMessageCodeConstants.BlankLineInMiddleOfResidues
                message = "A blank line was found in the middle of the residue block for the protein"
            Case eMessageCodeConstants.ResiduesFoundWithoutProteinHeader
                message = "Residues were found, but a protein header didn't precede them"
                'Case eMessageCodeConstants.ResiduesWithAsterisk
                '    message = "An asterisk was found in the residues"
                'Case eMessageCodeConstants.ResiduesWithSpace
                '    message = "A space was found in the residues"
                'Case eMessageCodeConstants.ResiduesWithTab
                '    message = "A tab character was found in the residues"
                'Case eMessageCodeConstants.ResiduesWithInvalidCharacters
                '    message = "Invalid residues found"
                '    If Not specifiedResidueChars Then
                '        If mAllowAsteriskInResidues Then
                '            message &= " (should be any capital letter except J, plus *)"
                '        Else
                '            message &= " (should be any capital letter except J)"
                '        End If
                '        specifiedResidueChars = True
                '    End If
            Case eMessageCodeConstants.ProteinEntriesNotFound
                message = "File does not contain any protein entries"
            Case eMessageCodeConstants.FinalProteinEntryMissingResidues
                message = "The last entry in the file is a protein header line, but there is no protein sequence line after it"
            Case eMessageCodeConstants.FileDoesNotEndWithLinefeed
                message = "File does not end in a blank line; this is a problem for Sequest"
            Case eMessageCodeConstants.DuplicateProteinName
                message = "Duplicate protein name found"

                ' Warning messages
            Case eMessageCodeConstants.ProteinNameIsTooShort
                message = "Protein name is shorter than the minimum suggested length of " & mMinimumProteinNameLength.ToString & " characters"
                'Case eMessageCodeConstants.ProteinNameContainsComma
                '    message = "Protein name contains a comma"
                'Case eMessageCodeConstants.ProteinNameContainsVerticalBars
                '    message = "Protein name contains two or more vertical bars"
                'Case eMessageCodeConstants.ProteinNameContainsWarningCharacters
                '    message = "Protein name contains undesirable characters"
                'Case eMessageCodeConstants.ProteinNameWithoutDescription
                '    message = "Line contains a protein name, but not a description"
            Case eMessageCodeConstants.BlankLineBeforeProteinName
                message = "Blank line found before the protein name; this is acceptable, but not preferred"
                'Case eMessageCodeConstants.ProteinNameAndDescriptionSeparatedByTab
                '    message = "Protein name is separated from the protein description by a tab"
            Case eMessageCodeConstants.ResiduesLineTooLong
                message = "Residues line is longer than the suggested maximum length of " & mMaximumResiduesPerLine.ToString & " characters"
                'Case eMessageCodeConstants.ProteinDescriptionWithTab
                '    message = "Protein description contains a tab character"
                'Case eMessageCodeConstants.ProteinDescriptionWithQuotationMark
                '    message = "Protein description contains a quotation mark"
                'Case eMessageCodeConstants.ProteinDescriptionWithEscapedSlash
                '    message = "Protein description contains escaped slash: \/"
                'Case eMessageCodeConstants.ProteinDescriptionWithUndesirableCharacter
                '    message = "Protein description contains undesirable characters"
                'Case eMessageCodeConstants.ResiduesLineTooLong
                '    message = "Residues line is longer than the suggested maximum length of " & mMaximumResiduesPerLine.ToString & " characters"
                'Case eMessageCodeConstants.ResiduesLineContainsU
                '    message = "Residues line contains U (selenocysteine); this residue is unsupported by Sequest"

            Case eMessageCodeConstants.DuplicateProteinSequence
                message = "Duplicate protein sequences found"

            Case eMessageCodeConstants.RenamedProtein
                message = "Renamed protein because duplicate name"

            Case eMessageCodeConstants.ProteinRemovedSinceDuplicateSequence
                message = "Removed protein since duplicate sequence"

            Case eMessageCodeConstants.DuplicateProteinNameRetained
                message = "Duplicate protein retained in fixed file"

            Case eMessageCodeConstants.UnspecifiedError
                message = "Unspecified error"
            Case Else
                message = "Unspecified error"

                ' Search the custom rules for the given code
                matchFound = SearchRulesForID(mHeaderLineRules, errorMessageCode, message)

                If Not matchFound Then
                    matchFound = SearchRulesForID(mProteinNameRules, errorMessageCode, message)
                End If

                If Not matchFound Then
                    matchFound = SearchRulesForID(mProteinDescriptionRules, errorMessageCode, message)
                End If

                If Not matchFound Then
                    SearchRulesForID(mProteinSequenceRules, errorMessageCode, message)
                End If

        End Select

        If Not String.IsNullOrWhiteSpace(extraInfo) Then
            message &= " (" & extraInfo & ")"
        End If

        Return message

    End Function

    Private Function LookupMessageType(EntryType As eMsgTypeConstants) As String

        Select Case EntryType
            Case eMsgTypeConstants.ErrorMsg
                Return "Error"
            Case eMsgTypeConstants.WarningMsg
                Return "Warning"
            Case Else
                Return "Status"
        End Select
    End Function

    ''' <summary>
    ''' Validate a single fasta file
    ''' </summary>
    ''' <returns>True if success; false if a fatal error</returns>
    ''' <remarks>
    ''' Note that .ProcessFile returns True if a file is successfully processed (even if errors are found)
    ''' Used by clsCustomValidateFastaFiles
    ''' </remarks>
    Protected Function SimpleProcessFile(inputFilePath As String) As Boolean
        Return Me.ProcessFile(inputFilePath, Nothing, Nothing, False)
    End Function

    ''' <summary>
    ''' Main processing function
    ''' </summary>
    ''' <param name="inputFilePath"></param>
    ''' <param name="outputFolderPath"></param>
    ''' <param name="parameterFilePath"></param>
    ''' <param name="resetErrorCode"></param>
    ''' <returns>True if success, False if failure</returns>
    Public Overloads Overrides Function ProcessFile(
      inputFilePath As String,
      outputFolderPath As String,
      parameterFilePath As String,
      resetErrorCode As Boolean) As Boolean

        Dim ioFile As FileInfo
        Dim swStatsOutFile As StreamWriter

        Dim inputFilePathFull As String
        Dim statusMessage As String

        If resetErrorCode Then
            SetLocalErrorCode(eValidateFastaFileErrorCodes.NoError)
        End If

        If Not LoadParameterFileSettings(parameterFilePath) Then
            statusMessage = "Parameter file load error: " & parameterFilePath
            OnWarningEvent(statusMessage)
            If MyBase.ErrorCode = ProcessFilesErrorCodes.NoError Then
                MyBase.SetBaseClassErrorCode(ProcessFilesErrorCodes.InvalidParameterFile)
            End If
            Return False
        End If

        Try
            If inputFilePath Is Nothing OrElse inputFilePath.Length = 0 Then
                ShowWarning("Input file name is empty")
                MyBase.SetBaseClassErrorCode(ProcessFilesErrorCodes.InvalidInputFilePath)
                Return False
            Else

                Console.WriteLine()
                ShowMessage("Parsing " & Path.GetFileName(inputFilePath))

                If Not CleanupFilePaths(inputFilePath, outputFolderPath) Then
                    MyBase.SetBaseClassErrorCode(ProcessFilesErrorCodes.FilePathError)
                    Return False
                Else
                    ' List of protein names to keep
                    ' Keys are protein names, values are the number of entries written to the fixed fasta file for the given protein name
                    Dim preloadedProteinNamesToKeep As clsNestedStringIntList = Nothing

                    If Not String.IsNullOrEmpty(ExistingProteinHashFile) Then
                        Dim loadSuccess = LoadExistingProteinHashFile(ExistingProteinHashFile, preloadedProteinNamesToKeep)
                        If Not loadSuccess Then
                            Return False
                        End If
                    End If

                    Try

                        ' Obtain the full path to the input file
                        ioFile = New FileInfo(inputFilePath)
                        inputFilePathFull = ioFile.FullName

                        Dim success = AnalyzeFastaFile(inputFilePathFull, preloadedProteinNamesToKeep)

                        If success Then
                            ReportResults(outputFolderPath, mOutputToStatsFile)
                            DeleteTempFiles()
                            Return True
                        Else
                            If mOutputToStatsFile Then
                                mStatsFilePath = ConstructStatsFilePath(outputFolderPath)
                                swStatsOutFile = New StreamWriter(mStatsFilePath, True)
                                swStatsOutFile.WriteLine(GetTimeStamp() & ControlChars.Tab &
                                                         "Error parsing " &
                                                         Path.GetFileName(inputFilePath) & ": " & Me.GetErrorMessage())
                                swStatsOutFile.Close()
                            Else
                                ShowMessage("Error parsing " &
                                  Path.GetFileName(inputFilePath) &
                                  ": " & Me.GetErrorMessage())
                            End If
                            Return False
                        End If

                    Catch ex As Exception
                        OnErrorEvent("Error calling AnalyzeFastaFile", ex)
                        Return False
                    End Try
                End If
            End If
        Catch ex As Exception
            OnErrorEvent("Error in ProcessFile", ex)
            Return False
        End Try

    End Function

    Private Sub PrependExtraTextToProteinDescription(extraProteinNameText As String, ByRef proteinDescription As String)
        Static extraCharsToTrim As Char() = New Char() {"|"c, " "c}

        If Not extraProteinNameText Is Nothing AndAlso extraProteinNameText.Length > 0 Then
            ' If extraProteinNameText ends in a vertical bar and/or space, them remove them
            extraProteinNameText = extraProteinNameText.TrimEnd(extraCharsToTrim)

            If Not proteinDescription Is Nothing AndAlso proteinDescription.Length > 0 Then
                If proteinDescription.Chars(0) = " "c OrElse proteinDescription.Chars(0) = "|"c Then
                    proteinDescription = extraProteinNameText & proteinDescription
                Else
                    proteinDescription = extraProteinNameText & " " & proteinDescription
                End If
            Else
                proteinDescription = String.Copy(extraProteinNameText)
            End If
        End If


    End Sub

    Private Sub ProcessResiduesForPreviousProtein(
      proteinName As String,
      sbCurrentResidues As StringBuilder,
      proteinSequenceHashes As clsNestedStringDictionary(Of Integer),
      ByRef proteinSequenceHashCount As Integer,
      ByRef proteinSeqHashInfo() As clsProteinHashInfo,
      consolidateDupsIgnoreILDiff As Boolean,
      fixedFastaWriter As TextWriter,
      currentValidResidueLineLengthMax As Integer,
      sequenceHashWriter As TextWriter)

        Dim wrapLength As Integer

        Dim index As Integer
        Dim length As Integer

        ' Check for and remove any asterisks at the end of the residues
        While sbCurrentResidues.Length > 0 AndAlso sbCurrentResidues.Chars(sbCurrentResidues.Length - 1) = "*"
            sbCurrentResidues.Remove(sbCurrentResidues.Length - 1, 1)
        End While

        If sbCurrentResidues.Length > 0 Then

            ' Remove any spaces from the residues

            If mCheckForDuplicateProteinSequences OrElse mSaveBasicProteinHashInfoFile Then
                ' Process the previous protein entry to store a hash of the protein sequence
                ProcessSequenceHashInfo(
                    proteinName, sbCurrentResidues,
                    proteinSequenceHashes,
                    proteinSequenceHashCount, proteinSeqHashInfo,
                    consolidateDupsIgnoreILDiff, sequenceHashWriter)
            End If

            If mGenerateFixedFastaFile AndAlso mFixedFastaOptions.WrapLongResidueLines Then
                ' Write out the residues
                ' Wrap the lines at currentValidResidueLineLengthMax characters (but do not allow to be longer than mMaximumResiduesPerLine residues)

                wrapLength = currentValidResidueLineLengthMax
                If wrapLength <= 0 OrElse wrapLength > mMaximumResiduesPerLine Then
                    wrapLength = mMaximumResiduesPerLine
                End If

                If wrapLength < 10 Then
                    ' Do not allow wrapLength to be less than 10
                    wrapLength = 10
                End If

                index = 0
                Dim proteinResidueCount = sbCurrentResidues.Length
                Do While index < sbCurrentResidues.Length
                    length = Math.Min(wrapLength, proteinResidueCount - index)
                    fixedFastaWriter.WriteLine(sbCurrentResidues.ToString(index, length))
                    index += wrapLength
                Loop

            End If

            sbCurrentResidues.Length = 0

        End If

    End Sub

    Private Sub ProcessSequenceHashInfo(
      proteinName As String,
      sbCurrentResidues As StringBuilder,
      proteinSequenceHashes As clsNestedStringDictionary(Of Integer),
      ByRef proteinSequenceHashCount As Integer,
      ByRef proteinSeqHashInfo() As clsProteinHashInfo,
      consolidateDupsIgnoreILDiff As Boolean,
      sequenceHashWriter As TextWriter)

        Dim computedHash As String

        Try
            If sbCurrentResidues.Length > 0 Then

                ' Compute the hash value for sbCurrentResidues
                computedHash = ComputeProteinHash(sbCurrentResidues, consolidateDupsIgnoreILDiff)

                If Not sequenceHashWriter Is Nothing Then

                    Dim dataValues = New List(Of String) From {
                        mProteinCount.ToString,
                        proteinName,
                        sbCurrentResidues.Length.ToString,
                        computedHash}

                    sequenceHashWriter.WriteLine(FlattenList(dataValues))
                End If

                If mCheckForDuplicateProteinSequences AndAlso Not proteinSequenceHashes Is Nothing Then

                    ' See if proteinSequenceHashes contains hash
                    Dim seqHashLookupPointer As Integer
                    If proteinSequenceHashes.TryGetValue(computedHash, seqHashLookupPointer) Then

                        ' Value exists; update the entry in proteinSeqHashInfo
                        CachedSequenceHashInfoUpdate(proteinSeqHashInfo(seqHashLookupPointer), proteinName)

                    Else
                        ' Value not yet present; add it
                        CachedSequenceHashInfoUpdateAppend(
                            proteinSequenceHashCount, proteinSeqHashInfo,
                            computedHash, sbCurrentResidues, proteinName)

                        proteinSequenceHashes.Add(computedHash, proteinSequenceHashCount)
                        proteinSequenceHashCount += 1

                    End If

                End If
            End If

        Catch ex As Exception
            ' Error caught; pass it up to the calling function
            ShowMessage(ex.Message)
            Throw
        End Try


    End Sub

    Private Sub CachedSequenceHashInfoUpdate(proteinSeqHashInfo As clsProteinHashInfo, proteinName As String)

        If proteinSeqHashInfo.ProteinNameFirst = proteinName Then
            proteinSeqHashInfo.DuplicateProteinNameCount += 1
        Else
            proteinSeqHashInfo.AddAdditionalProtein(proteinName)
        End If
    End Sub

    Private Sub CachedSequenceHashInfoUpdateAppend(
      ByRef proteinSequenceHashCount As Integer,
      ByRef proteinSeqHashInfo() As clsProteinHashInfo,
      computedHash As String,
      sbCurrentResidues As StringBuilder,
      proteinName As String)

        If proteinSequenceHashCount >= proteinSeqHashInfo.Length Then
            ' Need to reserve more space in proteinSeqHashInfo
            If proteinSeqHashInfo.Length < 1000000 Then
                ReDim Preserve proteinSeqHashInfo(proteinSeqHashInfo.Length * 2 - 1)
            Else
                ReDim Preserve proteinSeqHashInfo(CInt(proteinSeqHashInfo.Length * 1.2) - 1)
            End If

        End If

        Dim newProteinHashInfo = New clsProteinHashInfo(computedHash, sbCurrentResidues, proteinName)
        proteinSeqHashInfo(proteinSequenceHashCount) = newProteinHashInfo

    End Sub

    Private Function ReadRulesFromParameterFile(
      settingsFile As XmlSettingsFileAccessor,
      sectionName As String,
      ByRef rules() As udtRuleDefinitionType) As Boolean
        ' Returns True if the section named sectionName is present and if it contains an item with keyName = "RuleCount"
        ' Note: even if RuleCount = 0, this function will return True

        Dim success = False
        Dim ruleCount As Integer
        Dim ruleNumber As Integer

        Dim ruleBase As String

        Dim newRule As udtRuleDefinitionType

        ruleCount = settingsFile.GetParam(sectionName, XML_OPTION_ENTRY_RULE_COUNT, -1)

        If ruleCount >= 0 Then
            ClearRulesDataStructure(rules)

            For ruleNumber = 1 To ruleCount
                ruleBase = "Rule" & ruleNumber.ToString

                newRule.MatchRegEx = settingsFile.GetParam(sectionName, ruleBase & "MatchRegEx", String.Empty)

                If newRule.MatchRegEx.Length > 0 Then
                    ' Only read the rule settings if MatchRegEx contains 1 or more characters

                    With newRule
                        .MatchIndicatesProblem = settingsFile.GetParam(sectionName, ruleBase & "MatchIndicatesProblem", True)
                        .MessageWhenProblem = settingsFile.GetParam(sectionName, ruleBase & "MessageWhenProblem", "Error found with RegEx " & .MatchRegEx)
                        .Severity = settingsFile.GetParam(sectionName, ruleBase & "Severity", 3S)
                        .DisplayMatchAsExtraInfo = settingsFile.GetParam(sectionName, ruleBase & "DisplayMatchAsExtraInfo", False)

                        SetRule(rules, .MatchRegEx, .MatchIndicatesProblem, .MessageWhenProblem, .Severity, .DisplayMatchAsExtraInfo)

                    End With

                End If

            Next ruleNumber

            success = True
        End If

        Return success

    End Function

    Private Sub RecordFastaFileError(
      lineNumber As Integer,
      charIndex As Integer,
      proteinName As String,
      errorMessageCode As Integer)

        RecordFastaFileError(lineNumber, charIndex, proteinName,
         errorMessageCode, String.Empty, String.Empty)
    End Sub

    Private Sub RecordFastaFileError(
      lineNumber As Integer,
      charIndex As Integer,
      proteinName As String,
      errorMessageCode As Integer,
      extraInfo As String,
      context As String)
        RecordFastaFileProblemWork(
         mFileErrorStats,
         mFileErrorCount,
         mFileErrors,
         lineNumber,
         charIndex,
         proteinName,
         errorMessageCode,
         extraInfo, context)
    End Sub

    Private Sub RecordFastaFileWarning(
      lineNumber As Integer,
      charIndex As Integer,
      proteinName As String,
      warningMessageCode As Integer)
        RecordFastaFileWarning(
         lineNumber,
         charIndex,
         proteinName,
         warningMessageCode,
         String.Empty, String.Empty)
    End Sub

    Private Sub RecordFastaFileWarning(
      lineNumber As Integer,
      charIndex As Integer,
      proteinName As String,
      warningMessageCode As Integer,
      extraInfo As String, context As String)

        RecordFastaFileProblemWork(mFileWarningStats, mFileWarningCount,
         mFileWarnings, lineNumber, charIndex, proteinName,
         warningMessageCode, extraInfo, context)
    End Sub

    Private Sub RecordFastaFileProblemWork(
      ByRef itemSummaryIndexed As udtItemSummaryIndexedType,
      ByRef itemCountSpecified As Integer,
      ByRef items() As udtMsgInfoType,
      lineNumber As Integer,
      charIndex As Integer,
      proteinName As String,
      messageCode As Integer,
      extraInfo As String,
      context As String)

        ' Note that charIndex is the index in the source string at which the error occurred
        ' When storing in .ColNumber, we add 1 to charIndex

        ' Lookup the index of the entry with messageCode in itemSummaryIndexed.ErrorStats
        ' Add it if not present

        Try
            Dim itemIndex As Integer

            With itemSummaryIndexed
                If Not .MessageCodeToArrayIndex.TryGetValue(messageCode, itemIndex) Then
                    If .ErrorStats.Length <= 0 Then
                        ReDim .ErrorStats(1)
                    ElseIf .ErrorStatsCount = .ErrorStats.Length Then
                        ReDim Preserve .ErrorStats(.ErrorStats.Length * 2 - 1)
                    End If
                    itemIndex = .ErrorStatsCount
                    .ErrorStats(itemIndex).MessageCode = messageCode
                    .MessageCodeToArrayIndex.Add(messageCode, itemIndex)
                    .ErrorStatsCount += 1
                End If

            End With

            With itemSummaryIndexed.ErrorStats(itemIndex)
                If .CountSpecified >= mMaximumFileErrorsToTrack Then
                    .CountUnspecified += 1
                Else
                    If items.Length <= 0 Then
                        ' Initially reserve space for 10 errors
                        ReDim items(10)
                    ElseIf itemCountSpecified >= items.Length Then
                        ' Double the amount of space reserved for errors
                        ReDim Preserve items(items.Length * 2 - 1)
                    End If

                    With items(itemCountSpecified)
                        .LineNumber = lineNumber
                        .ColNumber = charIndex + 1
                        If proteinName Is Nothing Then
                            .ProteinName = String.Empty
                        Else
                            .ProteinName = proteinName
                        End If

                        .MessageCode = messageCode
                        If extraInfo Is Nothing Then
                            .ExtraInfo = String.Empty
                        Else
                            .ExtraInfo = extraInfo
                        End If

                        If extraInfo Is Nothing Then
                            .Context = String.Empty
                        Else
                            .Context = context
                        End If

                    End With
                    itemCountSpecified += 1

                    .CountSpecified += 1
                End If

            End With

        Catch ex As Exception
            ' Ignore any errors that occur, but output the error to the console
            OnWarningEvent("Error in RecordFastaFileProblemWork: " & ex.Message)
        End Try

    End Sub

    Private Sub ReplaceXMLCodesWithText(parameterFilePath As String)

        Dim outputFilePath As String
        Dim timeStamp As String
        Dim lineIn As String

        Try
            ' Define the output file path
            timeStamp = GetTimeStamp().Replace(" ", "_").Replace(":", "_").Replace("/", "_")

            outputFilePath = parameterFilePath & "_" & timeStamp & ".fixed"

            ' Open the input file and output file
            Using srInFile = New StreamReader(parameterFilePath),
                swOutFile = New StreamWriter(New FileStream(outputFilePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))

                ' Parse each line in the file
                Do While Not srInFile.EndOfStream
                    lineIn = srInFile.ReadLine()

                    If Not lineIn Is Nothing Then
                        lineIn = lineIn.Replace("&gt;", ">").Replace("&lt;", "<")
                        swOutFile.WriteLine(lineIn)
                    End If
                Loop

                ' Close the input and output files
            End Using

            ' Wait 100 msec
            Threading.Thread.Sleep(100)

            ' Delete the input file
            File.Delete(parameterFilePath)

            ' Wait 250 msec
            Threading.Thread.Sleep(250)

            ' Rename the output file to the input file
            File.Move(outputFilePath, parameterFilePath)

        Catch ex As Exception
            OnErrorEvent("Error in ReplaceXMLCodesWithText", ex)
        End Try

    End Sub

    Private Sub ReportMemoryUsage()

        ' ReSharper disable once VbUnreachableCode
        If REPORT_DETAILED_MEMORY_USAGE Then
            Console.WriteLine(MEM_USAGE_PREFIX & mMemoryUsageLogger.GetMemoryUsageSummary())
        Else
            Console.WriteLine(MEM_USAGE_PREFIX & GetProcessMemoryUsageWithTimestamp())
        End If

    End Sub

    Private Sub ReportMemoryUsage(
      preloadedProteinNamesToKeep As clsNestedStringIntList,
      proteinSequenceHashes As clsNestedStringDictionary(Of Integer),
      proteinNames As ICollection,
      proteinSeqHashInfo As IEnumerable(Of clsProteinHashInfo))

        Console.WriteLine()
        ReportMemoryUsage()

        If Not preloadedProteinNamesToKeep Is Nothing AndAlso preloadedProteinNamesToKeep.Count > 0 Then
            Console.WriteLine(" PreloadedProteinNamesToKeep: {0,12:#,##0} records", preloadedProteinNamesToKeep.Count)
        End If

        If proteinSequenceHashes.Count > 0 Then
            Console.WriteLine(" ProteinSequenceHashes:  {0,12:#,##0} records", proteinSequenceHashes.Count)
            Console.WriteLine("   {0}", proteinSequenceHashes.GetSizeSummary())
        End If

        Console.WriteLine(" ProteinNames:           {0,12:#,##0} records", proteinNames.Count)
        Console.WriteLine(" ProteinSeqHashInfo:       {0,12:#,##0} records", proteinSeqHashInfo.Count)

    End Sub

    Private Sub ReportMemoryUsage(
      proteinNameFirst As clsNestedStringDictionary(Of Integer),
      proteinsWritten As clsNestedStringDictionary(Of Integer),
      duplicateProteinList As clsNestedStringDictionary(Of String))

        Console.WriteLine()
        ReportMemoryUsage()
        Console.WriteLine(" ProteinNameFirst:      {0,12:#,##0} records", proteinNameFirst.Count)
        Console.WriteLine("   {0}", proteinNameFirst.GetSizeSummary())
        Console.WriteLine(" ProteinsWritten:       {0,12:#,##0} records", proteinsWritten.Count)
        Console.WriteLine("   {0}", proteinsWritten.GetSizeSummary())
        Console.WriteLine(" DuplicateProteinList:  {0,12:#,##0} records", duplicateProteinList.Count)
        Console.WriteLine("   {0}", duplicateProteinList.GetSizeSummary())

    End Sub

    Private Sub ReportResults(
      outputFolderPath As String,
      outputToStatsFile As Boolean)

        Dim iErrorInfoComparerClass As ErrorInfoComparerClass

        Dim proteinName As String

        Dim index As Integer
        Dim retryCount As Integer

        Dim success As Boolean
        Dim fileAlreadyExists As Boolean

        Try
            Dim outputOptions = New udtOutputOptionsType With {
                .OutputToStatsFile = outputToStatsFile,
                .SepChar = ControlChars.Tab
            }

            Try
                outputOptions.SourceFile = Path.GetFileName(mFastaFilePath)
            Catch ex As Exception
                outputOptions.SourceFile = "Unknown_filename_due_to_error.fasta"
            End Try

            If outputToStatsFile Then
                mStatsFilePath = ConstructStatsFilePath(outputFolderPath)
                fileAlreadyExists = File.Exists(mStatsFilePath)

                success = False
                retryCount = 0

                Do While Not success And retryCount < 5
                    Dim outStream As FileStream = Nothing
                    Try
                        outStream = New FileStream(mStatsFilePath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite)
                        Dim outFileWriter = New StreamWriter(outStream)

                        outputOptions.OutFile = outFileWriter

                        success = True
                    Catch ex As Exception
                        ' Failed to open file, wait 1 second, then try again
                        If Not outStream Is Nothing Then
                            outStream.Close()
                        End If

                        retryCount += 1
                        Threading.Thread.Sleep(1000)
                    End Try
                Loop

                If success Then
                    outputOptions.SepChar = ControlChars.Tab
                    If Not fileAlreadyExists Then
                        ' Write the header line
                        Dim headers = New List(Of String) From {
                                "Date",
                                "SourceFile",
                                "MessageType",
                                "LineNumber",
                                "ColumnNumber",
                                "Description_or_Protein",
                                "Info",
                                "Context"
                        }

                        outputOptions.OutFile.WriteLine(String.Join(outputOptions.SepChar, headers))

                    End If
                Else
                    outputOptions.SepChar = ", "
                    outputToStatsFile = False
                    SetLocalErrorCode(eValidateFastaFileErrorCodes.ErrorCreatingStatsFile)
                End If
            Else
                outputOptions.SepChar = ", "
            End If

            ReportResultAddEntry(
                outputOptions, eMsgTypeConstants.StatusMsg,
                "Full path to file", mFastaFilePath)

            ReportResultAddEntry(
                outputOptions, eMsgTypeConstants.StatusMsg,
                "Protein count", mProteinCount.ToString("#,##0"))

            ReportResultAddEntry(
                outputOptions, eMsgTypeConstants.StatusMsg,
                "Residue count", mResidueCount.ToString("#,##0"))

            If mFileErrorCount > 0 Then
                ReportResultAddEntry(
                    outputOptions, eMsgTypeConstants.ErrorMsg,
                    "Error count", Me.ErrorWarningCounts(eMsgTypeConstants.ErrorMsg, ErrorWarningCountTypes.Total).ToString)

                If mFileErrorCount > 1 Then
                    iErrorInfoComparerClass = New ErrorInfoComparerClass
                    Array.Sort(mFileErrors, 0, mFileErrorCount, iErrorInfoComparerClass)
                End If

                For index = 0 To mFileErrorCount - 1
                    With mFileErrors(index)
                        If .ProteinName Is Nothing OrElse .ProteinName.Length = 0 Then
                            proteinName = "N/A"
                        Else
                            proteinName = String.Copy(.ProteinName)
                        End If

                        Dim messageDescription = LookupMessageDescription(.MessageCode, .ExtraInfo)

                        ReportResultAddEntry(
                            outputOptions, eMsgTypeConstants.ErrorMsg,
                            .LineNumber,
                            .ColNumber,
                            proteinName,
                            messageDescription,
                           .Context)

                    End With
                Next index
            End If

            If mFileWarningCount > 0 Then
                ReportResultAddEntry(
                    outputOptions, eMsgTypeConstants.WarningMsg,
                    "Warning count",
                    Me.ErrorWarningCounts(eMsgTypeConstants.WarningMsg, ErrorWarningCountTypes.Total).ToString)

                If mFileWarningCount > 1 Then
                    iErrorInfoComparerClass = New ErrorInfoComparerClass
                    Array.Sort(mFileWarnings, 0, mFileWarningCount, iErrorInfoComparerClass)
                End If

                For index = 0 To mFileWarningCount - 1
                    With mFileWarnings(index)
                        If .ProteinName Is Nothing OrElse .ProteinName.Length = 0 Then
                            proteinName = "N/A"
                        Else
                            proteinName = String.Copy(.ProteinName)
                        End If

                        ReportResultAddEntry(
                            outputOptions, eMsgTypeConstants.WarningMsg,
                            .LineNumber,
                            .ColNumber,
                            proteinName,
                            LookupMessageDescription(.MessageCode, .ExtraInfo),
                            .Context)

                    End With
                Next index
            End If

            Dim fastaFile = New FileInfo(mFastaFilePath)

            ' # Proteins, # Peptides, FileSizeKB
            ReportResultAddEntry(
                outputOptions, eMsgTypeConstants.StatusMsg,
                "Summary line",
                 mProteinCount.ToString() & " proteins, " & mResidueCount.ToString() & " residues, " & (fastaFile.Length / 1024.0).ToString("0") & " KB")

            If outputToStatsFile AndAlso Not outputOptions.OutFile Is Nothing Then
                outputOptions.OutFile.Close()
            End If

        Catch ex As Exception
            OnErrorEvent("Error in ReportResults", ex)
        End Try

    End Sub

    Private Sub ReportResultAddEntry(
      outputOptions As udtOutputOptionsType,
      entryType As eMsgTypeConstants,
      descriptionOrProteinName As String,
      info As String,
      Optional context As String = "")

        ReportResultAddEntry(
            outputOptions,
            entryType, 0, 0,
            descriptionOrProteinName,
            info,
            context)
    End Sub

    Private Sub ReportResultAddEntry(
      outputOptions As udtOutputOptionsType,
      entryType As eMsgTypeConstants,
      lineNumber As Integer,
      colNumber As Integer,
      descriptionOrProteinName As String,
      info As String,
      context As String)

        Dim dataColumns = New List(Of String) From {
            outputOptions.SourceFile,
            LookupMessageType(entryType),
            lineNumber.ToString,
            colNumber.ToString,
            descriptionOrProteinName,
            info
        }

        If Not context Is Nothing AndAlso context.Length > 0 Then
            dataColumns.Add(context)
        End If

        Dim message = String.Join(outputOptions.SepChar, dataColumns)

        If outputOptions.OutputToStatsFile Then
            outputOptions.OutFile.WriteLine(GetTimeStamp() & outputOptions.SepChar & message)
        Else
            Console.WriteLine(message)
        End If

    End Sub

    Private Sub ResetStructures()
        ' This is used to reset the error arrays and stats variables

        mLineCount = 0
        mProteinCount = 0
        mResidueCount = 0

        With mFixedFastaStats
            .TruncatedProteinNameCount = 0
            .UpdatedResidueLines = 0
            .ProteinNamesInvalidCharsReplaced = 0
            .ProteinNamesMultipleRefsRemoved = 0
            .DuplicateNameProteinsSkipped = 0
            .DuplicateNameProteinsRenamed = 0
            .DuplicateSequenceProteinsSkipped = 0
        End With

        mFileErrorCount = 0
        ReDim mFileErrors(-1)
        ResetItemSummaryStructure(mFileErrorStats)

        mFileWarningCount = 0
        ReDim mFileWarnings(-1)
        ResetItemSummaryStructure(mFileWarningStats)

        MyBase.AbortProcessing = False
    End Sub

    Private Sub ResetItemSummaryStructure(ByRef itemSummary As udtItemSummaryIndexedType)
        With itemSummary
            .ErrorStatsCount = 0
            ReDim .ErrorStats(-1)
            If .MessageCodeToArrayIndex Is Nothing Then
                .MessageCodeToArrayIndex = New Dictionary(Of Integer, Integer)
            Else
                .MessageCodeToArrayIndex.Clear()
            End If
        End With

    End Sub

    Private Sub SaveRulesToParameterFile(settingsFile As XmlSettingsFileAccessor, sectionName As String, rules As IList(Of udtRuleDefinitionType))

        Dim ruleNumber As Integer
        Dim ruleBase As String

        If rules Is Nothing OrElse rules.Count <= 0 Then
            settingsFile.SetParam(sectionName, XML_OPTION_ENTRY_RULE_COUNT, 0)
        Else
            settingsFile.SetParam(sectionName, XML_OPTION_ENTRY_RULE_COUNT, rules.Count)

            For ruleNumber = 1 To rules.Count
                ruleBase = "Rule" & ruleNumber.ToString

                With rules(ruleNumber - 1)
                    settingsFile.SetParam(sectionName, ruleBase & "MatchRegEx", .MatchRegEx)
                    settingsFile.SetParam(sectionName, ruleBase & "MatchIndicatesProblem", .MatchIndicatesProblem)
                    settingsFile.SetParam(sectionName, ruleBase & "MessageWhenProblem", .MessageWhenProblem)
                    settingsFile.SetParam(sectionName, ruleBase & "Severity", .Severity)
                    settingsFile.SetParam(sectionName, ruleBase & "DisplayMatchAsExtraInfo", .DisplayMatchAsExtraInfo)
                End With

            Next ruleNumber
        End If

    End Sub

    Public Function SaveSettingsToParameterFile(parameterFilePath As String) As Boolean
        ' Save a model parameter file

        Dim srOutFile As StreamWriter
        Dim settingsFile As New XmlSettingsFileAccessor

        Try

            If parameterFilePath Is Nothing OrElse parameterFilePath.Length = 0 Then
                ' No parameter file specified; do not save the settings
                Return True
            End If

            If Not File.Exists(parameterFilePath) Then
                ' Need to generate a blank XML settings file

                srOutFile = New StreamWriter(parameterFilePath, False)

                srOutFile.WriteLine("<?xml version=""1.0"" encoding=""UTF-8""?>")
                srOutFile.WriteLine("<sections>")
                srOutFile.WriteLine("  <section name=""" & XML_SECTION_OPTIONS & """>")
                srOutFile.WriteLine("  </section>")
                srOutFile.WriteLine("</sections>")

                srOutFile.Close()
            End If

            settingsFile.LoadSettings(parameterFilePath)

            ' Save the general settings

            settingsFile.SetParam(XML_SECTION_OPTIONS, "AddMissingLinefeedAtEOF", Me.OptionSwitch(SwitchOptions.AddMissingLinefeedatEOF))
            settingsFile.SetParam(XML_SECTION_OPTIONS, "AllowAsteriskInResidues", Me.OptionSwitch(SwitchOptions.AllowAsteriskInResidues))
            settingsFile.SetParam(XML_SECTION_OPTIONS, "AllowDashInResidues", Me.OptionSwitch(SwitchOptions.AllowDashInResidues))

            settingsFile.SetParam(XML_SECTION_OPTIONS, "CheckForDuplicateProteinNames", Me.OptionSwitch(SwitchOptions.CheckForDuplicateProteinNames))
            settingsFile.SetParam(XML_SECTION_OPTIONS, "CheckForDuplicateProteinSequences", Me.OptionSwitch(SwitchOptions.CheckForDuplicateProteinSequences))
            settingsFile.SetParam(XML_SECTION_OPTIONS, "SaveProteinSequenceHashInfoFiles", Me.OptionSwitch(SwitchOptions.SaveProteinSequenceHashInfoFiles))
            settingsFile.SetParam(XML_SECTION_OPTIONS, "SaveBasicProteinHashInfoFile", Me.OptionSwitch(SwitchOptions.SaveBasicProteinHashInfoFile))

            settingsFile.SetParam(XML_SECTION_OPTIONS, "MaximumFileErrorsToTrack", Me.MaximumFileErrorsToTrack)
            settingsFile.SetParam(XML_SECTION_OPTIONS, "MinimumProteinNameLength", Me.MinimumProteinNameLength)
            settingsFile.SetParam(XML_SECTION_OPTIONS, "MaximumProteinNameLength", Me.MaximumProteinNameLength)
            settingsFile.SetParam(XML_SECTION_OPTIONS, "MaximumResiduesPerLine", Me.MaximumResiduesPerLine)

            settingsFile.SetParam(XML_SECTION_OPTIONS, "WarnBlankLinesBetweenProteins", Me.OptionSwitch(SwitchOptions.WarnBlankLinesBetweenProteins))
            settingsFile.SetParam(XML_SECTION_OPTIONS, "WarnLineStartsWithSpace", Me.OptionSwitch(SwitchOptions.WarnLineStartsWithSpace))
            settingsFile.SetParam(XML_SECTION_OPTIONS, "OutputToStatsFile", Me.OptionSwitch(SwitchOptions.OutputToStatsFile))
            settingsFile.SetParam(XML_SECTION_OPTIONS, "NormalizeFileLineEndCharacters", Me.OptionSwitch(SwitchOptions.NormalizeFileLineEndCharacters))


            settingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "GenerateFixedFASTAFile", Me.OptionSwitch(SwitchOptions.GenerateFixedFASTAFile))
            settingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "SplitOutMultipleRefsInProteinName", Me.OptionSwitch(SwitchOptions.SplitOutMultipleRefsInProteinName))

            settingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "RenameDuplicateNameProteins", Me.OptionSwitch(SwitchOptions.FixedFastaRenameDuplicateNameProteins))
            settingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "KeepDuplicateNamedProteins", Me.OptionSwitch(SwitchOptions.FixedFastaKeepDuplicateNamedProteins))

            settingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ConsolidateDuplicateProteinSeqs", Me.OptionSwitch(SwitchOptions.FixedFastaConsolidateDuplicateProteinSeqs))
            settingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ConsolidateDupsIgnoreILDiff", Me.OptionSwitch(SwitchOptions.FixedFastaConsolidateDupsIgnoreILDiff))

            settingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "TruncateLongProteinNames", Me.OptionSwitch(SwitchOptions.FixedFastaTruncateLongProteinNames))
            settingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "SplitOutMultipleRefsForKnownAccession", Me.OptionSwitch(SwitchOptions.FixedFastaSplitOutMultipleRefsForKnownAccession))
            settingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "WrapLongResidueLines", Me.OptionSwitch(SwitchOptions.FixedFastaWrapLongResidueLines))
            settingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "RemoveInvalidResidues", Me.OptionSwitch(SwitchOptions.FixedFastaRemoveInvalidResidues))


            settingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "LongProteinNameSplitChars", Me.LongProteinNameSplitChars)
            settingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ProteinNameInvalidCharsToRemove", Me.ProteinNameInvalidCharsToRemove)
            settingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ProteinNameFirstRefSepChars", Me.ProteinNameFirstRefSepChars)
            settingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ProteinNameSubsequentRefSepChars", Me.ProteinNameSubsequentRefSepChars)


            ' Save the rules
            SaveRulesToParameterFile(settingsFile, XML_SECTION_FASTA_HEADER_LINE_RULES, mHeaderLineRules)
            SaveRulesToParameterFile(settingsFile, XML_SECTION_FASTA_PROTEIN_NAME_RULES, mProteinNameRules)
            SaveRulesToParameterFile(settingsFile, XML_SECTION_FASTA_PROTEIN_DESCRIPTION_RULES, mProteinDescriptionRules)
            SaveRulesToParameterFile(settingsFile, XML_SECTION_FASTA_PROTEIN_SEQUENCE_RULES, mProteinSequenceRules)

            ' Commit the new settings to disk
            settingsFile.SaveSettings()

            ' Need to re-open the parameter file and replace instances of "&gt;" with ">" and "&lt;" with "<"
            ReplaceXMLCodesWithText(parameterFilePath)

        Catch ex As Exception
            OnErrorEvent("Error in SaveSettingsToParameterFile", ex)
            Return False
        End Try

        Return True

    End Function

    Private Function SearchRulesForID(
      rules As IList(Of udtRuleDefinitionType),
      errorMessageCode As Integer,
      <Out> ByRef message As String) As Boolean

        If Not rules Is Nothing Then
            For index = 0 To rules.Count - 1
                If rules(index).CustomRuleID = errorMessageCode Then
                    message = rules(index).MessageWhenProblem
                    Return True
                End If
            Next index
        End If

        message = Nothing
        Return False

    End Function

    ''' <summary>
    ''' Updates the validation rules using the current options
    ''' </summary>
    ''' <remarks>Call this function after setting new options using SetOptionSwitch</remarks>
    Public Sub SetDefaultRules()

        Me.ClearAllRules()

        ' For the rules, severity level 1 to 4 is warning; severity 5 or higher is an error

        ' Header line errors
        Me.SetRule(RuleTypes.HeaderLine, "^>[ \t]*$", True, "Line starts with > but does not contain a protein name", DEFAULT_ERROR_SEVERITY)
        Me.SetRule(RuleTypes.HeaderLine, "^>[ \t].+", True, "Space or tab found directly after the > symbol", DEFAULT_ERROR_SEVERITY)

        ' Header line warnings
        Me.SetRule(RuleTypes.HeaderLine, "^>[^ \t]+[ \t]*$", True, MESSAGE_TEXT_PROTEIN_DESCRIPTION_MISSING, DEFAULT_WARNING_SEVERITY)
        Me.SetRule(RuleTypes.HeaderLine, "^>[^ \t]+\t", True, "Protein name is separated from the protein description by a tab", DEFAULT_WARNING_SEVERITY)

        ' Protein Name error characters
        Dim allowedChars = "A-Za-z0-9.\-_:,\|/()\[\]\=\+#"

        If mAllowAllSymbolsInProteinNames Then
            allowedChars &= "!@$%^&*<>?,\\"
        End If

        Dim allowedCharsMatchSpec = "[^" & allowedChars & "]"

        Me.SetRule(RuleTypes.ProteinName, allowedCharsMatchSpec, True, "Protein name contains invalid characters", DEFAULT_ERROR_SEVERITY, True)

        ' Protein name warnings

        ' Note that .*? changes .* from being greedy to being lazy
        Me.SetRule(RuleTypes.ProteinName, "[:|].*?[:|;].*?[:|;]", True, "Protein name contains 3 or more vertical bars", DEFAULT_WARNING_SEVERITY + 1, True)

        If Not mAllowAllSymbolsInProteinNames Then
            Me.SetRule(RuleTypes.ProteinName, "[/()\[\],]", True, "Protein name contains undesirable characters", DEFAULT_WARNING_SEVERITY, True)
        End If

        ' Protein description warnings
        Me.SetRule(RuleTypes.ProteinDescription, """", True, "Protein description contains a quotation mark", DEFAULT_WARNING_SEVERITY)
        Me.SetRule(RuleTypes.ProteinDescription, "\t", True, "Protein description contains a tab character", DEFAULT_WARNING_SEVERITY)
        Me.SetRule(RuleTypes.ProteinDescription, "\\/", True, "Protein description contains an escaped slash: \/", DEFAULT_WARNING_SEVERITY)
        Me.SetRule(RuleTypes.ProteinDescription, "[\x00-\x08\x0E-\x1F]", True, "Protein description contains an escape code character", DEFAULT_ERROR_SEVERITY)
        Me.SetRule(RuleTypes.ProteinDescription, ".{900,}", True, MESSAGE_TEXT_PROTEIN_DESCRIPTION_TOO_LONG, DEFAULT_WARNING_SEVERITY + 1, False)

        ' Protein sequence errors
        Me.SetRule(RuleTypes.ProteinSequence, "[ \t]", True, "A space or tab was found in the residues", DEFAULT_ERROR_SEVERITY)

        If Not mAllowAsteriskInResidues Then
            Me.SetRule(RuleTypes.ProteinSequence, "\*", True, MESSAGE_TEXT_ASTERISK_IN_RESIDUES, DEFAULT_ERROR_SEVERITY)
        End If

        If Not mAllowDashInResidues Then
            Me.SetRule(RuleTypes.ProteinSequence, "\-", True, MESSAGE_TEXT_DASH_IN_RESIDUES, DEFAULT_ERROR_SEVERITY)
        End If

        ' Note: we look for a space, tab, asterisk, and dash with separate rules (defined above)
        ' Thus they are "allowed" by this RegEx, even though we may flag them as warnings with a different RegEx
        ' We look for non-standard amino acids with warning rules (defined below)
        Me.SetRule(RuleTypes.ProteinSequence, "[^A-Z \t\*\-]", True, "Invalid residues found", DEFAULT_ERROR_SEVERITY, True)

        ' Protein residue warnings
        ' MS-GF+ treats these residues as stop characters(meaning no identified peptide will ever contain B, J, O, U, X, or Z)

        ' SEQUEST uses mass 114.53494 for B (average of N and D)
        Me.SetRule(RuleTypes.ProteinSequence, "B", True, "Residues line contains B (non-standard amino acid for N or D)", DEFAULT_WARNING_SEVERITY - 1)

        ' Unsupported by SEQUEST
        Me.SetRule(RuleTypes.ProteinSequence, "J", True, "Residues line contains J (non-standard amino acid)", DEFAULT_WARNING_SEVERITY - 1)

        ' SEQUEST uses mass 114.07931 for O
        Me.SetRule(RuleTypes.ProteinSequence, "O", True, "Residues line contains O (non-standard amino acid, ornithine)", DEFAULT_WARNING_SEVERITY - 1)

        ' Unsupported by SEQUEST
        Me.SetRule(RuleTypes.ProteinSequence, "U", True, "Residues line contains U (non-standard amino acid, selenocysteine)", DEFAULT_WARNING_SEVERITY)

        ' SEQUEST uses mass 113.08406 for X (same as L and I)
        Me.SetRule(RuleTypes.ProteinSequence, "X", True, "Residues line contains X (non-standard amino acid for L or I)", DEFAULT_WARNING_SEVERITY - 1)

        ' SEQUEST uses mass 128.55059 for Z (average of Q and E)
        Me.SetRule(RuleTypes.ProteinSequence, "Z", True, "Residues line contains Z (non-standard amino acid for Q or E)", DEFAULT_WARNING_SEVERITY - 1)

    End Sub

    Private Sub SetLocalErrorCode(eNewErrorCode As eValidateFastaFileErrorCodes)
        SetLocalErrorCode(eNewErrorCode, False)
    End Sub

    Private Sub SetLocalErrorCode(
      eNewErrorCode As eValidateFastaFileErrorCodes,
      leaveExistingErrorCodeUnchanged As Boolean)

        If leaveExistingErrorCodeUnchanged AndAlso mLocalErrorCode <> eValidateFastaFileErrorCodes.NoError Then
            ' An error code is already defined; do not change it
        Else
            mLocalErrorCode = eNewErrorCode

            If eNewErrorCode = eValidateFastaFileErrorCodes.NoError Then
                If MyBase.ErrorCode = ProcessFilesErrorCodes.LocalizedError Then
                    MyBase.SetBaseClassErrorCode(ProcessFilesErrorCodes.NoError)
                End If
            Else
                MyBase.SetBaseClassErrorCode(ProcessFilesErrorCodes.LocalizedError)
            End If
        End If

    End Sub

    Private Sub SetRule(
      ruleType As RuleTypes,
      regexToMatch As String,
      doesMatchIndicateProblem As Boolean,
      problemReturnMessage As String,
      severityLevel As Short)

        Me.SetRule(
         ruleType, regexToMatch,
         doesMatchIndicateProblem,
         problemReturnMessage,
         severityLevel, False)

    End Sub

    Private Sub SetRule(
      ruleType As RuleTypes,
      regexToMatch As String,
      doesMatchIndicateProblem As Boolean,
      problemReturnMessage As String,
      severityLevel As Short,
      displayMatchAsExtraInfo As Boolean)

        Select Case ruleType
            Case RuleTypes.HeaderLine
                SetRule(Me.mHeaderLineRules, regexToMatch, doesMatchIndicateProblem, problemReturnMessage, severityLevel, displayMatchAsExtraInfo)
            Case RuleTypes.ProteinDescription
                SetRule(Me.mProteinDescriptionRules, regexToMatch, doesMatchIndicateProblem, problemReturnMessage, severityLevel, displayMatchAsExtraInfo)
            Case RuleTypes.ProteinName
                SetRule(Me.mProteinNameRules, regexToMatch, doesMatchIndicateProblem, problemReturnMessage, severityLevel, displayMatchAsExtraInfo)
            Case RuleTypes.ProteinSequence
                SetRule(Me.mProteinSequenceRules, regexToMatch, doesMatchIndicateProblem, problemReturnMessage, severityLevel, displayMatchAsExtraInfo)
        End Select

    End Sub

    Private Sub SetRule(
      ByRef rules() As udtRuleDefinitionType,
      matchRegEx As String,
      matchIndicatesProblem As Boolean,
      messageWhenProblem As String,
      severity As Short,
      displayMatchAsExtraInfo As Boolean)

        If rules Is Nothing OrElse rules.Length = 0 Then
            ReDim rules(0)
        Else
            ReDim Preserve rules(rules.Length)
        End If

        With rules(rules.Length - 1)
            .MatchRegEx = matchRegEx
            .MatchIndicatesProblem = matchIndicatesProblem
            .MessageWhenProblem = messageWhenProblem
            .Severity = severity
            .DisplayMatchAsExtraInfo = displayMatchAsExtraInfo
            .CustomRuleID = mMasterCustomRuleID
        End With

        mMasterCustomRuleID += 1

    End Sub

    Private Function SortFile(proteinHashFile As FileInfo, sortColumnIndex As Integer, sortedFilePath As String) As Boolean

        Dim sortUtility = New FlexibleFileSortUtility.TextFileSorter

        mSortUtilityErrorMessage = String.Empty

        sortUtility.WorkingDirectoryPath = proteinHashFile.Directory.FullName
        sortUtility.HasHeaderLine = True
        sortUtility.ColumnDelimiter = ControlChars.Tab
        sortUtility.MaxFileSizeMBForInMemorySort = 250
        sortUtility.ChunkSizeMB = 250
        sortUtility.SortColumn = sortColumnIndex + 1
        sortUtility.SortColumnIsNumeric = False

        ' The sort utility uses CompareOrdinal (StringComparison.Ordinal)
        sortUtility.IgnoreCase = False

        RegisterEvents(sortUtility)
        AddHandler sortUtility.ErrorEvent, AddressOf mSortUtility_ErrorEvent

        Dim success = sortUtility.SortFile(proteinHashFile.FullName, sortedFilePath)

        If success Then
            Console.WriteLine()
            Return True
        End If

        If String.IsNullOrWhiteSpace(mSortUtilityErrorMessage) Then
            ShowErrorMessage("Unknown error sorting " & proteinHashFile.Name)
        Else
            ShowErrorMessage("Sort error: " & mSortUtilityErrorMessage)
        End If

        Console.WriteLine()
        Return False

    End Function

    Private Sub SplitFastaProteinHeaderLine(
      headerLine As String,
      <Out> ByRef proteinName As String,
      <Out> ByRef proteinDescription As String,
      <Out> ByRef descriptionStartIndex As Integer)

        proteinDescription = String.Empty
        descriptionStartIndex = 0

        ' Make sure the protein name and description are valid
        ' Find the first space and/or tab
        Dim spaceIndex = GetBestSpaceIndex(headerLine)

        ' At this point, spaceIndex will contain the location of the space or tab separating the protein name and description
        ' However, if the space or tab is directly after the > sign, then we cannot continue (if this is the case, then spaceIndex will be 1)
        If spaceIndex > 1 Then

            proteinName = headerLine.Substring(1, spaceIndex - 1)
            proteinDescription = headerLine.Substring(spaceIndex + 1)
            descriptionStartIndex = spaceIndex

        Else
            ' Line does not contain a description
            If spaceIndex <= 0 Then
                If headerLine.Trim.Length <= 1 Then
                    proteinName = String.Empty
                Else
                    ' The line contains a protein name, but not a description
                    proteinName = headerLine.Substring(1)
                End If
            Else
                ' Space or tab found directly after the > symbol
                proteinName = String.Empty
            End If
        End If

    End Sub

    Private Function VerifyLinefeedAtEOF(strInputFilePath As String, blnAddCrLfIfMissing As Boolean) As Boolean

        Dim blnNeedToAddCrLf As Boolean
        Dim blnSuccess As Boolean

        Try
            ' Open the input file and validate that the final characters are CrLf, simply CR, or simply LF
            Using fsInFile = New FileStream(strInputFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)

                If fsInFile.Length > 2 Then
                    fsInFile.Seek(-1, SeekOrigin.End)

                    Dim lastByte = fsInFile.ReadByte()

                    If lastByte = 10 Or lastByte = 13 Then
                        ' File ends in a linefeed or carriage return character; that's good
                        blnNeedToAddCrLf = False
                    Else
                        blnNeedToAddCrLf = True
                    End If
                End If

                If blnNeedToAddCrLf Then
                    If blnAddCrLfIfMissing Then
                        ShowMessage("Appending CrLf return to: " & Path.GetFileName(strInputFilePath))
                        fsInFile.WriteByte(13)

                        fsInFile.WriteByte(10)
                    End If
                End If

            End Using

            blnSuccess = True

        Catch ex As Exception
            SetLocalErrorCode(eValidateFastaFileErrorCodes.ErrorVerifyingLinefeedAtEOF)
            blnSuccess = False
        End Try

        Return blnSuccess

    End Function

    Private Sub WriteCachedProtein(
      cachedProteinName As String,
      cachedProteinDescription As String,
      consolidatedFastaWriter As TextWriter,
      proteinSeqHashInfo As IList(Of clsProteinHashInfo),
      sbCachedProteinResidueLines As StringBuilder,
      sbCachedProteinResidues As StringBuilder,
      consolidateDuplicateProteinSeqsInFasta As Boolean,
      consolidateDupsIgnoreILDiff As Boolean,
      proteinNameFirst As clsNestedStringDictionary(Of Integer),
      duplicateProteinList As clsNestedStringDictionary(Of String),
      lineCountRead As Integer,
      proteinsWritten As clsNestedStringDictionary(Of Integer))

        Static reAdditionalProtein As Regex = New Regex("(.+)-[a-z]\d*", RegexOptions.Compiled)

        Dim masterProteinName As String = String.Empty
        Dim masterProteinInfo As String

        Dim proteinHash As String
        Dim lineOut As String = String.Empty
        Dim reMatch As Match

        Dim keepProtein As Boolean
        Dim skipDupProtein As Boolean

        Dim additionalProteinNames = New List(Of String)

        Dim seqIndex As Integer
        If proteinNameFirst.TryGetValue(cachedProteinName, seqIndex) Then
            ' cachedProteinName was found in proteinNameFirst

            Dim cachedSeqIndex As Integer
            If proteinsWritten.TryGetValue(cachedProteinName, cachedSeqIndex) Then
                If mFixedFastaOptions.KeepDuplicateNamedProteinsUnlessMatchingSequence Then
                    ' Keep this protein if its sequence hash differs from the first protein with this name
                    proteinHash = ComputeProteinHash(sbCachedProteinResidues, consolidateDupsIgnoreILDiff)
                    If proteinSeqHashInfo(seqIndex).SequenceHash <> proteinHash Then
                        RecordFastaFileWarning(lineCountRead, 1, cachedProteinName, eMessageCodeConstants.DuplicateProteinNameRetained)
                        keepProtein = True
                    Else
                        keepProtein = False
                    End If
                Else
                    keepProtein = False
                End If
            Else
                keepProtein = True
                proteinsWritten.Add(cachedProteinName, seqIndex)
            End If

            If keepProtein AndAlso seqIndex >= 0 Then
                If proteinSeqHashInfo(seqIndex).AdditionalProteins.Count > 0 Then
                    ' The protein has duplicate proteins
                    ' Construct a list of the duplicate protein names

                    additionalProteinNames.Clear()
                    For Each additionalProtein As String In proteinSeqHashInfo(seqIndex).AdditionalProteins
                        ' Add the additional protein name if it is not of the form "BaseName-b", "BaseName-c", etc.
                        skipDupProtein = False

                        If additionalProtein Is Nothing Then
                            skipDupProtein = True
                        Else
                            If additionalProtein.ToLower() = cachedProteinName.ToLower() Then
                                ' Names match; do not add to the list
                                skipDupProtein = True
                            Else
                                ' Check whether additionalProtein looks like one of the following
                                ' ProteinX-b
                                ' ProteinX-a2
                                ' ProteinX-d3
                                reMatch = reAdditionalProtein.Match(additionalProtein)

                                If reMatch.Success Then

                                    If cachedProteinName.ToLower() = reMatch.Groups(1).Value.ToLower() Then
                                        ' Base names match; do not add to the list
                                        ' For example, ProteinX and ProteinX-b
                                        skipDupProtein = True
                                    End If
                                End If

                            End If
                        End If

                        If Not skipDupProtein Then
                            additionalProteinNames.Add(additionalProtein)
                        End If
                    Next

                    If additionalProteinNames.Count > 0 AndAlso consolidateDuplicateProteinSeqsInFasta Then
                        ' Append the duplicate protein names to the description
                        ' However, do not let the description get over 7995 characters in length
                        Dim updatedDescription = cachedProteinDescription & "; Duplicate proteins: " & FlattenArray(additionalProteinNames, ","c)
                        If updatedDescription.Length > MAX_PROTEIN_DESCRIPTION_LENGTH Then
                            updatedDescription = updatedDescription.Substring(0, MAX_PROTEIN_DESCRIPTION_LENGTH - 3) & "..."
                        End If

                        lineOut = ConstructFastaHeaderLine(cachedProteinName, updatedDescription)
                    End If
                End If
            End If
        Else
            keepProtein = False
            mFixedFastaStats.DuplicateSequenceProteinsSkipped += 1

            If Not duplicateProteinList.TryGetValue(cachedProteinName, masterProteinName) Then
                masterProteinInfo = "same as ??"
            Else
                masterProteinInfo = "same as " & masterProteinName
            End If
            RecordFastaFileWarning(lineCountRead, 0, cachedProteinName, eMessageCodeConstants.ProteinRemovedSinceDuplicateSequence, masterProteinInfo, String.Empty)
        End If

        If keepProtein Then
            If String.IsNullOrEmpty(lineOut) Then
                lineOut = ConstructFastaHeaderLine(cachedProteinName, cachedProteinDescription)
            End If
            consolidatedFastaWriter.WriteLine(lineOut)
            consolidatedFastaWriter.Write(sbCachedProteinResidueLines.ToString())
        End If

    End Sub

    Private Sub WriteCachedProteinHashMetadata(
      proteinNamesToKeepWriter As TextWriter,
      uniqueProteinSeqsWriter As TextWriter,
      uniqueProteinSeqDuplicateWriter As TextWriter,
      currentSequenceIndex As Integer,
      sequenceHash As String,
      sequenceLength As String,
      proteinNames As ICollection(Of String))

        If proteinNames.Count < 1 Then
            Throw New Exception("proteinNames is empty in WriteCachedProteinHashMetadata; this indicates a logic error")
        End If

        Dim firstProtein = proteinNames(0)
        Dim duplicateProteinList As String

        ' Append the first protein name to the _ProteinsToKeep.tmp file
        proteinNamesToKeepWriter.WriteLine(firstProtein)

        If proteinNames.Count = 1 Then
            duplicateProteinList = String.Empty
        Else

            Dim duplicateProteins = New StringBuilder()

            For proteinIndex = 1 To proteinNames.Count - 1

                If duplicateProteins.Length > 0 Then
                    duplicateProteins.Append(",")
                End If
                duplicateProteins.Append(proteinNames(proteinIndex))

                Dim dataValuesSeqDuplicate = New List(Of String) From {
                    currentSequenceIndex.ToString(),
                    firstProtein,
                    sequenceLength,
                    proteinNames(proteinIndex)}

                uniqueProteinSeqDuplicateWriter.WriteLine(FlattenList(dataValuesSeqDuplicate))

            Next

            duplicateProteinList = duplicateProteins.ToString()
        End If

        Dim dataValues = New List(Of String) From {
           currentSequenceIndex.ToString(),
           proteinNames(0),
           sequenceLength,
           sequenceHash,
           proteinNames.Count.ToString(),
           duplicateProteinList}

        uniqueProteinSeqsWriter.WriteLine(FlattenList(dataValues))

    End Sub

#Region "Event Handlers"

    Private Sub mSortUtility_ErrorEvent(message As String, ex As Exception)
        mSortUtilityErrorMessage = message
    End Sub

#End Region

    ' IComparer class to allow comparison of udtMsgInfoType items
    Private Class ErrorInfoComparerClass
        Implements IComparer

        Public Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare

            Dim errorInfo1, errorInfo2 As udtMsgInfoType

            errorInfo1 = CType(x, udtMsgInfoType)
            errorInfo2 = CType(y, udtMsgInfoType)

            If errorInfo1.MessageCode > errorInfo2.MessageCode Then
                Return 1
            ElseIf errorInfo1.MessageCode < errorInfo2.MessageCode Then
                Return -1
            Else
                If errorInfo1.LineNumber > errorInfo2.LineNumber Then
                    Return 1
                ElseIf errorInfo1.LineNumber < errorInfo2.LineNumber Then
                    Return -1
                Else
                    If errorInfo1.ColNumber > errorInfo2.ColNumber Then
                        Return 1
                    ElseIf errorInfo1.ColNumber < errorInfo2.ColNumber Then
                        Return -1
                    Else
                        Return 0
                    End If
                End If
            End If

        End Function
    End Class

End Class
