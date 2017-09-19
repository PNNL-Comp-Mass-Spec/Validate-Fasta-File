Option Strict On

' This class will read a protein fasta file and validate its contents
'
' -------------------------------------------------------------------------------
' Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
' Program started March 21, 2005

' E-mail: matthew.monroe@pnnl.gov or matt@alchemistmatt.com
' Website: http://omics.pnl.gov/ or http://www.sysbio.org/resources/staff/ or http://panomics.pnnl.gov/
' -------------------------------------------------------------------------------
'
' Licensed under the Apache License, Version 2.0; you may not use this file except
' in compliance with the License.  You may obtain a copy of the License at
' http://www.apache.org/licenses/LICENSE-2.0
'
' Notice: This computer software was prepared by Battelle Memorial Institute,
' hereinafter the Contractor, under Contract No. DE-AC05-76RL0 1830 with the
' Department of Energy (DOE).  All rights in the computer software are reserved
' by DOE on behalf of the United States Government and the Contractor as
' provided in the Contract.  NEITHER THE GOVERNMENT NOR THE CONTRACTOR MAKES ANY
' WARRANTY, EXPRESS OR IMPLIED, OR ASSUMES ANY LIABILITY FOR THE USE OF THIS
' SOFTWARE.  This notice including this sentence must appear on any copies of
' this computer software.

Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports PRISM

Public Class clsValidateFastaFile
    Inherits clsProcessFilesBaseClass
    Implements IValidateFastaFile

    Public Sub New()
        MyBase.mFileDate = "March 23, 2016"
        InitializeLocalVariables()
    End Sub

    Public Sub New(ParameterFilePath As String)
        Me.New()
        LoadParameterFileSettings(ParameterFilePath)
    End Sub


#Region "Constants and Enums"
    Private Const DEFAULT_MINIMUM_PROTEIN_NAME_LENGTH As Integer = 3

    ''' <summary>
    ''' The maximum suggested value when using SEQUEST is 34 characters
    ''' In contrast, MSGF+ supports long protein names
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
    Private mFileErrors() As IValidateFastaFile.udtMsgInfoType
    Private mFileErrorStats As udtItemSummaryIndexedType

    Private mFileWarningCount As Integer

    Private mFileWarnings() As IValidateFastaFile.udtMsgInfoType
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
    ''' The number of characters at the start of keystrings to use when adding items to clsNestedStringDictionary instances
    ''' </summary>
    ''' <remarks>
    ''' If this value is too short, all of the items added to the clsNestedStringDictionary instance
    ''' will be tracked by the same dictionary, which could result in a dictionary surpassing the 2 GB boundary
    ''' </remarks>
    Private mProteinNameSpannerCharLength As Byte = 1

    Private mLocalErrorCode As IValidateFastaFile.eValidateFastaFileErrorCodes

    Private mMemoryUsageLogger As clsMemoryUsageLogger

    Private mProcessMemoryUsageMBAtStart As Single

    Private mSortUtilityErrorMessage As String
    Private mLastSortUtilityProgress As DateTime

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
    Public Property OptionSwitch(SwitchName As IValidateFastaFile.SwitchOptions) As Boolean _
     Implements IValidateFastaFile.OptionSwitches
        Get
            Return GetOptionSwitchValue(SwitchName)
        End Get
        Set(value As Boolean)
            SetOptionSwitch(SwitchName, value)
        End Set
    End Property

    ''' <summary>
    ''' Set a processing option
    ''' </summary>
    ''' <param name="SwitchName"></param>
    ''' <param name="State"></param>
    ''' <remarks>Be sure to call SetDefaultRules() after setting all of the options</remarks>
    Public Sub SetOptionSwitch(SwitchName As IValidateFastaFile.SwitchOptions, State As Boolean)

        Select Case SwitchName
            Case IValidateFastaFile.SwitchOptions.AddMissingLinefeedatEOF
                mAddMissingLinefeedAtEOF = State
            Case IValidateFastaFile.SwitchOptions.AllowAsteriskInResidues
                mAllowAsteriskInResidues = State
            Case IValidateFastaFile.SwitchOptions.CheckForDuplicateProteinNames
                mCheckForDuplicateProteinNames = State
            Case IValidateFastaFile.SwitchOptions.GenerateFixedFASTAFile
                mGenerateFixedFastaFile = State
            Case IValidateFastaFile.SwitchOptions.OutputToStatsFile
                mOutputToStatsFile = State
            Case IValidateFastaFile.SwitchOptions.SplitOutMultipleRefsInProteinName
                mFixedFastaOptions.SplitOutMultipleRefsInProteinName = State
            Case IValidateFastaFile.SwitchOptions.WarnBlankLinesBetweenProteins
                mWarnBlankLinesBetweenProteins = State
            Case IValidateFastaFile.SwitchOptions.WarnLineStartsWithSpace
                mWarnLineStartsWithSpace = State
            Case IValidateFastaFile.SwitchOptions.NormalizeFileLineEndCharacters
                mNormalizeFileLineEndCharacters = State
            Case IValidateFastaFile.SwitchOptions.CheckForDuplicateProteinSequences
                mCheckForDuplicateProteinSequences = State
            Case IValidateFastaFile.SwitchOptions.FixedFastaRenameDuplicateNameProteins
                mFixedFastaOptions.RenameProteinsWithDuplicateNames = State
            Case IValidateFastaFile.SwitchOptions.FixedFastaKeepDuplicateNamedProteins
                mFixedFastaOptions.KeepDuplicateNamedProteinsUnlessMatchingSequence = State
            Case IValidateFastaFile.SwitchOptions.SaveProteinSequenceHashInfoFiles
                mSaveProteinSequenceHashInfoFiles = State
            Case IValidateFastaFile.SwitchOptions.FixedFastaConsolidateDuplicateProteinSeqs
                mFixedFastaOptions.ConsolidateProteinsWithDuplicateSeqs = State
            Case IValidateFastaFile.SwitchOptions.FixedFastaConsolidateDupsIgnoreILDiff
                mFixedFastaOptions.ConsolidateDupsIgnoreILDiff = State
            Case IValidateFastaFile.SwitchOptions.FixedFastaTruncateLongProteinNames
                mFixedFastaOptions.TruncateLongProteinNames = State
            Case IValidateFastaFile.SwitchOptions.FixedFastaSplitOutMultipleRefsForKnownAccession
                mFixedFastaOptions.SplitOutMultipleRefsForKnownAccession = State
            Case IValidateFastaFile.SwitchOptions.FixedFastaWrapLongResidueLines
                mFixedFastaOptions.WrapLongResidueLines = State
            Case IValidateFastaFile.SwitchOptions.FixedFastaRemoveInvalidResidues
                mFixedFastaOptions.RemoveInvalidResidues = State
            Case IValidateFastaFile.SwitchOptions.SaveBasicProteinHashInfoFile
                mSaveBasicProteinHashInfoFile = State
            Case IValidateFastaFile.SwitchOptions.AllowDashInResidues
                mAllowDashInResidues = State
            Case IValidateFastaFile.SwitchOptions.AllowAllSymbolsInProteinNames
                mAllowAllSymbolsInProteinNames = State
        End Select

    End Sub

    Public Function GetOptionSwitchValue(SwitchName As IValidateFastaFile.SwitchOptions) As Boolean

        Select Case SwitchName
            Case IValidateFastaFile.SwitchOptions.AddMissingLinefeedatEOF
                Return mAddMissingLinefeedAtEOF
            Case IValidateFastaFile.SwitchOptions.AllowAsteriskInResidues
                Return mAllowAsteriskInResidues
            Case IValidateFastaFile.SwitchOptions.CheckForDuplicateProteinNames
                Return mCheckForDuplicateProteinNames
            Case IValidateFastaFile.SwitchOptions.GenerateFixedFASTAFile
                Return mGenerateFixedFastaFile
            Case IValidateFastaFile.SwitchOptions.OutputToStatsFile
                Return mOutputToStatsFile
            Case IValidateFastaFile.SwitchOptions.SplitOutMultipleRefsInProteinName
                Return mFixedFastaOptions.SplitOutMultipleRefsInProteinName
            Case IValidateFastaFile.SwitchOptions.WarnBlankLinesBetweenProteins
                Return mWarnBlankLinesBetweenProteins
            Case IValidateFastaFile.SwitchOptions.WarnLineStartsWithSpace
                Return mWarnLineStartsWithSpace
            Case IValidateFastaFile.SwitchOptions.NormalizeFileLineEndCharacters
                Return mNormalizeFileLineEndCharacters
            Case IValidateFastaFile.SwitchOptions.CheckForDuplicateProteinSequences
                Return mCheckForDuplicateProteinSequences
            Case IValidateFastaFile.SwitchOptions.FixedFastaRenameDuplicateNameProteins
                Return mFixedFastaOptions.RenameProteinsWithDuplicateNames
            Case IValidateFastaFile.SwitchOptions.FixedFastaKeepDuplicateNamedProteins
                Return mFixedFastaOptions.KeepDuplicateNamedProteinsUnlessMatchingSequence
            Case IValidateFastaFile.SwitchOptions.SaveProteinSequenceHashInfoFiles
                Return mSaveProteinSequenceHashInfoFiles
            Case IValidateFastaFile.SwitchOptions.FixedFastaConsolidateDuplicateProteinSeqs
                Return mFixedFastaOptions.ConsolidateProteinsWithDuplicateSeqs
            Case IValidateFastaFile.SwitchOptions.FixedFastaConsolidateDupsIgnoreILDiff
                Return mFixedFastaOptions.ConsolidateDupsIgnoreILDiff
            Case IValidateFastaFile.SwitchOptions.FixedFastaTruncateLongProteinNames
                Return mFixedFastaOptions.TruncateLongProteinNames
            Case IValidateFastaFile.SwitchOptions.FixedFastaSplitOutMultipleRefsForKnownAccession
                Return mFixedFastaOptions.SplitOutMultipleRefsForKnownAccession
            Case IValidateFastaFile.SwitchOptions.FixedFastaWrapLongResidueLines
                Return mFixedFastaOptions.WrapLongResidueLines
            Case IValidateFastaFile.SwitchOptions.FixedFastaRemoveInvalidResidues
                Return mFixedFastaOptions.RemoveInvalidResidues
            Case IValidateFastaFile.SwitchOptions.SaveBasicProteinHashInfoFile
                Return mSaveBasicProteinHashInfoFile
            Case IValidateFastaFile.SwitchOptions.AllowDashInResidues
                Return mAllowDashInResidues
            Case IValidateFastaFile.SwitchOptions.AllowAllSymbolsInProteinNames
                Return mAllowAllSymbolsInProteinNames
        End Select

        Return False

    End Function

    Public ReadOnly Property ErrorWarningCounts(
      messageType As IValidateFastaFile.eMsgTypeConstants,
      CountType As IValidateFastaFile.ErrorWarningCountTypes) As Integer

        Get
            Dim tmpValue As Integer
            Select Case CountType
                Case IValidateFastaFile.ErrorWarningCountTypes.Total
                    Select Case messageType
                        Case IValidateFastaFile.eMsgTypeConstants.ErrorMsg
                            tmpValue = mFileErrorCount + ComputeTotalUnspecifiedCount(mFileErrorStats)
                        Case IValidateFastaFile.eMsgTypeConstants.WarningMsg
                            tmpValue = mFileWarningCount + ComputeTotalUnspecifiedCount(mFileWarningStats)
                        Case IValidateFastaFile.eMsgTypeConstants.StatusMsg
                            tmpValue = 0
                    End Select
                Case IValidateFastaFile.ErrorWarningCountTypes.Unspecified
                    Select Case messageType
                        Case IValidateFastaFile.eMsgTypeConstants.ErrorMsg
                            tmpValue = ComputeTotalUnspecifiedCount(mFileErrorStats)
                        Case IValidateFastaFile.eMsgTypeConstants.WarningMsg
                            tmpValue = ComputeTotalSpecifiedCount(mFileWarningStats)
                        Case IValidateFastaFile.eMsgTypeConstants.StatusMsg
                            tmpValue = 0
                    End Select
                Case IValidateFastaFile.ErrorWarningCountTypes.Specified
                    Select Case messageType
                        Case IValidateFastaFile.eMsgTypeConstants.ErrorMsg
                            tmpValue = mFileErrorCount
                        Case IValidateFastaFile.eMsgTypeConstants.WarningMsg
                            tmpValue = mFileWarningCount
                        Case IValidateFastaFile.eMsgTypeConstants.StatusMsg
                            tmpValue = 0
                    End Select
            End Select

            Return tmpValue
        End Get
    End Property

    Public ReadOnly Property FixedFASTAFileStats(
     ValueType As IValidateFastaFile.FixedFASTAFileValues) As Integer _
     Implements IValidateFastaFile.FixedFASTAFileStats

        Get
            Dim tmpValue As Integer
            Select Case ValueType
                Case IValidateFastaFile.FixedFASTAFileValues.DuplicateProteinNamesSkippedCount
                    tmpValue = mFixedFastaStats.DuplicateNameProteinsSkipped
                Case IValidateFastaFile.FixedFASTAFileValues.ProteinNamesInvalidCharsReplaced
                    tmpValue = mFixedFastaStats.ProteinNamesInvalidCharsReplaced
                Case IValidateFastaFile.FixedFASTAFileValues.ProteinNamesMultipleRefsRemoved
                    tmpValue = mFixedFastaStats.ProteinNamesMultipleRefsRemoved
                Case IValidateFastaFile.FixedFASTAFileValues.TruncatedProteinNameCount
                    tmpValue = mFixedFastaStats.TruncatedProteinNameCount
                Case IValidateFastaFile.FixedFASTAFileValues.UpdatedResidueLines
                    tmpValue = mFixedFastaStats.UpdatedResidueLines
                Case IValidateFastaFile.FixedFASTAFileValues.DuplicateProteinNamesRenamedCount
                    tmpValue = mFixedFastaStats.DuplicateNameProteinsRenamed
                Case IValidateFastaFile.FixedFASTAFileValues.DuplicateProteinSeqsSkippedCount
                    tmpValue = mFixedFastaStats.DuplicateSequenceProteinsSkipped
            End Select
            Return tmpValue

        End Get
    End Property

    Public ReadOnly Property ProteinCount() As Integer Implements IValidateFastaFile.ProteinCount
        Get
            Return mProteinCount
        End Get
    End Property

    Public ReadOnly Property LineCount() As Integer Implements IValidateFastaFile.FileLineCount
        Get
            Return mLineCount
        End Get
    End Property

    Public ReadOnly Property LocalErrorCode() As IValidateFastaFile.eValidateFastaFileErrorCodes _
     Implements IValidateFastaFile.LocalErrorCode
        Get
            Return mLocalErrorCode
        End Get
    End Property

    Public ReadOnly Property ResidueCount() As Long Implements IValidateFastaFile.ResidueCount
        Get
            Return mResidueCount
        End Get
    End Property

    Public ReadOnly Property FastaFilePath() As String Implements IValidateFastaFile.FASTAFilePath
        Get
            Return mFastaFilePath
        End Get
    End Property

    Public ReadOnly Property ErrorMessageTextByIndex(
     index As Integer,
     valueSeparator As String) As String _
      Implements IValidateFastaFile.ErrorMessageTextByIndex
        Get
            Return GetFileErrorTextByIndex(index, valueSeparator)
        End Get
    End Property

    Public ReadOnly Property WarningMessageTextByIndex(
     index As Integer,
     valueSeparator As String) As String _
      Implements IValidateFastaFile.WarningMessageTextByIndex
        Get
            Return GetFileWarningTextByIndex(index, valueSeparator)
        End Get
    End Property

    Public ReadOnly Property ErrorsByIndex(errorIndex As Integer) As IValidateFastaFile.udtMsgInfoType _
     Implements IValidateFastaFile.FileErrorByIndex
        Get
            Return (GetFileErrorByIndex(errorIndex))
        End Get
    End Property

    Public ReadOnly Property WarningsByIndex(warningIndex As Integer) As IValidateFastaFile.udtMsgInfoType _
     Implements IValidateFastaFile.FileWarningByIndex
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

    Public Property MaximumFileErrorsToTrack() As Integer _
     Implements IValidateFastaFile.MaximumFileErrorsToTrack
        Get
            Return mMaximumFileErrorsToTrack
        End Get
        Set(Value As Integer)
            If Value < 1 Then Value = 1
            mMaximumFileErrorsToTrack = Value
        End Set
    End Property

    Public Property MaximumProteinNameLength() As Integer _
     Implements IValidateFastaFile.MaximumProteinNameLength
        Get
            Return mMaximumProteinNameLength
        End Get
        Set(Value As Integer)
            If Value < 8 Then
                ' Do not allow maximum lengths less than 8; use the default
                Value = DEFAULT_MAXIMUM_PROTEIN_NAME_LENGTH
            End If
            mMaximumProteinNameLength = Value
        End Set
    End Property

    Public Property MinimumProteinNameLength() As Integer _
     Implements IValidateFastaFile.MinimumProteinNameLength
        Get
            Return mMinimumProteinNameLength
        End Get
        Set(Value As Integer)
            If Value < 1 Then Value = DEFAULT_MINIMUM_PROTEIN_NAME_LENGTH
            mMinimumProteinNameLength = Value
        End Set
    End Property

    Public Property MaximumResiduesPerLine() As Integer _
     Implements IValidateFastaFile.MaximumResiduesPerLine
        Get
            Return mMaximumResiduesPerLine
        End Get
        Set(Value As Integer)
            If Value = 0 Then
                Value = DEFAULT_MAXIMUM_RESIDUES_PER_LINE
            ElseIf Value < 40 Then
                Value = 40
            End If

            mMaximumResiduesPerLine = Value
        End Set
    End Property

    Public Property ProteinLineStartChar() As Char _
     Implements IValidateFastaFile.ProteinLineStartCharacter
        Get
            Return mProteinLineStartChar
        End Get
        Set(Value As Char)
            mProteinLineStartChar = Value
        End Set
    End Property

    Public ReadOnly Property StatsFilePath() As String
        Get
            If mStatsFilePath Is Nothing Then
                Return String.Empty
            Else
                Return mStatsFilePath
            End If
        End Get
    End Property

    Public Property ProteinNameInvalidCharsToRemove() As String _
     Implements IValidateFastaFile.ProteinNameInvalidCharactersToRemove
        Get
            Return CharArrayToString(mFixedFastaOptions.ProteinNameInvalidCharsToRemove)
        End Get
        Set(Value As String)
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

    Public Property ProteinNameFirstRefSepChars() As String _
       Implements IValidateFastaFile.ProteinNameFirstRefSepChars
        Get
            Return CharArrayToString(mProteinNameFirstRefSepChars)
        End Get
        Set(Value As String)
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

    Public Property ProteinNameSubsequentRefSepChars() As String _
       Implements IValidateFastaFile.ProteinNameSubsequentRefSepChars
        Get
            Return CharArrayToString(mProteinNameSubsequentRefSepChars)
        End Get
        Set(Value As String)
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

    Public Property LongProteinNameSplitChars() As String _
     Implements IValidateFastaFile.LongProteinNameSplitChars
        Get
            Return CharArrayToString(mFixedFastaOptions.LongProteinNameSplitChars)
        End Get
        Set(Value As String)
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

    Public ReadOnly Property FileWarningList() As IValidateFastaFile.udtMsgInfoType() _
     Implements IValidateFastaFile.FileWarningList
        Get
            Return GetFileWarnings()
        End Get
    End Property

    Public ReadOnly Property FileErrorList() As IValidateFastaFile.udtMsgInfoType() _
     Implements IValidateFastaFile.FileErrorList
        Get
            Return GetFileErrors()
        End Get
    End Property

    Public Shadows Property ShowMessages() As Boolean Implements IValidateFastaFile.ShowMessages
        Get
            Return MyBase.ShowMessages
        End Get
        Set(Value As Boolean)
            MyBase.ShowMessages = Value
        End Set
    End Property

#End Region

    Private Event ProgressUpdated(
     taskDescription As String,
     percentComplete As Single) Implements IValidateFastaFile.ProgressChanged

    Private Event ProgressCompleted() Implements IValidateFastaFile.ProgressCompleted

    Private Event WroteLineEndNormalizedFASTA(newFilePath As String) Implements IValidateFastaFile.WroteLineEndNormalizedFASTA

    Private Sub OnProgressUpdate(
     taskDescription As String,
     percentComplete As Single) Handles MyBase.ProgressChanged

        RaiseEvent ProgressUpdated(taskDescription, percentComplete)

    End Sub

    Private Sub OnProgressComplete() Handles MyBase.ProgressComplete
        RaiseEvent ProgressCompleted()
    End Sub

    Private Sub OnWroteLineEndNormalizedFASTA(newFilePath As String)
        RaiseEvent WroteLineEndNormalizedFASTA(newFilePath)
    End Sub

    ''' <summary>
    ''' Examine the given fasta file to look for problems.
    ''' Optionally create a new, fixed fasta file
    ''' Optionally also consolidate proteins with duplicate sequences
    ''' </summary>
    ''' <param name="strFastaFilePath"></param>
    ''' <param name="lstPreloadedProteinNamesToKeep">
    ''' Preloaded list of protein names to include in the fixed fasta file
    ''' Keys are protein names, values are the number of entries written to the fixed fasta file for the given protein name
    ''' </param>
    ''' <returns>True if the file was successfully analyzed (even if errors were found)</returns>
    ''' <remarks>Assumes strFastaFilePath exists</remarks>
    Private Function AnalyzeFastaFile(strFastaFilePath As String, lstPreloadedProteinNamesToKeep As clsNestedStringIntList) As Boolean

        Dim swFixedFastaOut As StreamWriter = Nothing
        Dim swProteinSequenceHashBasic As StreamWriter = Nothing

        Dim strFastaFilePathOut = "UndefinedFilePath.xyz"

        Dim blnSuccess As Boolean

        Dim blnConsolidateDuplicateProteinSeqsInFasta = False
        Dim blnKeepDuplicateNamedProteinsUnlessMatchingSequence = False
        Dim blnConsolidateDupsIgnoreILDiff = False

        ' This array tracks protein hash details
        Dim intProteinSequenceHashCount As Integer
        Dim oProteinSeqHashInfo() As clsProteinHashInfo

        Dim udtHeaderLineRuleDetails() As udtRuleDefinitionExtendedType
        Dim udtProteinNameRuleDetails() As udtRuleDefinitionExtendedType
        Dim udtProteinDescriptionRuleDetails() As udtRuleDefinitionExtendedType
        Dim udtProteinSequenceRuleDetails() As udtRuleDefinitionExtendedType

        Try
            ' Reset the data structures and variables
            ResetStructures()
            ReDim oProteinSeqHashInfo(0)

            ReDim udtHeaderLineRuleDetails(1)
            ReDim udtProteinNameRuleDetails(1)
            ReDim udtProteinDescriptionRuleDetails(1)
            ReDim udtProteinSequenceRuleDetails(1)

            ' This is a dictionary of dictionaries, with one dictionary for each letter or number that a SHA-1 hash could start with
            ' This dictionary of dictionaries provides a quick lookup for existing protein hashes
            ' This dictionary is not used if lstPreloadedProteinNamesToKeep contains data
            Const SPANNER_CHAR_LENGTH = 1
            Dim lstProteinSequenceHashes = New clsNestedStringDictionary(Of Integer)(False, SPANNER_CHAR_LENGTH)
            Dim usingPreloadedProteinNames = False

            If Not lstPreloadedProteinNamesToKeep Is Nothing AndAlso lstPreloadedProteinNamesToKeep.Count > 0 Then
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
                 strFastaFilePath,
                 "CRLF_" & Path.GetFileName(strFastaFilePath),
                 IValidateFastaFile.eLineEndingCharacters.CRLF)

                If mFastaFilePath <> strFastaFilePath Then
                    strFastaFilePath = String.Copy(mFastaFilePath)
                    OnWroteLineEndNormalizedFASTA(strFastaFilePath)
                End If
            Else
                mFastaFilePath = String.Copy(strFastaFilePath)
            End If

            OnProgressUpdate("Parsing " & Path.GetFileName(mFastaFilePath), 0)

            Dim blnProteinHeaderFound = False
            Dim blnProcessingResidueBlock = False
            Dim blnBlankLineProcessed = False

            Dim strProteinName = String.Empty
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
            InitializeRuleDetails(mHeaderLineRules, udtHeaderLineRuleDetails)
            InitializeRuleDetails(mProteinNameRules, udtProteinNameRuleDetails)
            InitializeRuleDetails(mProteinDescriptionRules, udtProteinDescriptionRuleDetails)
            InitializeRuleDetails(mProteinSequenceRules, udtProteinSequenceRuleDetails)

            ' Open the file and read, at most, the first 100,000 characters to see if it contains CrLf or just Lf
            Dim intTerminatorSize = DetermineLineTerminatorSize(strFastaFilePath)

            ' Pre-scan a portion of the fasta file to determine the appropriate value for mProteinNameSpannerCharLength
            AutoDetermineFastaProteinNameSpannerCharLength(mFastaFilePath, intTerminatorSize)

            ' Open the input file

            Using fsInFile = New FileStream(strFastaFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite),
                srFastaInFile = New StreamReader(fsInFile)

                ' Optionally, open the output fasta file
                If mGenerateFixedFastaFile Then

                    Try
                        strFastaFilePathOut =
                         Path.Combine(Path.GetDirectoryName(strFastaFilePath),
                         Path.GetFileNameWithoutExtension(strFastaFilePath) & "_new.fasta")
                        swFixedFastaOut = New StreamWriter(New FileStream(strFastaFilePathOut, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
                    Catch ex As Exception
                        ' Error opening output file
                        RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError,
                         "Error creating output file " & strFastaFilePathOut & ": " & ex.Message, String.Empty)
                        ShowMessage(ex.Message)
                        ShowExceptionStackTrace("AnalyzeFastaFile (Create _new.fasta)", ex)
                        Return False
                    End Try
                End If

                ' Optionally, open the Sequence Hash file
                If mSaveBasicProteinHashInfoFile Then
                    Dim strBasicProteinHashInfoFilePath = "<undefined>"

                    Try
                        strBasicProteinHashInfoFilePath =
                         Path.Combine(Path.GetDirectoryName(strFastaFilePath),
                         Path.GetFileNameWithoutExtension(strFastaFilePath) & PROTEIN_HASHES_FILENAME_SUFFIX)
                        swProteinSequenceHashBasic = New StreamWriter(New FileStream(strBasicProteinHashInfoFilePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))

                        Dim headerNames = New List(Of String) From {
                            "Protein_ID",
                            PROTEIN_NAME_COLUMN,
                            SEQUENCE_LENGTH_COLUMN,
                            SEQUENCE_HASH_COLUMN}

                        swProteinSequenceHashBasic.WriteLine(FlattenList(headerNames))

                    Catch ex As Exception
                        ' Error opening output file
                        RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError,
                         "Error creating output file " & strBasicProteinHashInfoFilePath & ": " & ex.Message, String.Empty)
                        ShowMessage(ex.Message)
                        ShowExceptionStackTrace("AnalyzeFastaFile (Create " & PROTEIN_HASHES_FILENAME_SUFFIX & ")", ex)
                    End Try

                End If

                If mGenerateFixedFastaFile And (mFixedFastaOptions.RenameProteinsWithDuplicateNames OrElse mFixedFastaOptions.KeepDuplicateNamedProteinsUnlessMatchingSequence) Then
                    ' Make sure mCheckForDuplicateProteinNames is enabled
                    mCheckForDuplicateProteinNames = True
                End If

                ' Initialize lstProteinNames
                Dim lstProteinNames = New SortedSet(Of String)(StringComparer.CurrentCultureIgnoreCase)

                ' Optionally, initialize the protein sequence hash objects
                If mSaveProteinSequenceHashInfoFiles Then
                    mCheckForDuplicateProteinSequences = True
                End If

                If mGenerateFixedFastaFile And mFixedFastaOptions.ConsolidateProteinsWithDuplicateSeqs Then
                    mCheckForDuplicateProteinSequences = True
                    mSaveProteinSequenceHashInfoFiles = Not usingPreloadedProteinNames
                    blnConsolidateDuplicateProteinSeqsInFasta = True
                    blnKeepDuplicateNamedProteinsUnlessMatchingSequence = mFixedFastaOptions.KeepDuplicateNamedProteinsUnlessMatchingSequence
                    blnConsolidateDupsIgnoreILDiff = mFixedFastaOptions.ConsolidateDupsIgnoreILDiff
                ElseIf mGenerateFixedFastaFile And mFixedFastaOptions.KeepDuplicateNamedProteinsUnlessMatchingSequence Then
                    mCheckForDuplicateProteinSequences = True
                    mSaveProteinSequenceHashInfoFiles = Not usingPreloadedProteinNames
                    blnConsolidateDuplicateProteinSeqsInFasta = False
                    blnKeepDuplicateNamedProteinsUnlessMatchingSequence = mFixedFastaOptions.KeepDuplicateNamedProteinsUnlessMatchingSequence
                    blnConsolidateDupsIgnoreILDiff = mFixedFastaOptions.ConsolidateDupsIgnoreILDiff
                End If

                If mCheckForDuplicateProteinSequences Then
                    lstProteinSequenceHashes.Clear()
                    intProteinSequenceHashCount = 0
                    ReDim oProteinSeqHashInfo(99)
                End If

                ' Parse each line in the file
                Dim lngBytesRead As Int64 = 0
                Dim dtLastMemoryUsageReport = DateTime.UtcNow

                ' Note: This value is updated only if the line length is < mMaximumResiduesPerLine
                Dim intCurrentValidResidueLineLengthMax = 0
                Dim blnProcessingDuplicateOrInvalidProtein = False

                Do While Not srFastaInFile.EndOfStream

                    Dim strLineIn = srFastaInFile.ReadLine()
                    lngBytesRead += strLineIn.Length + intTerminatorSize

                    If mLineCount Mod 50 = 0 Then
                        If MyBase.AbortProcessing Then Exit Do

                        Dim sngPercentComplete = CType(lngBytesRead / CType(srFastaInFile.BaseStream.Length, Single) * 100.0, Single)
                        If blnConsolidateDuplicateProteinSeqsInFasta OrElse blnKeepDuplicateNamedProteinsUnlessMatchingSequence Then
                            ' Bump the % complete down so that 100% complete in this routine will equate to 75% complete
                            ' The remaining 25% will occur in ConsolidateDuplicateProteinSeqsInFasta
                            sngPercentComplete = sngPercentComplete * 3 / 4
                        End If

                        MyBase.UpdateProgress("Validating FASTA File (" & Math.Round(sngPercentComplete, 0) & "% Done)", sngPercentComplete)

                        If DateTime.UtcNow.Subtract(dtLastMemoryUsageReport).TotalMinutes >= 1 Then
                            dtLastMemoryUsageReport = DateTime.UtcNow
                            ReportMemoryUsage(lstPreloadedProteinNamesToKeep, lstProteinSequenceHashes, lstProteinNames, oProteinSeqHashInfo)
                        End If
                    End If

                    mLineCount += 1

                    If strLineIn Is Nothing Then Continue Do

                    If strLineIn.Trim.Length = 0 Then
                        ' We typically only want blank lines at the end of the fasta file or between two protein entries
                        blnBlankLineProcessed = True
                        Continue Do
                    End If

                    If strLineIn.Chars(0) = " "c Then
                        If mWarnLineStartsWithSpace Then
                            RecordFastaFileError(mLineCount, 0, String.Empty,
                                                 eMessageCodeConstants.LineStartsWithSpace, String.Empty, ExtractContext(strLineIn, 0))
                        End If
                    End If

                    ' Note: Only trim the start of the line; do not trim the end of the line since Sequest incorrectly notates the peptide terminal state if a residue has a space after it
                    strLineIn = strLineIn.TrimStart

                    If strLineIn.Chars(0) = mProteinLineStartChar Then
                        ' Protein entry

                        If sbCurrentResidues.Length > 0 Then
                            ProcessResiduesForPreviousProtein(
                                strProteinName, sbCurrentResidues,
                                lstProteinSequenceHashes,
                                intProteinSequenceHashCount, oProteinSeqHashInfo,
                                blnConsolidateDupsIgnoreILDiff,
                                swFixedFastaOut, intCurrentValidResidueLineLengthMax,
                                swProteinSequenceHashBasic)

                            intCurrentValidResidueLineLengthMax = 0
                        End If

                        ' Now process this protein entry
                        mProteinCount += 1
                        blnProteinHeaderFound = True
                        blnProcessingResidueBlock = False
                        blnProcessingDuplicateOrInvalidProtein = False

                        strProteinName = String.Empty

                        AnalyzeFastaProcessProteinHeader(
                            swFixedFastaOut,
                            strLineIn,
                            strProteinName,
                            blnProcessingDuplicateOrInvalidProtein,
                            lstPreloadedProteinNamesToKeep,
                            lstProteinNames,
                            udtHeaderLineRuleDetails,
                            udtProteinNameRuleDetails,
                            udtProteinDescriptionRuleDetails,
                            reProteinNameTruncation)

                        If blnBlankLineProcessed Then
                            ' The previous line was blank; raise a warning
                            If mWarnBlankLinesBetweenProteins Then
                                RecordFastaFileWarning(mLineCount, 0, strProteinName, eMessageCodeConstants.BlankLineBeforeProteinName)
                            End If
                        End If

                    Else
                        ' Protein residues

                        If Not blnProcessingResidueBlock Then
                            If blnProteinHeaderFound Then
                                blnProteinHeaderFound = False

                                If blnBlankLineProcessed Then
                                    RecordFastaFileError(mLineCount, 0, strProteinName, eMessageCodeConstants.BlankLineBetweenProteinNameAndResidues)
                                End If
                            Else
                                RecordFastaFileError(mLineCount, 0, String.Empty, eMessageCodeConstants.ResiduesFoundWithoutProteinHeader)
                            End If

                            blnProcessingResidueBlock = True
                        Else
                            If blnBlankLineProcessed Then
                                RecordFastaFileError(mLineCount, 0, strProteinName, eMessageCodeConstants.BlankLineInMiddleOfResidues)
                            End If
                        End If

                        Dim intNewResidueCount = strLineIn.Length
                        mResidueCount += intNewResidueCount

                        ' Check the line length; raise a warning if longer than suggested
                        If intNewResidueCount > mMaximumResiduesPerLine Then
                            RecordFastaFileWarning(mLineCount, 0, strProteinName, eMessageCodeConstants.ResiduesLineTooLong, intNewResidueCount.ToString, String.Empty)
                        End If

                        ' Test the protein sequence rules
                        EvaluateRules(udtProteinSequenceRuleDetails, strProteinName, strLineIn, 0, strLineIn, 5)

                        If mGenerateFixedFastaFile OrElse mCheckForDuplicateProteinSequences OrElse mSaveBasicProteinHashInfoFile Then
                            Dim strResiduesClean As String

                            If mFixedFastaOptions.RemoveInvalidResidues Then
                                ' Auto-fix residues to remove any non-letter characters (spaces, asterisks, etc.)
                                strResiduesClean = reNonLetterResidues.Replace(strLineIn, String.Empty)
                            Else
                                ' Do not remove non-letter characters, but do remove leading or trailing whitespace
                                strResiduesClean = String.Copy(strLineIn.Trim())
                            End If

                            If Not swFixedFastaOut Is Nothing AndAlso Not blnProcessingDuplicateOrInvalidProtein Then
                                If strResiduesClean <> strLineIn Then
                                    mFixedFastaStats.UpdatedResidueLines += 1
                                End If

                                If Not mFixedFastaOptions.WrapLongResidueLines Then
                                    ' Only write out this line if not auto-wrapping long residue lines
                                    ' If we are auto-wrapping, then the residues will be written out by the call to ProcessResiduesForPreviousProtein
                                    swFixedFastaOut.WriteLine(strResiduesClean)
                                End If
                            End If

                            If mCheckForDuplicateProteinSequences OrElse mFixedFastaOptions.WrapLongResidueLines Then
                                ' Only add the residues if this is not a duplicate/invalid protein
                                If Not blnProcessingDuplicateOrInvalidProtein Then
                                    sbCurrentResidues.Append(strResiduesClean)
                                    If strResiduesClean.Length > intCurrentValidResidueLineLengthMax Then
                                        intCurrentValidResidueLineLengthMax = strResiduesClean.Length
                                    End If
                                End If
                            End If
                        End If

                        ' Reset the blank line tracking variable
                        blnBlankLineProcessed = False

                    End If

                Loop

                If sbCurrentResidues.Length > 0 Then
                    ProcessResiduesForPreviousProtein(
                       strProteinName, sbCurrentResidues,
                       lstProteinSequenceHashes,
                       intProteinSequenceHashCount, oProteinSeqHashInfo,
                       blnConsolidateDupsIgnoreILDiff,
                       swFixedFastaOut, intCurrentValidResidueLineLengthMax,
                       swProteinSequenceHashBasic)
                End If

                If mCheckForDuplicateProteinSequences Then
                    ' Step through oProteinSeqHashInfo and look for duplicate sequences
                    For intIndex = 0 To intProteinSequenceHashCount - 1
                        If oProteinSeqHashInfo(intIndex).AdditionalProteins.Count > 0 Then
                            With oProteinSeqHashInfo(intIndex)
                                RecordFastaFileWarning(mLineCount, 0, .ProteinNameFirst, eMessageCodeConstants.DuplicateProteinSequence,
                                  .ProteinNameFirst & ", " & FlattenArray(.AdditionalProteins, ","c), .SequenceStart)
                            End With
                        End If
                    Next intIndex
                End If

                Dim memoryUsageMB = clsMemoryUsageLogger.GetProcessMemoryUsageMB
                If memoryUsageMB > mProcessMemoryUsageMBAtStart * 4 OrElse
                   memoryUsageMB - mProcessMemoryUsageMBAtStart > 50 Then
                    ReportMemoryUsage(lstPreloadedProteinNamesToKeep, lstProteinSequenceHashes, lstProteinNames, oProteinSeqHashInfo)
                End If

            End Using

            ' Close the output files
            If Not swFixedFastaOut Is Nothing Then
                swFixedFastaOut.Close()
            End If

            If Not swProteinSequenceHashBasic Is Nothing Then
                swProteinSequenceHashBasic.Close()
            End If

            If mProteinCount = 0 Then
                RecordFastaFileError(mLineCount, 0, String.Empty, eMessageCodeConstants.ProteinEntriesNotFound)
            ElseIf blnProteinHeaderFound Then
                RecordFastaFileError(mLineCount, 0, strProteinName, eMessageCodeConstants.FinalProteinEntryMissingResidues)
            ElseIf Not blnBlankLineProcessed Then
                ' File does not end in multiple blank lines; need to re-open it using a binary reader and check the last two characters to make sure they're valid
                Threading.Thread.Sleep(100)

                If Not VerifyLinefeedAtEOF(strFastaFilePath, mAddMissingLinefeedAtEOF) Then
                    RecordFastaFileError(mLineCount, 0, String.Empty, eMessageCodeConstants.FileDoesNotEndWithLinefeed)
                End If
            End If

            If usingPreloadedProteinNames Then
                ' Report stats on the number of proteins read, the number written, and any that had duplicate protein names in the original fasta file
                Dim nameCountNotFound = 0
                Dim duplicateProteinNameCount = 0
                Dim proteinCountWritten = 0
                Dim preloadedProteinNameCount = 0

                For Each spanningKey In lstPreloadedProteinNamesToKeep.GetSpanningKeys
                    Dim proteinsForKey = lstPreloadedProteinNamesToKeep.GetListForSpanningKey(spanningKey)
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

                blnSuccess = True

            ElseIf mSaveProteinSequenceHashInfoFiles Then
                Dim sngPercentComplete = 98.0!
                If blnConsolidateDuplicateProteinSeqsInFasta OrElse blnKeepDuplicateNamedProteinsUnlessMatchingSequence Then
                    sngPercentComplete = sngPercentComplete * 3 / 4
                End If
                MyBase.UpdateProgress("Validating FASTA File (" & Math.Round(sngPercentComplete, 0) & "% Done)", sngPercentComplete)

                blnSuccess = AnalyzeFastaSaveHashInfo(
                  strFastaFilePath,
                  intProteinSequenceHashCount,
                  oProteinSeqHashInfo,
                  blnConsolidateDuplicateProteinSeqsInFasta,
                  blnConsolidateDupsIgnoreILDiff,
                  blnKeepDuplicateNamedProteinsUnlessMatchingSequence,
                  strFastaFilePathOut)
            Else
                blnSuccess = True
            End If

            If MyBase.AbortProcessing Then
                MyBase.UpdateProgress("Parsing aborted")
            Else
                MyBase.UpdateProgress("Parsing complete", 100)
            End If

        Catch ex As Exception
            If MyBase.ShowMessages Then
                ShowErrorMessage("Error in AnalyzeFastaFile:" & ex.Message)
                ShowExceptionStackTrace("AnalyzeFastaFile", ex)
            Else
                Throw New Exception("Error in AnalyzeFastaFile", ex)
            End If
            blnSuccess = False
        Finally
            ' These close statements will typically be redundant,
            ' However, if an exception occurs, then they will be needed to close the files

            If Not swFixedFastaOut Is Nothing Then
                swFixedFastaOut.Close()
            End If

            If Not swProteinSequenceHashBasic Is Nothing Then
                swProteinSequenceHashBasic.Close()
            End If

        End Try

        Return blnSuccess

    End Function

    Private Sub AnalyzeFastaProcessProteinHeader(
     swFixedFastaOut As TextWriter,
     strLineIn As String,
     <Out()> ByRef strProteinName As String,
     <Out()> ByRef blnProcessingDuplicateOrInvalidProtein As Boolean,
     lstPreloadedProteinNamesToKeep As clsNestedStringIntList,
     lstProteinNames As ISet(Of String),
     udtHeaderLineRuleDetails As IList(Of udtRuleDefinitionExtendedType),
     udtProteinNameRuleDetails As IList(Of udtRuleDefinitionExtendedType),
     udtProteinDescriptionRuleDetails As IList(Of udtRuleDefinitionExtendedType),
     reProteinNameTruncation As udtProteinNameTruncationRegex)

        Dim intDescriptionStartIndex As Integer

        Dim strProteinDescription As String = String.Empty

        Dim blnSkipDuplicateProtein = False

        strProteinName = String.Empty
        blnProcessingDuplicateOrInvalidProtein = True

        Try
            SplitFastaProteinHeaderLine(strLineIn, strProteinName, strProteinDescription, intDescriptionStartIndex)

            If strProteinName.Length = 0 Then
                blnProcessingDuplicateOrInvalidProtein = True
            Else
                blnProcessingDuplicateOrInvalidProtein = False
            End If


            If strProteinName = "contig_4_2" OrElse strProteinName = "contig_4_3" OrElse strProteinName = "contig_10083_1" Then
                Console.WriteLine("check this")
            End If


            ' Test the header line rules
            EvaluateRules(udtHeaderLineRuleDetails, strProteinName, strLineIn, 0, strLineIn, DEFAULT_CONTEXT_LENGTH)

            If strProteinDescription.Length > 0 Then
                ' Test the protein description rules

                EvaluateRules(udtProteinDescriptionRuleDetails, strProteinName, strProteinDescription,
                 intDescriptionStartIndex, strLineIn, DEFAULT_CONTEXT_LENGTH)
            End If

            If strProteinName.Length > 0 Then

                ' Check for protein names that are too long or too short
                If strProteinName.Length < mMinimumProteinNameLength Then
                    RecordFastaFileWarning(mLineCount, 1, strProteinName,
                     eMessageCodeConstants.ProteinNameIsTooShort, strProteinName.Length.ToString, String.Empty)
                ElseIf strProteinName.Length > mMaximumProteinNameLength Then
                    RecordFastaFileError(mLineCount, 1, strProteinName,
                     eMessageCodeConstants.ProteinNameIsTooLong, strProteinName.Length.ToString, String.Empty)
                End If

                ' Test the protein name rules
                EvaluateRules(udtProteinNameRuleDetails, strProteinName, strProteinName, 1, strLineIn, DEFAULT_CONTEXT_LENGTH)

                If Not lstPreloadedProteinNamesToKeep Is Nothing AndAlso lstPreloadedProteinNamesToKeep.Count > 0 Then
                    ' See if lstPreloadedProteinNamesToKeep contains strProteinName
                    Dim matchCount As Integer = lstPreloadedProteinNamesToKeep.GetValueForItem(strProteinName, -1)

                    If matchCount >= 0 Then
                        ' Name is known; increment the value for this protein

                        If matchCount = 0 Then
                            blnSkipDuplicateProtein = False
                        Else
                            ' An entry with this protein name has already been written
                            ' Do not include the duplicate
                            blnSkipDuplicateProtein = True
                        End If

                        If Not lstPreloadedProteinNamesToKeep.SetValueForItem(strProteinName, matchCount + 1) Then
                            ShowMessage("WARNING: protein " & strProteinName & " not found in lstPreloadedProteinNamesToKeep")
                        End If

                    Else
                        ' Unknown protein name; do not keep this protein
                        blnSkipDuplicateProtein = True
                        blnProcessingDuplicateOrInvalidProtein = True
                    End If

                    If mGenerateFixedFastaFile Then
                        ' Make sure strProteinDescription doesn't start with a | or space
                        If strProteinDescription.Length > 0 Then
                            strProteinDescription = strProteinDescription.TrimStart(New Char() {"|"c, " "c})
                        End If
                    End If

                Else

                    If mGenerateFixedFastaFile Then
                        strProteinName = AutoFixProteinNameAndDescription(strProteinName, strProteinDescription, reProteinNameTruncation)
                    End If

                    ' Optionally, check for duplicate protein names
                    If mCheckForDuplicateProteinNames Then
                        strProteinName = ExamineProteinName(strProteinName, lstProteinNames, blnSkipDuplicateProtein, blnProcessingDuplicateOrInvalidProtein)

                        If blnSkipDuplicateProtein Then
                            blnProcessingDuplicateOrInvalidProtein = True
                        End If
                    End If

                End If

                If Not swFixedFastaOut Is Nothing AndAlso Not blnSkipDuplicateProtein Then
                    swFixedFastaOut.WriteLine(ConstructFastaHeaderLine(strProteinName.Trim, strProteinDescription.Trim))
                End If
            End If


        Catch ex As Exception
            RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError,
             "Error parsing protein header line '" & strLineIn & "': " & ex.Message, String.Empty)
            ShowMessage(ex.Message)
            ShowExceptionStackTrace("AnalyzeFastaFileProcesssProteinHeader", ex)
        End Try


    End Sub

    Private Function AnalyzeFastaSaveHashInfo(
      strFastaFilePath As String,
      intProteinSequenceHashCount As Integer,
      oProteinSeqHashInfo As IList(Of clsProteinHashInfo),
      blnConsolidateDuplicateProteinSeqsInFasta As Boolean,
      blnConsolidateDupsIgnoreILDiff As Boolean,
      blnKeepDuplicateNamedProteinsUnlessMatchingSequence As Boolean,
      strFastaFilePathOut As String) As Boolean

        Dim swUniqueProteinSeqsOut As StreamWriter
        Dim swDuplicateProteinMapping As StreamWriter = Nothing

        Dim strUniqueProteinSeqsFileOut As String = String.Empty
        Dim strDuplicateProteinMappingFileOut As String = String.Empty

        Dim intIndex As Integer
        Dim intDuplicateIndex As Integer

        Dim blnDuplicateProteinSeqsFound As Boolean
        Dim blnSuccess As Boolean

        blnDuplicateProteinSeqsFound = False

        Try
            strUniqueProteinSeqsFileOut =
             Path.Combine(Path.GetDirectoryName(strFastaFilePath),
             Path.GetFileNameWithoutExtension(strFastaFilePath) & "_UniqueProteinSeqs.txt")

            ' Create swUniqueProteinSeqsOut
            swUniqueProteinSeqsOut = New StreamWriter(New FileStream(strUniqueProteinSeqsFileOut, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
        Catch ex As Exception
            ' Error opening output file
            RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError,
             "Error creating output file " & strUniqueProteinSeqsFileOut & ": " & ex.Message, String.Empty)
            ShowMessage(ex.Message)
            ShowExceptionStackTrace("AnalyzeFastaSaveHashInfo (to _UniqueProteinSeqs.txt)", ex)
            Return False
        End Try

        Try
            ' Define the path to the protein mapping file, but don't create it yet; just delete it if it exists
            ' We'll only create it if two or more proteins have the same protein sequence
            strDuplicateProteinMappingFileOut =
              Path.Combine(Path.GetDirectoryName(strFastaFilePath),
              Path.GetFileNameWithoutExtension(strFastaFilePath) & "_UniqueProteinSeqDuplicates.txt")                       ' Look for strDuplicateProteinMappingFileOut and erase it if it exists

            If File.Exists(strDuplicateProteinMappingFileOut) Then
                File.Delete(strDuplicateProteinMappingFileOut)
            End If
        Catch ex As Exception
            ' Error deleting output file
            RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError,
             "Error deleting output file " & strDuplicateProteinMappingFileOut & ": " & ex.Message, String.Empty)
            ShowMessage(ex.Message)
            ShowExceptionStackTrace("AnalyzeFastaSaveHashInfo (to _UniqueProteinSeqDuplicates.txt)", ex)
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

            For intIndex = 0 To intProteinSequenceHashCount - 1
                With oProteinSeqHashInfo(intIndex)

                    Dim dataValues = New List(Of String) From {
                        (intIndex + 1).ToString,
                        .ProteinNameFirst,
                        .SequenceLength.ToString(),
                        .SequenceHash,
                        (.AdditionalProteins.Count + 1).ToString(),
                        FlattenArray(.AdditionalProteins, ","c)}

                    swUniqueProteinSeqsOut.WriteLine(FlattenList(dataValues))

                    If .AdditionalProteins.Count > 0 Then
                        blnDuplicateProteinSeqsFound = True

                        If swDuplicateProteinMapping Is Nothing Then
                            ' Need to create swDuplicateProteinMapping
                            swDuplicateProteinMapping = New StreamWriter(New FileStream(strDuplicateProteinMappingFileOut, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))

                            Dim proteinHeaderColumns = New List(Of String) From {
                                "Sequence_Index",
                                "Protein_Name_First",
                                SEQUENCE_LENGTH_COLUMN,
                                "Duplicate_Protein"}

                            swDuplicateProteinMapping.WriteLine(FlattenList(proteinHeaderColumns))
                        End If

                        For Each strAdditionalProtein As String In .AdditionalProteins
                            If Not .AdditionalProteins(intDuplicateIndex) Is Nothing Then
                                If strAdditionalProtein.Trim.Length > 0 Then
                                    Dim proteinDataValues = New List(Of String) From {
                                        (intIndex + 1).ToString,
                                        .ProteinNameFirst,
                                        .SequenceLength.ToString(),
                                        strAdditionalProtein}

                                    swDuplicateProteinMapping.WriteLine(FlattenList(proteinDataValues))
                                End If
                            End If
                        Next

                    ElseIf .DuplicateProteinNameCount > 0 Then
                        blnDuplicateProteinSeqsFound = True
                    End If
                End With

            Next intIndex

            swUniqueProteinSeqsOut.Close()
            If Not swDuplicateProteinMapping Is Nothing Then swDuplicateProteinMapping.Close()

            blnSuccess = True

        Catch ex As Exception
            ' Error writing results
            RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError,
             "Error writing results to " & strUniqueProteinSeqsFileOut & " or " & strDuplicateProteinMappingFileOut & ": " & ex.Message, String.Empty)
            ShowMessage(ex.Message)
            ShowExceptionStackTrace("AnalyzeFastaSaveHashInfo", ex)
            blnSuccess = False
        End Try

        If blnSuccess And intProteinSequenceHashCount > 0 And blnDuplicateProteinSeqsFound Then
            If blnConsolidateDuplicateProteinSeqsInFasta OrElse blnKeepDuplicateNamedProteinsUnlessMatchingSequence Then
                blnSuccess = CorrectForDuplicateProteinSeqsInFasta(blnConsolidateDuplicateProteinSeqsInFasta, blnConsolidateDupsIgnoreILDiff, strFastaFilePathOut, intProteinSequenceHashCount, oProteinSeqHashInfo)
            End If
        End If

        Return blnSuccess

    End Function

    ''' <summary>
    ''' Pre-scan a portion of the fasta file to determine the appropriate value for mProteinNameSpannerCharLength
    ''' </summary>
    ''' <param name="fastaFilePathToTest">Fasta file to examine</param>
    ''' <param name="intTerminatorSize">Linefeed length (1 for LF or 2 for CRLF)</param>
    ''' <remarks>
    ''' Reads 50 MB chunks from 10 sections of the Fasta file (or the entire Fasta file if under 500 MB in size)
    ''' Keeps track of the portion of protein names in common between adjacent proteins
    ''' Uses this information to determine an appropriate value for mProteinNameSpannerCharLength
    ''' </remarks>
    Private Sub AutoDetermineFastaProteinNameSpannerCharLength(fastaFilePathToTest As String, intTerminatorSize As Integer)

        Const PARTS_TO_SAMPLE = 10
        Const KBYTES_PER_SAMPLE = 51200

        Dim proteinStartLetters = New Dictionary(Of String, Integer)
        Dim dtStartTime = DateTime.UtcNow
        Dim showStats = False

        Dim fiFastaFile = New FileInfo(fastaFilePathToTest)
        If Not fiFastaFile.Exists Then Return

        Dim fullScanLengthBytes = 1024L * PARTS_TO_SAMPLE * KBYTES_PER_SAMPLE
        Dim linesReadTotal As Int64

        If fiFastaFile.Length < fullScanLengthBytes Then
            fullScanLengthBytes = fiFastaFile.Length
            linesReadTotal = AutoDetermineFastaProteinNameSpannerCharLength(fiFastaFile, intTerminatorSize, proteinStartLetters, 0, fiFastaFile.Length)
        Else

            Dim stepSizeBytes = CLng(Math.Round(fiFastaFile.Length / PARTS_TO_SAMPLE, 0))

            For byteOffsetStart As Int64 = 0 To fiFastaFile.Length Step stepSizeBytes
                Dim linesRead = AutoDetermineFastaProteinNameSpannerCharLength(fiFastaFile, intTerminatorSize, proteinStartLetters, byteOffsetStart, KBYTES_PER_SAMPLE * 1024)
                linesReadTotal += linesRead

                If Not showStats AndAlso DateTime.UtcNow.Subtract(dtStartTime).TotalMilliseconds > 500 Then
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
                Dim percentFileProcessed = fullScanLengthBytes / fiFastaFile.Length * 100
                Console.WriteLine("  parsed {0:0}% of the file, reading {1:#,##0} lines and finding {2:#,##0} proteins",
                                  percentFileProcessed, linesReadTotal, preScanProteinCount)
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
    ''' <param name="fiFastaFile"></param>
    ''' <param name="intTerminatorSize"></param>
    ''' <param name="proteinStartLetters"></param>
    ''' <param name="startOffset"></param>
    ''' <param name="bytesToRead"></param>
    ''' <returns>The number of lines read</returns>
    ''' <remarks></remarks>
    Private Function AutoDetermineFastaProteinNameSpannerCharLength(
      fiFastaFile As FileInfo,
      intTerminatorSize As Integer,
      proteinStartLetters As IDictionary(Of String, Integer),
      startOffset As Int64,
      bytesToRead As Int64) As Int64

        Dim linesRead = 0L

        Try
            Dim previousProteinLength = 0
            Dim previousProteinName = String.Empty

            If startOffset >= fiFastaFile.Length Then
                ShowMessage("Ignoring byte offset of " & startOffset &
                            " in AutoDetermineProteinNameSpannerCharLength since past the end of the file " &
                            "(" & fiFastaFile.Length & " bytes")
                Return 0
            End If

            Dim bytesRead As Int64 = 0

            Dim firstLineDiscarded As Boolean
            If startOffset = 0 Then
                firstLineDiscarded = True
            End If

            Using fsInstream = New FileStream(fiFastaFile.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                fsInstream.Position = startOffset

                Using srFastaInFile = New StreamReader(fsInstream)

                    Do While Not srFastaInFile.EndOfStream

                        Dim strLineIn = srFastaInFile.ReadLine()
                        bytesRead += intTerminatorSize
                        linesRead += 1

                        If String.IsNullOrEmpty(strLineIn) Then
                            Continue Do
                        End If

                        bytesRead += strLineIn.Length
                        If Not firstLineDiscarded Then
                            ' We can't trust that this was a full line of text; skip it
                            firstLineDiscarded = True
                            Continue Do
                        End If

                        If bytesRead > bytesToRead Then
                            Exit Do
                        End If

                        If Not strLineIn.Chars(0) = mProteinLineStartChar Then
                            Continue Do
                        End If

                        ' Make sure the protein name and description are valid
                        ' Find the first space and/or tab
                        Dim intSpaceIndex = GetBestSpaceIndex(strLineIn)
                        Dim strProteinName As String

                        If intSpaceIndex > 1 Then
                            strProteinName = strLineIn.Substring(1, intSpaceIndex - 1)
                        Else
                            ' Line does not contain a description
                            If intSpaceIndex <= 0 Then
                                If strLineIn.Trim.Length <= 1 Then
                                    Continue Do
                                Else
                                    ' The line contains a protein name, but not a description
                                    strProteinName = strLineIn.Substring(1)
                                End If
                            Else
                                ' Space or tab found directly after the > symbol
                                Continue Do
                            End If
                        End If

                        If previousProteinLength = 0 Then
                            previousProteinName = String.Copy(strProteinName)
                            previousProteinLength = previousProteinName.Length
                            Continue Do
                        End If

                        Dim currentNameLength = strProteinName.Length
                        Dim charIndex = 0

                        While charIndex < previousProteinLength
                            If charIndex >= currentNameLength Then
                                Exit While
                            End If

                            If previousProteinName(charIndex) <> strProteinName.Chars(charIndex) Then
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

                        previousProteinName = String.Copy(strProteinName)
                        previousProteinLength = previousProteinName.Length
                    Loop

                End Using

            End Using

        Catch ex As Exception
            If MyBase.ShowMessages Then
                ShowErrorMessage("Error in AutoDetermineProteinNameSpannerCharLength:" & ex.Message)
            Else
                Throw New Exception("Error in AutoDetermineProteinNameSpannerCharLength", ex)
            End If
        End Try

        Return linesRead

    End Function

    Private Function AutoFixProteinNameAndDescription(
       ByRef strProteinName As String,
       ByRef strProteinDescription As String,
       reProteinNameTruncation As udtProteinNameTruncationRegex) As String

        Dim blnProteinNameTooLong As Boolean
        Dim reMatch As Match
        Dim strNewProteinName As String
        Dim intCharIndex As Integer
        Dim intMinCharIndex As Integer
        Dim strExtraProteinNameText As String
        Dim chInvalidChar As Char

        Dim blnMultipleRefsSplitOutFromKnownAccession = False

        ' Auto-fix potential errors in the protein name

        ' Possibly truncate to mMaximumProteinNameLength characters
        If strProteinName.Length > mMaximumProteinNameLength Then
            blnProteinNameTooLong = True
        Else
            blnProteinNameTooLong = False
        End If

        If mFixedFastaOptions.SplitOutMultipleRefsForKnownAccession OrElse
           (mFixedFastaOptions.TruncateLongProteinNames And blnProteinNameTooLong) Then

            ' First see if the name fits the pattern IPI:IPI00048500.11|
            ' Next see if the name fits the pattern gi|7110699|
            ' Next see if the name fits the pattern jgi
            ' Next see if the name fits the generic pattern defined by reProteinNameTruncation.reMatchGeneric
            ' Otherwise, use mFixedFastaOptions.LongProteinNameSplitChars to define where to truncate

            strNewProteinName = String.Copy(strProteinName)
            strExtraProteinNameText = String.Empty

            reMatch = reProteinNameTruncation.reMatchIPI.Match(strProteinName)
            If reMatch.Success Then
                blnMultipleRefsSplitOutFromKnownAccession = True
            Else
                ' IPI didn't match; try gi
                reMatch = reProteinNameTruncation.reMatchGI.Match(strProteinName)
            End If

            If reMatch.Success Then
                blnMultipleRefsSplitOutFromKnownAccession = True
            Else
                ' GI didn't match; try jgi
                reMatch = reProteinNameTruncation.reMatchJGI.Match(strProteinName)
            End If

            If reMatch.Success Then
                blnMultipleRefsSplitOutFromKnownAccession = True
            Else
                ' jgi didn't match; try generic (text separated by a series of colons or bars),
                '  but only if the name is too long
                If mFixedFastaOptions.TruncateLongProteinNames And blnProteinNameTooLong Then
                    reMatch = reProteinNameTruncation.reMatchGeneric.Match(strProteinName)
                End If
            End If

            If reMatch.Success Then
                ' Trunctate the protein name
                ' Truncate the protein name, but move the truncated portion into the next group
                strNewProteinName = reMatch.Groups(1).Value
                strExtraProteinNameText = reMatch.Groups(2).Value

            ElseIf mFixedFastaOptions.TruncateLongProteinNames And blnProteinNameTooLong Then

                ' Name is too long, but it didn't match the known patterns
                ' Find the last occurrence of mFixedFastaOptions.LongProteinNameSplitChars (default is vertical bar)
                '   and truncate the text following the match
                ' Repeat the process until the protein name length >= mMaximumProteinNameLength

                ' See if any of the characters in chProteinNameSplitChars is present after
                ' character 6 but less than character mMaximumProteinNameLength
                intMinCharIndex = 6

                Do
                    intCharIndex = strNewProteinName.LastIndexOfAny(mFixedFastaOptions.LongProteinNameSplitChars)
                    If intCharIndex >= intMinCharIndex Then
                        If strExtraProteinNameText.Length > 0 Then
                            strExtraProteinNameText = "|" & strExtraProteinNameText
                        End If
                        strExtraProteinNameText = strNewProteinName.Substring(intCharIndex + 1) & strExtraProteinNameText
                        strNewProteinName = strNewProteinName.Substring(0, intCharIndex)
                    Else
                        intCharIndex = -1
                    End If

                Loop While intCharIndex > 0 And strNewProteinName.Length > mMaximumProteinNameLength

            End If

            If strExtraProteinNameText.Length > 0 Then
                If blnProteinNameTooLong Then
                    mFixedFastaStats.TruncatedProteinNameCount += 1
                Else
                    mFixedFastaStats.ProteinNamesMultipleRefsRemoved += 1
                End If

                strProteinName = String.Copy(strNewProteinName)

                PrependExtraTextToProteinDescription(strExtraProteinNameText, strProteinDescription)
            End If

        End If

        If mFixedFastaOptions.ProteinNameInvalidCharsToRemove.Length > 0 Then
            strNewProteinName = String.Copy(strProteinName)

            ' First remove invalid characters from the beginning or end of the protein name
            strNewProteinName = strNewProteinName.Trim(mFixedFastaOptions.ProteinNameInvalidCharsToRemove)

            If strNewProteinName.Length >= 1 Then
                For Each chInvalidChar In mFixedFastaOptions.ProteinNameInvalidCharsToRemove
                    ' Next, replace any remaining instances of the character with an underscore
                    strNewProteinName = strNewProteinName.Replace(chInvalidChar, INVALID_PROTEIN_NAME_CHAR_REPLACEMENT)
                Next

                If strProteinName <> strNewProteinName Then
                    If strNewProteinName.Length >= 3 Then
                        strProteinName = String.Copy(strNewProteinName)
                        mFixedFastaStats.ProteinNamesInvalidCharsReplaced += 1
                    End If
                End If
            End If
        End If

        If mFixedFastaOptions.SplitOutMultipleRefsInProteinName AndAlso Not blnMultipleRefsSplitOutFromKnownAccession Then
            ' Look for multiple refs in the protein name, but only if we didn't already split out multiple refs above

            reMatch = reProteinNameTruncation.reMatchDoubleBarOrColonAndBar.Match(strProteinName)
            If reMatch.Success Then
                ' Protein name contains 2 or more vertical bars, or a colon and a bar
                ' Split out the multiple refs and place them in the description
                ' However, jgi names are supposed to have two vertical bars, so we need to treat that data differently

                strExtraProteinNameText = String.Empty

                reMatch = reProteinNameTruncation.reMatchJGIBaseAndID.Match(strProteinName)
                If reMatch.Success Then
                    ' ProteinName is similar to jgi|Organism|00000
                    ' Check whether there is any text following the match
                    If reMatch.Length < strProteinName.Length Then
                        ' Extra text exists; populate strExtraProteinNameText
                        strExtraProteinNameText = strProteinName.Substring(reMatch.Length + 1)
                        strProteinName = reMatch.ToString
                    End If
                Else
                    ' Find the first vertical bar or colon
                    intCharIndex = strProteinName.IndexOfAny(mProteinNameFirstRefSepChars)

                    If intCharIndex > 0 Then
                        ' Find the second vertical bar, colon, or semicolon
                        intCharIndex = strProteinName.IndexOfAny(mProteinNameSubsequentRefSepChars, intCharIndex + 1)

                        If intCharIndex > 0 Then
                            ' Split the protein name
                            strExtraProteinNameText = strProteinName.Substring(intCharIndex + 1)
                            strProteinName = strProteinName.Substring(0, intCharIndex)
                        End If
                    End If

                End If

                If strExtraProteinNameText.Length > 0 Then
                    PrependExtraTextToProteinDescription(strExtraProteinNameText, strProteinDescription)
                    mFixedFastaStats.ProteinNamesMultipleRefsRemoved += 1
                End If

            End If
        End If

        ' Make sure strProteinDescription doesn't start with a | or space
        If strProteinDescription.Length > 0 Then
            strProteinDescription = strProteinDescription.TrimStart(New Char() {"|"c, " "c})
        End If

        Return strProteinName

    End Function

    Private Shadows Sub AbortProcessingNow() Implements IValidateFastaFile.AbortProcessingNow
        MyBase.AbortProcessingNow()
    End Sub

    Private Function BoolToStringInt(value As Boolean) As String
        If value Then
            Return "1"
        Else
            Return "0"
        End If
    End Function

    Private Function CharArrayToString(chCharArray() As Char) As String
        Return CStr(chCharArray)
    End Function


    Private Sub ClearAllRules() Implements IValidateFastaFile.ClearAllRules
        Me.ClearRules(IValidateFastaFile.RuleTypes.HeaderLine)
        Me.ClearRules(IValidateFastaFile.RuleTypes.ProteinDescription)
        Me.ClearRules(IValidateFastaFile.RuleTypes.ProteinName)
        Me.ClearRules(IValidateFastaFile.RuleTypes.ProteinSequence)

        mMasterCustomRuleID = CUSTOM_RULE_ID_START
    End Sub

    Private Sub ClearRules(ruleType As IValidateFastaFile.RuleTypes) Implements IValidateFastaFile.ClearRules
        Select Case ruleType
            Case IValidateFastaFile.RuleTypes.HeaderLine
                Me.ClearRulesDataStructure(mHeaderLineRules)
            Case IValidateFastaFile.RuleTypes.ProteinDescription
                Me.ClearRulesDataStructure(mProteinDescriptionRules)
            Case IValidateFastaFile.RuleTypes.ProteinName
                Me.ClearRulesDataStructure(mProteinNameRules)
            Case IValidateFastaFile.RuleTypes.ProteinSequence
                Me.ClearRulesDataStructure(mProteinSequenceRules)
        End Select
    End Sub

    Private Sub ClearRulesDataStructure(ByRef udtRules() As udtRuleDefinitionType)
        ReDim udtRules(-1)
    End Sub

    Public Function ComputeProteinHash(sbResidues As StringBuilder, blnConsolidateDupsIgnoreILDiff As Boolean) As String

        Static objHashGenerator As clsHashGenerator

        If objHashGenerator Is Nothing Then
            objHashGenerator = New clsHashGenerator
        End If

        If sbResidues.Length > 0 Then
            ' Compute the hash value for sbCurrentResidues
            If blnConsolidateDupsIgnoreILDiff Then
                Return objHashGenerator.GenerateHash(sbResidues.ToString.Replace("L"c, "I"c))
            Else
                Return objHashGenerator.GenerateHash(sbResidues.ToString)
            End If
        Else
            Return String.Empty
        End If

    End Function

    Private Function ComputeTotalSpecifiedCount(udtErrorStats As udtItemSummaryIndexedType) As Integer
        Dim intTotal As Integer
        Dim intIndex As Integer

        intTotal = 0
        For intIndex = 0 To udtErrorStats.ErrorStatsCount - 1
            intTotal += udtErrorStats.ErrorStats(intIndex).CountSpecified
        Next intIndex

        Return intTotal

    End Function

    Private Function ComputeTotalUnspecifiedCount(udtErrorStats As udtItemSummaryIndexedType) As Integer
        Dim intTotal As Integer
        Dim intIndex As Integer

        intTotal = 0
        For intIndex = 0 To udtErrorStats.ErrorStatsCount - 1
            intTotal += udtErrorStats.ErrorStats(intIndex).CountUnspecified
        Next intIndex

        Return intTotal

    End Function

    ''' <summary>
    ''' Looks for duplicate proteins in the Fasta file
    ''' Creates a new fasta file that has exact duplicates removed
    ''' Will consolidate proteins with the same sequence if blnConsolidateDuplicateProteinSeqsInFasta=True
    ''' </summary>
    ''' <param name="blnConsolidateDuplicateProteinSeqsInFasta"></param>
    ''' <param name="strFixedFastaFilePath"></param>
    ''' <param name="intProteinSequenceHashCount"></param>
    ''' <param name="oProteinSeqHashInfo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CorrectForDuplicateProteinSeqsInFasta(
      blnConsolidateDuplicateProteinSeqsInFasta As Boolean,
      blnConsolidateDupsIgnoreILDiff As Boolean,
      strFixedFastaFilePath As String,
      intProteinSequenceHashCount As Integer,
      oProteinSeqHashInfo As IList(Of clsProteinHashInfo)) As Boolean

        Dim fsInFile As Stream
        Dim swConsolidatedFastaOut As StreamWriter = Nothing

        Dim lngBytesRead As Int64
        Dim intTerminatorSize As Integer
        Dim sngPercentComplete As Single
        Dim intLineCount As Integer

        Dim strFixedFastaFilePathTemp As String = String.Empty
        Dim strLineIn As String

        Dim strCachedProteinName As String = String.Empty
        Dim strCachedProteinDescription As String = String.Empty
        Dim sbCachedProteinResidueLines = New StringBuilder(250)
        Dim sbCachedProteinResidues = New StringBuilder(250)

        ' This list contains the protein names that we will keep; values are the index values pointing into oProteinSeqHashInfo
        ' If blnConsolidateDuplicateProteinSeqsInFasta=False, this will contain all protein names
        ' If blnConsolidateDuplicateProteinSeqsInFasta=True, we only keep the first name found for a given sequence
        Dim lstProteinNameFirst As clsNestedStringDictionary(Of Integer)

        ' This list keeps track of the protein names that have been written out to the new fasta file
        ' Keys are the protein names; values are the index of the entry in oProteinSeqHashInfo()
        Dim lstProteinsWritten As clsNestedStringDictionary(Of Integer)

        ' This list contains the names of duplicate proteins; the hash values are the protein names of the master protein that has the same sequence
        Dim lstDuplicateProteinList As clsNestedStringDictionary(Of String)

        Dim intDescriptionStartIndex As Integer

        Dim blnSuccess As Boolean

        If intProteinSequenceHashCount <= 0 Then
            Return True
        End If

        ''''''''''''''''''''''
        ' Processing Steps
        ''''''''''''''''''''''
        '
        ' Open strFixedFastaFilePath with the fasta file reader
        ' Create a new fasta file with a writer

        ' For each protein, check whether it has duplicates
        ' If not, just write it out to the new fasta file

        ' If it does have duplicates and it is the master, then append the duplicate protein names to the end of the description for the protein
        '  and write out the name, new description, and sequence to the new fasta file

        ' Otherwise, check if it is a duplicate of a master protein
        ' If it is, then do not write the name, description, or sequence to the new fasta file

        Try
            strFixedFastaFilePathTemp = strFixedFastaFilePath & ".TempFixed"

            If File.Exists(strFixedFastaFilePathTemp) Then
                File.Delete(strFixedFastaFilePathTemp)
            End If

            File.Move(strFixedFastaFilePath, strFixedFastaFilePathTemp)
        Catch ex As Exception
            RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError,
             "Error renaming " & strFixedFastaFilePath & " to " & strFixedFastaFilePathTemp & ": " & ex.Message, String.Empty)
            ShowMessage(ex.Message)
            ShowExceptionStackTrace("CorrectForDuplicateProteinSeqsInFasta (rename fixed fasta to .tempfixed)", ex)
            Return False
        End Try

        Dim srFastaInFile As StreamReader

        Try
            ' Open the file and read, at most, the first 100,000 characters to see if it contains CrLf or just Lf
            intTerminatorSize = DetermineLineTerminatorSize(strFixedFastaFilePathTemp)

            ' Open the Fixed fasta file
            fsInFile = New FileStream(
               strFixedFastaFilePathTemp,
               FileMode.Open,
               FileAccess.Read,
               FileShare.ReadWrite)

            srFastaInFile = New StreamReader(fsInFile)

        Catch ex As Exception
            RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError,
             "Error opening " & strFixedFastaFilePathTemp & ": " & ex.Message, String.Empty)
            ShowMessage(ex.Message)
            ShowExceptionStackTrace("CorrectForDuplicateProteinSeqsInFasta (create strFixedFastafilePathTemp)", ex)
            Return False
        End Try

        Try
            ' Create the new fasta file
            swConsolidatedFastaOut = New StreamWriter(New FileStream(strFixedFastaFilePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
        Catch ex As Exception
            RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError,
             "Error creating consolidated fasta output file " & strFixedFastaFilePath & ": " & ex.Message, String.Empty)
            ShowMessage(ex.Message)
            ShowExceptionStackTrace("CorrectForDuplicateProteinSeqsInFasta (create strFixedFastaFilePath)", ex)
        End Try

        Try
            ' Populate lstProteinNameFirst with the protein names in oProteinSeqHashInfo().ProteinNameFirst
            lstProteinNameFirst = New clsNestedStringDictionary(Of Integer)(True, mProteinNameSpannerCharLength)

            ' Populate htDuplicateProteinList with the protein names in oProteinSeqHashInfo().AdditionalProteins
            lstDuplicateProteinList = New clsNestedStringDictionary(Of String)(True, mProteinNameSpannerCharLength)

            For intIndex As Integer = 0 To intProteinSequenceHashCount - 1
                With oProteinSeqHashInfo(intIndex)

                    If Not lstProteinNameFirst.ContainsKey(.ProteinNameFirst) Then
                        lstProteinNameFirst.Add(.ProteinNameFirst, intIndex)
                    Else
                        ' .ProteinNameFirst is already present in lstProteinNameFirst
                        ' The fixed fasta file will only actually contain the first occurrence of .ProteinNameFirst, so we can effectively ignore this entry
                        ' but we should increment the DuplicateNameSkipCount

                    End If

                    If .AdditionalProteins.Count > 0 Then
                        For Each strAdditionalProtein As String In .AdditionalProteins

                            If blnConsolidateDuplicateProteinSeqsInFasta Then
                                ' Update the duplicate protein name list
                                If Not lstDuplicateProteinList.ContainsKey(strAdditionalProtein) Then
                                    lstDuplicateProteinList.Add(strAdditionalProtein, .ProteinNameFirst)
                                End If

                            Else
                                ' We are not consolidating proteins with the same sequence but different protein names
                                ' Append this entry to lstProteinNameFirst

                                If Not lstProteinNameFirst.ContainsKey(strAdditionalProtein) Then
                                    lstProteinNameFirst.Add(strAdditionalProtein, intIndex)
                                Else
                                    ' .AdditionalProteins(intDupIndex) is already present in lstProteinNameFirst
                                    ' Increment the DuplicateNameSkipCount
                                End If

                            End If

                        Next
                    End If
                End With
            Next intIndex

            lstProteinsWritten = New clsNestedStringDictionary(Of Integer)(False, mProteinNameSpannerCharLength)

            Dim dtLastMemoryUsageReport = DateTime.UtcNow

            ' Parse each line in the file
            intLineCount = 0
            lngBytesRead = 0
            mFixedFastaStats.DuplicateSequenceProteinsSkipped = 0

            Do While Not srFastaInFile.EndOfStream
                strLineIn = srFastaInFile.ReadLine()
                lngBytesRead += strLineIn.Length + intTerminatorSize

                If intLineCount Mod 50 = 0 Then
                    If MyBase.AbortProcessing Then Exit Do

                    sngPercentComplete = 75 + CType(lngBytesRead / CType(srFastaInFile.BaseStream.Length, Single) * 100.0, Single) / 4
                    MyBase.UpdateProgress("Consolidating duplicate proteins to create a new FASTA File (" & Math.Round(sngPercentComplete, 0) & "% Done)", sngPercentComplete)

                    If DateTime.UtcNow.Subtract(dtLastMemoryUsageReport).TotalMinutes >= 1 Then
                        dtLastMemoryUsageReport = DateTime.UtcNow
                        ReportMemoryUsage(lstProteinNameFirst, lstProteinsWritten, lstDuplicateProteinList)
                    End If
                End If

                intLineCount += 1

                If Not strLineIn Is Nothing Then
                    If strLineIn.Trim.Length > 0 Then
                        ' Note: Trim the start of the line (however, since this is a fixed fasta file it should not start with a space)
                        strLineIn = strLineIn.TrimStart

                        If strLineIn.Chars(0) = mProteinLineStartChar Then
                            ' Protein entry line

                            If Not String.IsNullOrEmpty(strCachedProteinName) Then
                                ' Write out the cached protein and it's residues

                                WriteCachedProtein(strCachedProteinName, strCachedProteinDescription,
                                 swConsolidatedFastaOut, oProteinSeqHashInfo,
                                 sbCachedProteinResidueLines, sbCachedProteinResidues,
                                 blnConsolidateDuplicateProteinSeqsInFasta, blnConsolidateDupsIgnoreILDiff,
                                 lstProteinNameFirst, lstDuplicateProteinList,
                                 intLineCount, lstProteinsWritten)

                                strCachedProteinName = String.Empty
                                sbCachedProteinResidueLines.Length = 0
                                sbCachedProteinResidues.Length = 0
                            End If

                            ' Extract the protein name and description
                            SplitFastaProteinHeaderLine(strLineIn, strCachedProteinName, strCachedProteinDescription, intDescriptionStartIndex)

                        Else
                            ' Protein residues
                            sbCachedProteinResidueLines.AppendLine(strLineIn)
                            sbCachedProteinResidues.Append(strLineIn.Trim())
                        End If
                    End If
                End If

            Loop

            If Not String.IsNullOrEmpty(strCachedProteinName) Then
                ' Write out the cached protein and it's residues
                WriteCachedProtein(strCachedProteinName, strCachedProteinDescription,
                 swConsolidatedFastaOut, oProteinSeqHashInfo,
                 sbCachedProteinResidueLines, sbCachedProteinResidues,
                 blnConsolidateDuplicateProteinSeqsInFasta, blnConsolidateDupsIgnoreILDiff,
                 lstProteinNameFirst, lstDuplicateProteinList,
                 intLineCount, lstProteinsWritten)
            End If

            ReportMemoryUsage(lstProteinNameFirst, lstProteinsWritten, lstDuplicateProteinList)

            blnSuccess = True

        Catch ex As Exception
            RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError,
             "Error writing to consolidated fasta file " & strFixedFastaFilePath & ": " & ex.Message, String.Empty)
            ShowMessage(ex.Message)
            ShowExceptionStackTrace("CorrectForDuplicateProteinSeqsInFasta", ex)
            Return False
        Finally
            Try
                If Not srFastaInFile Is Nothing Then srFastaInFile.Close()
                If Not swConsolidatedFastaOut Is Nothing Then swConsolidatedFastaOut.Close()

                Threading.Thread.Sleep(100)

                File.Delete(strFixedFastaFilePathTemp)
            Catch ex As Exception
                ' Ignore errors here
                ShowMessage(ex.Message)
                ShowExceptionStackTrace("CorrectForDuplicateProteinSeqsInFasta (closing file handles)", ex)
            End Try
        End Try

        Return blnSuccess

    End Function

    Private Function ConstructFastaHeaderLine(ByRef strProteinName As String, strProteinDescription As String) As String

        If strProteinName Is Nothing Then strProteinName = "????"

        If String.IsNullOrWhiteSpace(strProteinDescription) Then
            Return mProteinLineStartChar & strProteinName
        Else
            Return mProteinLineStartChar & strProteinName & " " & strProteinDescription
        End If

    End Function

    Private Function ConstructStatsFilePath(strOutputFolderPath As String) As String

        Dim strStatsFilePath As String = String.Empty

        Try
            ' Record the current time in dtNow
            strStatsFilePath = "FastaFileStats_" & DateTime.Now.ToString("yyyy-MM-dd") & ".txt"

            If Not strOutputFolderPath Is Nothing AndAlso strOutputFolderPath.Length > 0 Then
                strStatsFilePath = Path.Combine(strOutputFolderPath, strStatsFilePath)
            End If
        Catch ex As Exception
            If strStatsFilePath Is Nothing OrElse strStatsFilePath.Length = 0 Then
                strStatsFilePath = "FastaFileStats.txt"
            End If
        End Try

        Return strStatsFilePath

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

    Private Function DetermineLineTerminatorSize(strInputFilePath As String) As Integer

        Dim endCharType As IValidateFastaFile.eLineEndingCharacters = Me.DetermineLineTerminatorType(strInputFilePath)

        Select Case endCharType
            Case IValidateFastaFile.eLineEndingCharacters.CR
                Return 1
            Case IValidateFastaFile.eLineEndingCharacters.LF
                Return 1
            Case IValidateFastaFile.eLineEndingCharacters.CRLF
                Return 2
            Case IValidateFastaFile.eLineEndingCharacters.LFCR
                Return 2
        End Select

        Return 2

    End Function

    Private Function DetermineLineTerminatorType(strInputFilePath As String) As IValidateFastaFile.eLineEndingCharacters
        Dim intByte As Integer

        Dim endCharacterType As IValidateFastaFile.eLineEndingCharacters

        Try
            ' Open the input file and look for the first carriage return or line feed
            Using fsInFile = New FileStream(strInputFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)

                Do While fsInFile.Position < fsInFile.Length AndAlso fsInFile.Position < 100000

                    intByte = fsInFile.ReadByte()

                    If intByte = 10 Then
                        ' Found linefeed
                        If fsInFile.Position < fsInFile.Length Then
                            intByte = fsInFile.ReadByte()
                            If intByte = 13 Then
                                ' LfCr
                                endCharacterType = IValidateFastaFile.eLineEndingCharacters.LFCR
                            Else
                                ' Lf only
                                endCharacterType = IValidateFastaFile.eLineEndingCharacters.LF
                            End If
                        Else
                            endCharacterType = IValidateFastaFile.eLineEndingCharacters.LF
                        End If
                        Exit Do
                    ElseIf intByte = 13 Then
                        ' Found carriage return
                        If fsInFile.Position < fsInFile.Length Then
                            intByte = fsInFile.ReadByte()
                            If intByte = 10 Then
                                ' CrLf
                                endCharacterType = IValidateFastaFile.eLineEndingCharacters.CRLF
                            Else
                                ' Cr only
                                endCharacterType = IValidateFastaFile.eLineEndingCharacters.CR
                            End If
                        Else
                            endCharacterType = IValidateFastaFile.eLineEndingCharacters.CR
                        End If
                        Exit Do
                    End If

                Loop

            End Using

        Catch ex As Exception
            SetLocalErrorCode(IValidateFastaFile.eValidateFastaFileErrorCodes.ErrorVerifyingLinefeedAtEOF)
        End Try

        Return endCharacterType

    End Function

    ' Unused function
    ''Private Function ExtractListItem(strList As String, intItem As Integer) As String
    ''    Dim strItems() As String
    ''    Dim strItem As String

    ''    strItem = String.Empty
    ''    If intItem >= 1 And Not strList Is Nothing Then
    ''        strItems = strList.Split(","c)
    ''        If strItems.Length >= intItem Then
    ''            strItem = strItems(intItem - 1)
    ''        End If
    ''    End If

    ''    Return strItem

    ''End Function

    Private Function NormalizeFileLineEndings(
     pathOfFileToFix As String,
     newFileName As String,
     desiredLineEndCharacterType As IValidateFastaFile.eLineEndingCharacters) As String

        Dim newEndChar As String = ControlChars.CrLf

        Dim endCharType As IValidateFastaFile.eLineEndingCharacters =
         Me.DetermineLineTerminatorType(pathOfFileToFix)

        Dim fi As FileInfo
        Dim tr As TextReader
        Dim s As String
        Dim origEndCharCount As Integer

        Dim fileLength As Long
        Dim currFilePos As Long

        Dim sw As StreamWriter

        If endCharType <> desiredLineEndCharacterType Then
            Select Case desiredLineEndCharacterType
                Case IValidateFastaFile.eLineEndingCharacters.CRLF
                    newEndChar = ControlChars.CrLf
                Case IValidateFastaFile.eLineEndingCharacters.CR
                    newEndChar = ControlChars.Cr
                Case IValidateFastaFile.eLineEndingCharacters.LF
                    newEndChar = ControlChars.Lf
                Case IValidateFastaFile.eLineEndingCharacters.LFCR
                    newEndChar = ControlChars.CrLf
            End Select

            Select Case endCharType
                Case IValidateFastaFile.eLineEndingCharacters.CR
                    origEndCharCount = 2
                Case IValidateFastaFile.eLineEndingCharacters.CRLF
                    origEndCharCount = 1
                Case IValidateFastaFile.eLineEndingCharacters.LF
                    origEndCharCount = 1
                Case IValidateFastaFile.eLineEndingCharacters.LFCR
                    origEndCharCount = 2
            End Select

            If Not Path.IsPathRooted(newFileName) Then
                newFileName = Path.Combine(Path.GetDirectoryName(pathOfFileToFix), Path.GetFileName(newFileName))
            End If

            fi = New FileInfo(pathOfFileToFix)
            fileLength = fi.Length

            tr = fi.OpenText

            sw = New StreamWriter(New FileStream(newFileName, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))

            Me.OnProgressUpdate("Normalizing Line Endings...", 0.0)

            s = tr.ReadLine
            Dim linesRead As Long = 0
            Do While Not s Is Nothing
                sw.Write(s)
                sw.Write(newEndChar)

                currFilePos = s.Length + origEndCharCount
                linesRead += 1

                If linesRead Mod 1000 = 0 Then
                    Me.OnProgressUpdate("Normalizing Line Endings (" &
                     Math.Round(CDbl(currFilePos / fileLength * 100), 1).ToString &
                     " % complete", CSng(currFilePos / fileLength * 100))
                End If

                s = tr.ReadLine
            Loop
            tr.Close()
            sw.Close()

            Return newFileName
        Else
            Return pathOfFileToFix
        End If
    End Function

    Private Sub EvaluateRules(
     udtRuleDetails As IList(Of udtRuleDefinitionExtendedType),
     strProteinName As String,
     strTextToTest As String,
     intTestTextOffsetInLine As Integer,
     strEntireLine As String,
     intContextLength As Integer)

        Dim intIndex As Integer
        Dim reMatch As Match
        Dim strExtraInfo As String
        Dim intCharIndexOfMatch As Integer

        For intIndex = 0 To udtRuleDetails.Count - 1
            With udtRuleDetails(intIndex)

                reMatch = .reRule.Match(strTextToTest)

                If (.RuleDefinition.MatchIndicatesProblem And reMatch.Success) OrElse
                 Not (.RuleDefinition.MatchIndicatesProblem And Not reMatch.Success) Then

                    If .RuleDefinition.DisplayMatchAsExtraInfo Then
                        strExtraInfo = reMatch.ToString
                    Else
                        strExtraInfo = String.Empty
                    End If

                    intCharIndexOfMatch = intTestTextOffsetInLine + reMatch.Index
                    If .RuleDefinition.Severity >= 5 Then
                        RecordFastaFileError(mLineCount, intCharIndexOfMatch, strProteinName,
                         .RuleDefinition.CustomRuleID, strExtraInfo,
                         ExtractContext(strEntireLine, intCharIndexOfMatch, intContextLength))
                    Else
                        RecordFastaFileWarning(mLineCount, intCharIndexOfMatch, strProteinName,
                         .RuleDefinition.CustomRuleID, strExtraInfo,
                         ExtractContext(strEntireLine, intCharIndexOfMatch, intContextLength))
                    End If

                End If

            End With
        Next intIndex

    End Sub

    Private Function ExamineProteinName(
       ByRef strProteinName As String,
       lstProteinNames As ISet(Of String),
       <Out()> ByRef blnSkipDuplicateProtein As Boolean,
       ByRef blnProcessingDuplicateOrInvalidProtein As Boolean) As String

        Dim blnDuplicateName = lstProteinNames.Contains(strProteinName)
        blnSkipDuplicateProtein = False

        If blnDuplicateName AndAlso mGenerateFixedFastaFile Then
            If mFixedFastaOptions.RenameProteinsWithDuplicateNames Then

                Dim chLetterToAppend = "b"c
                Dim intNumberToAppend = 0
                Dim strNewProteinName As String

                Do
                    strNewProteinName = strProteinName & "-"c & chLetterToAppend
                    If intNumberToAppend > 0 Then
                        strNewProteinName &= intNumberToAppend.ToString
                    End If

                    blnDuplicateName = lstProteinNames.Contains(strNewProteinName)

                    If blnDuplicateName Then
                        ' Increment chLetterToAppend to the next letter and then try again to rename the protein
                        If chLetterToAppend = "z"c Then
                            ' We've reached "z"
                            ' Change back to "a" but increment intNumberToAppend
                            chLetterToAppend = "a"c
                            intNumberToAppend += 1
                        Else
                            ' chLetterToAppend = Chr(Asc(chLetterToAppend) + 1)
                            chLetterToAppend = Convert.ToChar(Convert.ToInt32(chLetterToAppend) + 1)
                        End If
                    End If
                Loop While blnDuplicateName

                RecordFastaFileWarning(mLineCount, 1, strProteinName, eMessageCodeConstants.RenamedProtein, "--> " & strNewProteinName, String.Empty)

                strProteinName = String.Copy(strNewProteinName)
                mFixedFastaStats.DuplicateNameProteinsRenamed += 1
                blnSkipDuplicateProtein = False

            ElseIf mFixedFastaOptions.KeepDuplicateNamedProteinsUnlessMatchingSequence Then
                blnSkipDuplicateProtein = False
            Else
                blnSkipDuplicateProtein = True
            End If

        End If

        If blnDuplicateName Then
            If blnSkipDuplicateProtein Or Not mGenerateFixedFastaFile Then
                RecordFastaFileError(mLineCount, 1, strProteinName, eMessageCodeConstants.DuplicateProteinName)
                If mSaveBasicProteinHashInfoFile Then
                    blnProcessingDuplicateOrInvalidProtein = False
                Else
                    blnProcessingDuplicateOrInvalidProtein = True
                    mFixedFastaStats.DuplicateNameProteinsSkipped += 1
                End If
            Else
                RecordFastaFileWarning(mLineCount, 1, strProteinName, eMessageCodeConstants.DuplicateProteinName)
                blnProcessingDuplicateOrInvalidProtein = False
            End If
        Else
            blnProcessingDuplicateOrInvalidProtein = False
        End If

        If Not lstProteinNames.Contains(strProteinName) Then
            lstProteinNames.Add(strProteinName)
        End If

        Return strProteinName

    End Function

    Private Function ExtractContext(strText As String, intStartIndex As Integer) As String
        Return ExtractContext(strText, intStartIndex, DEFAULT_CONTEXT_LENGTH)
    End Function

    Private Function ExtractContext(strText As String, intStartIndex As Integer, intContextLength As Integer) As String
        ' Note that intContextLength should be an odd number; if it isn't, we'll add 1 to it

        Dim intContextStartIndex As Integer
        Dim intContextEndIndex As Integer

        If intContextLength Mod 2 = 0 Then
            intContextLength += 1
        ElseIf intContextLength < 1 Then
            intContextLength = 1
        End If

        If strText Is Nothing Then
            Return String.Empty
        ElseIf strText.Length <= 1 Then
            Return strText
        Else
            ' Define the start index for extracting the context from strText
            intContextStartIndex = intStartIndex - CInt((intContextLength - 1) / 2)
            If intContextStartIndex < 0 Then intContextStartIndex = 0

            ' Define the end index for extracting the context from strText
            intContextEndIndex = Math.Max(intStartIndex + CInt((intContextLength - 1) / 2), intContextStartIndex + intContextLength - 1)
            If intContextEndIndex >= strText.Length Then
                intContextEndIndex = strText.Length - 1
            End If

            ' Return the context portion of strText
            Return strText.Substring(intContextStartIndex, intContextEndIndex - intContextStartIndex + 1)
        End If

    End Function

    Private Function FlattenAdditionalProteinList(oProteinSeqHashEntry As clsProteinHashInfo, chSepChar As Char) As String
        Return FlattenArray(oProteinSeqHashEntry.AdditionalProteins, chSepChar)
    End Function

    Private Function FlattenArray(lstItems As IEnumerable(Of String)) As String
        Return FlattenArray(lstItems, ControlChars.Tab)
    End Function

    Private Function FlattenArray(lstItems As IEnumerable(Of String), chSepChar As Char) As String
        If lstItems Is Nothing Then
            Return String.Empty
        Else
            Return FlattenArray(lstItems, lstItems.Count, chSepChar)
        End If
    End Function

    Private Function FlattenArray(lstItems As IEnumerable(Of String), intDataCount As Integer, chSepChar As Char) As String
        Dim intIndex As Integer
        Dim strResult As String

        If lstItems Is Nothing Then
            Return String.Empty
        ElseIf lstItems.Count = 0 OrElse intDataCount <= 0 Then
            Return String.Empty
        Else
            If intDataCount > lstItems.Count Then
                intDataCount = lstItems.Count
            End If

            strResult = lstItems(0)
            If strResult Is Nothing Then strResult = String.Empty

            For intIndex = 1 To intDataCount - 1
                If lstItems(intIndex) Is Nothing Then
                    strResult &= chSepChar
                Else
                    strResult &= chSepChar & lstItems(intIndex)
                End If
            Next intIndex
            Return strResult
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
    ''' <param name="strHeaderLine"></param>
    ''' <returns></returns>
    ''' <remarks>Used for determining protein name</remarks>
    Private Function GetBestSpaceIndex(strHeaderLine As String) As Integer

        Dim spaceIndex = strHeaderLine.IndexOf(" "c)
        Dim tabIndex = strHeaderLine.IndexOf(ControlChars.Tab)

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

    Public Overrides Function GetDefaultExtensionsToParse() As String() Implements IValidateFastaFile.GetDefaultExtensionsToParse
        Dim strExtensionsToParse(0) As String

        strExtensionsToParse(0) = ".fasta"

        Return strExtensionsToParse

    End Function

    Public Overrides Function GetErrorMessage() As String Implements IValidateFastaFile.GetCurrentErrorMessage
        ' Returns "" if no error

        Dim strErrorMessage As String

        If MyBase.ErrorCode = eProcessFilesErrorCodes.LocalizedError Or
           MyBase.ErrorCode = eProcessFilesErrorCodes.NoError Then
            Select Case mLocalErrorCode
                Case IValidateFastaFile.eValidateFastaFileErrorCodes.NoError
                    strErrorMessage = ""
                Case IValidateFastaFile.eValidateFastaFileErrorCodes.OptionsSectionNotFound
                    strErrorMessage = "The section " & XML_SECTION_OPTIONS & " was not found in the parameter file"
                Case IValidateFastaFile.eValidateFastaFileErrorCodes.ErrorReadingInputFile
                    strErrorMessage = "Error reading input file"
                Case IValidateFastaFile.eValidateFastaFileErrorCodes.ErrorCreatingStatsFile
                    strErrorMessage = "Error creating stats output file"
                Case IValidateFastaFile.eValidateFastaFileErrorCodes.ErrorVerifyingLinefeedAtEOF
                    strErrorMessage = "Error verifying linefeed at end of file"
                Case IValidateFastaFile.eValidateFastaFileErrorCodes.UnspecifiedError
                    strErrorMessage = "Unspecified localized error"
                Case Else
                    ' This shouldn't happen
                    strErrorMessage = "Unknown error state"
            End Select
        Else
            strErrorMessage = MyBase.GetBaseClassErrorMessage()
        End If

        Return strErrorMessage

    End Function

    Private Function GetFileErrorTextByIndex(intFileErrorIndex As Integer, strSepChar As String) As String

        Dim strProteinName As String

        If mFileErrorCount <= 0 Or intFileErrorIndex < 0 Or intFileErrorIndex >= mFileErrorCount Then
            Return String.Empty
        Else
            With mFileErrors(intFileErrorIndex)
                If .ProteinName Is Nothing OrElse .ProteinName.Length = 0 Then
                    strProteinName = "N/A"
                Else
                    strProteinName = String.Copy(.ProteinName)
                End If

                Return LookupMessageType(IValidateFastaFile.eMsgTypeConstants.ErrorMsg) & strSepChar &
                 "Line " & .LineNumber.ToString & strSepChar &
                 "Col " & .ColNumber.ToString & strSepChar &
                 strProteinName & strSepChar &
                 LookupMessageDescription(.MessageCode, .ExtraInfo) & strSepChar & .Context

            End With
        End If

    End Function

    Private Function GetFileErrorByIndex(intFileErrorIndex As Integer) As IValidateFastaFile.udtMsgInfoType

        If mFileErrorCount <= 0 Or intFileErrorIndex < 0 Or intFileErrorIndex >= mFileErrorCount Then
            Return New IValidateFastaFile.udtMsgInfoType
        Else
            Return mFileErrors(intFileErrorIndex)
        End If

    End Function

    ''' <summary>
    ''' Retrieve the errors reported by the validator
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>Used by clsCustomValidateFastaFiles</remarks>
    Protected Function GetFileErrors() As IValidateFastaFile.udtMsgInfoType()

        Dim udtFileErrors() As IValidateFastaFile.udtMsgInfoType

        If mFileErrorCount > 0 Then
            ReDim udtFileErrors(mFileErrorCount - 1)
            Array.Copy(mFileErrors, udtFileErrors, mFileErrorCount)
        Else
            ReDim udtFileErrors(0)
        End If

        Return udtFileErrors

    End Function


    Private Function GetFileWarningTextByIndex(intFileWarningIndex As Integer, strSepChar As String) As String
        Dim strProteinName As String

        If mFileWarningCount <= 0 Or intFileWarningIndex < 0 Or intFileWarningIndex >= mFileWarningCount Then
            Return String.Empty
        Else
            With mFileWarnings(intFileWarningIndex)
                If .ProteinName Is Nothing OrElse .ProteinName.Length = 0 Then
                    strProteinName = "N/A"
                Else
                    strProteinName = String.Copy(.ProteinName)
                End If

                Return LookupMessageType(IValidateFastaFile.eMsgTypeConstants.WarningMsg) & strSepChar &
                 "Line " & .LineNumber.ToString & strSepChar &
                 "Col " & .ColNumber.ToString & strSepChar &
                 strProteinName & strSepChar &
                 LookupMessageDescription(.MessageCode, .ExtraInfo) &
                 strSepChar & .Context

            End With
        End If

    End Function

    Private Function GetFileWarningByIndex(intFileWarningIndex As Integer) As IValidateFastaFile.udtMsgInfoType

        If mFileWarningCount <= 0 Or intFileWarningIndex < 0 Or intFileWarningIndex >= mFileWarningCount Then
            Return New IValidateFastaFile.udtMsgInfoType
        Else
            Return mFileWarnings(intFileWarningIndex)
        End If

    End Function

    ''' <summary>
    ''' Retrieve the warnings reported by the validator
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>Used by clsCustomValidateFastaFiles</remarks>
    Protected Function GetFileWarnings() As IValidateFastaFile.udtMsgInfoType()

        Dim udtFileWarnings() As IValidateFastaFile.udtMsgInfoType

        If mFileWarningCount > 0 Then
            ReDim udtFileWarnings(mFileWarningCount - 1)
            Array.Copy(mFileWarnings, udtFileWarnings, mFileWarningCount)
        Else
            ReDim udtFileWarnings(0)
        End If

        Return udtFileWarnings

    End Function

    Private Function GetProcessMemoryUsageWithTimestamp() As String
        Return GetTimeStamp() & ControlChars.Tab & clsMemoryUsageLogger.GetProcessMemoryUsageMB.ToString("0.0") & " MB in use"
    End Function

    Private Function GetTimeStamp() As String
        ' Record the current time
        Return DateTime.Now.ToShortDateString & " " & DateTime.Now.ToLongTimeString
    End Function

    Private Sub InitializeLocalVariables()

        MyBase.ShowMessages = False

        mLocalErrorCode = IValidateFastaFile.eValidateFastaFileErrorCodes.NoError

        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.AddMissingLinefeedatEOF) = False

        Me.MaximumFileErrorsToTrack = 5

        Me.MinimumProteinNameLength = DEFAULT_MINIMUM_PROTEIN_NAME_LENGTH
        Me.MaximumProteinNameLength = DEFAULT_MAXIMUM_PROTEIN_NAME_LENGTH
        Me.MaximumResiduesPerLine = DEFAULT_MAXIMUM_RESIDUES_PER_LINE
        Me.ProteinLineStartChar = DEFAULT_PROTEIN_LINE_START_CHAR

        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.AllowAsteriskInResidues) = False
        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.AllowDashInResidues) = False
        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.WarnBlankLinesBetweenProteins) = False
        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.WarnLineStartsWithSpace) = True

        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.CheckForDuplicateProteinNames) = True
        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.CheckForDuplicateProteinSequences) = True

        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.GenerateFixedFASTAFile) = False

        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaSplitOutMultipleRefsForKnownAccession) = True
        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.SplitOutMultipleRefsInProteinName) = False

        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaRenameDuplicateNameProteins) = False
        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaKeepDuplicateNamedProteins) = False

        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaConsolidateDuplicateProteinSeqs) = False
        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaConsolidateDupsIgnoreILDiff) = False

        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaTruncateLongProteinNames) = True
        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaWrapLongResidueLines) = True
        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaRemoveInvalidResidues) = False

        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.SaveProteinSequenceHashInfoFiles) = False

        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.SaveBasicProteinHashInfoFile) = False


        mProteinNameFirstRefSepChars = DEFAULT_PROTEIN_NAME_FIRST_REF_SEP_CHARS.ToCharArray
        mProteinNameSubsequentRefSepChars = DEFAULT_PROTEIN_NAME_SUBSEQUENT_REF_SEP_CHARS.ToCharArray

        mFixedFastaOptions.LongProteinNameSplitChars = New Char() {DEFAULT_LONG_PROTEIN_NAME_SPLIT_CHAR}
        mFixedFastaOptions.ProteinNameInvalidCharsToRemove = New Char() {}          ' Default to an empty character array

        SetDefaultRules()

        ResetStructures()
        mFastaFilePath = String.Empty

        mMemoryUsageLogger = New clsMemoryUsageLogger(String.Empty)
        mProcessMemoryUsageMBAtStart = clsMemoryUsageLogger.GetProcessMemoryUsageMB()

        If REPORT_DETAILED_MEMORY_USAGE Then
            ' mMemoryUsageMBAtStart = mMemoryUsageLogger.GetFreeMemoryMB()
            Console.WriteLine(MEM_USAGE_PREFIX & mMemoryUsageLogger.GetMemoryUsageHeader())
            Console.WriteLine(MEM_USAGE_PREFIX & mMemoryUsageLogger.GetMemoryUsageSummary())
        End If

        mTempFilesToDelete = New List(Of String)

    End Sub

    Private Sub InitializeRuleDetails(
     ByRef udtRuleDefs() As udtRuleDefinitionType,
     ByRef udtRuleDetails() As udtRuleDefinitionExtendedType)
        Dim intIndex As Integer

        If udtRuleDefs Is Nothing OrElse udtRuleDefs.Length = 0 Then
            ReDim udtRuleDetails(-1)
        Else
            ReDim udtRuleDetails(udtRuleDefs.Length - 1)

            For intIndex = 0 To udtRuleDefs.Length - 1
                Try
                    With udtRuleDetails(intIndex)
                        .RuleDefinition = udtRuleDefs(intIndex)
                        .reRule = New Regex(
                         .RuleDefinition.MatchRegEx,
                         RegexOptions.Singleline Or
                         RegexOptions.Compiled)
                        .Valid = True
                    End With
                Catch ex As Exception
                    ' Ignore the error, but mark .Valid = false
                    udtRuleDetails(intIndex).Valid = False
                End Try
            Next intIndex
        End If

    End Sub

    Private Function LoadExistingProteinHashFile(
      proteinHashFilePath As String,
      <Out()> ByRef lstPreloadedProteinNamesToKeep As clsNestedStringIntList) As Boolean

        ' List of protein names to keep
        ' Keys are protein names, values are the number of entries written to the fixed fasta file for the given protein name
        lstPreloadedProteinNamesToKeep = Nothing

        Try

            Dim fiProteinHashFile = New FileInfo(proteinHashFilePath)

            If Not fiProteinHashFile.Exists Then
                ShowErrorMessage("Protein hash file not found: " & proteinHashFilePath)
                Return False
            End If

            ' Sort the protein has file on the Sequence_Hash column
            ' First cache the column names from the header line

            Dim headerInfo = New Dictionary(Of String, Integer)(StringComparer.InvariantCultureIgnoreCase)
            Dim proteinHashFileLines As Long = 0
            Dim cachedHeaderLine As String = String.Empty

            ShowMessage("Examining pre-existing protein hash file to count the number of entries: " & Path.GetFileName(proteinHashFilePath))
            Dim dtLastStatus = DateTime.UtcNow
            Dim progressDotShown = False

            Using hashFileReader = New StreamReader(New FileStream(fiProteinHashFile.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
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
                        If DateTime.UtcNow.Subtract(dtLastStatus).TotalSeconds >= 10 Then
                            Console.Write(".")
                            progressDotShown = True
                            dtLastStatus = DateTime.UtcNow
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

            Dim baseHashFileName = Path.GetFileNameWithoutExtension(fiProteinHashFile.Name)
            Dim sortedProteinHashAddon As String
            Dim proteinHashFilenameSuffixNoExtension = Path.GetFileNameWithoutExtension(PROTEIN_HASHES_FILENAME_SUFFIX)

            If baseHashFileName.EndsWith(proteinHashFilenameSuffixNoExtension) Then
                baseHashFileName = baseHashFileName.Substring(0, baseHashFileName.Length - proteinHashFilenameSuffixNoExtension.Length)
                sortedProteinHashAddon = proteinHashFilenameSuffixNoExtension
            Else
                sortedProteinHashAddon = String.Empty
            End If

            Dim baseDataFileName = Path.Combine(fiProteinHashFile.Directory.FullName, baseHashFileName)

            ' Note: do not add sortedProteinHashFilePath to mTempFilesToDelete
            Dim sortedProteinHashFilePath = baseDataFileName & sortedProteinHashAddon & "_Sorted.tmp"

            Dim fiSortedProteinHashFile = New FileInfo(sortedProteinHashFilePath)
            Dim sortedHashFileLines As Long = 0
            Dim sortRequired = True

            If fiSortedProteinHashFile.Exists Then
                dtLastStatus = DateTime.UtcNow
                progressDotShown = False
                ShowMessage("Validating existing sorted protein hash file: " + fiSortedProteinHashFile.Name)

                ' The sorted file exists; if it has the same number of lines as the sort file, assume that it is complete
                Using sortedHashFileReader = New StreamReader(New FileStream(fiSortedProteinHashFile.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
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
                                If DateTime.UtcNow.Subtract(dtLastStatus).TotalSeconds >= 10 Then
                                    Console.Write((sortedHashFileLines / proteinHashFileLines * 100.0).ToString("0") & "% ")
                                    progressDotShown = True
                                    dtLastStatus = DateTime.UtcNow
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

                    fiSortedProteinHashFile.Delete()
                    Threading.Thread.Sleep(50)
                End If
            End If

            ' Create the sorted protein sequence hash file if necessary
            If sortRequired Then
                Console.WriteLine()
                ShowMessage("Sorting the existing protein hash file to create " + Path.GetFileName(sortedProteinHashFilePath))
                Dim sortHashSuccess = SortFile(fiProteinHashFile, sequenceHashColumnIndex, sortedProteinHashFilePath)
                If Not sortHashSuccess Then
                    Return False
                End If
            End If

            ShowMessage("Determining the best spanner length for protein names")

            ' Examine the protein names in the sequence hash file to determine the appropriate spanner length for the names
            Dim spannerCharLength = clsNestedStringIntList.AutoDetermineSpannerCharLength(fiProteinHashFile, proteinNameColumnIndex, True)
            Const RAISE_EXCEPTION_IF_ADDED_DATA_NOT_SORTED = True

            ' List of protein names to keep
            lstPreloadedProteinNamesToKeep = New clsNestedStringIntList(spannerCharLength, RAISE_EXCEPTION_IF_ADDED_DATA_NOT_SORTED)

            dtLastStatus = DateTime.UtcNow
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
                        If DateTime.UtcNow.Subtract(dtLastStatus).TotalSeconds >= 10 Then
                            Console.Write((linesRead / proteinHashFileLines * 100.0).ToString("0") & "% ")
                            progressDotShown = True
                            dtLastStatus = DateTime.UtcNow
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
                    lstPreloadedProteinNamesToKeep.Add(proteinName, 0)

                    lastProteinAdded = String.Copy(proteinName)
                End While
            End Using

            ' Confirm that the data is sorted
            lstPreloadedProteinNamesToKeep.Sort()

            ShowMessage("Cached " & lstPreloadedProteinNamesToKeep.Count.ToString("#,##0") & " protein names into memory")
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
                dtLastStatus = DateTime.UtcNow
                progressDotShown = False

                While Not sortedProteinNamesReader.EndOfStream
                    Dim dataLine = sortedProteinNamesReader.ReadLine()
                    Dim dataValues = dataLine.Split(ControlChars.Tab)

                    Dim currentProtein As String = dataValues(0)
                    Dim sequenceHash As String = dataValues(1)

                    Dim firstProteinForHash = lstPreloadedProteinNamesToKeep.Contains(currentProtein)
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
                        If DateTime.UtcNow.Subtract(dtLastStatus).TotalSeconds >= 10 Then
                            Console.Write((linesRead / proteinNamesUnsortedCount * 100.0).ToString("0") & "% ")
                            progressDotShown = True
                            dtLastStatus = DateTime.UtcNow
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
            If MyBase.ShowMessages Then
                ShowErrorMessage("Error in LoadExistingProteinHashFile: " + ex.Message)
                ShowExceptionStackTrace("LoadExistingProteinHashFile", ex)
                Return False
            Else
                ShowExceptionStackTrace("LoadExistingProteinHashFile", ex)
                Throw New Exception("Error in LoadExistingProteinHashFile", ex)
            End If
        End Try

    End Function

    Public Function LoadParameterFileSettings(
     strParameterFilePath As String) As Boolean Implements IValidateFastaFile.LoadParameterFileSettings

        Dim objSettingsFile As New XmlSettingsFileAccessor

        Dim blnCustomRulesLoaded As Boolean
        Dim blnSuccess As Boolean

        Dim strCharacterList As String

        Try

            If strParameterFilePath Is Nothing OrElse strParameterFilePath.Length = 0 Then
                ' No parameter file specified; nothing to load
                Return True
            End If

            If Not File.Exists(strParameterFilePath) Then
                ' See if strParameterFilePath points to a file in the same directory as the application
                strParameterFilePath = Path.Combine(
                 Path.GetDirectoryName(Reflection.Assembly.GetExecutingAssembly().Location),
                  Path.GetFileName(strParameterFilePath))
                If Not File.Exists(strParameterFilePath) Then
                    MyBase.SetBaseClassErrorCode(eProcessFilesErrorCodes.ParameterFileNotFound)
                    Return False
                End If
            End If

            If objSettingsFile.LoadSettings(strParameterFilePath) Then
                If Not objSettingsFile.SectionPresent(XML_SECTION_OPTIONS) Then
                    If MyBase.ShowMessages Then
                        ShowErrorMessage("The node '<section name=""" & XML_SECTION_OPTIONS & """> was not found in the parameter file: " & strParameterFilePath)
                    End If
                    MyBase.SetBaseClassErrorCode(eProcessFilesErrorCodes.InvalidParameterFile)
                    Return False
                Else
                    ' Read customized settings

                    Me.OptionSwitch(IValidateFastaFile.SwitchOptions.AddMissingLinefeedatEOF) =
                     objSettingsFile.GetParam(XML_SECTION_OPTIONS, "AddMissingLinefeedAtEOF",
                     Me.OptionSwitch(IValidateFastaFile.SwitchOptions.AddMissingLinefeedatEOF))
                    Me.OptionSwitch(IValidateFastaFile.SwitchOptions.AllowAsteriskInResidues) =
                     objSettingsFile.GetParam(XML_SECTION_OPTIONS, "AllowAsteriskInResidues",
                     Me.OptionSwitch(IValidateFastaFile.SwitchOptions.AllowAsteriskInResidues))
                    Me.OptionSwitch(IValidateFastaFile.SwitchOptions.AllowDashInResidues) =
                     objSettingsFile.GetParam(XML_SECTION_OPTIONS, "AllowDashInResidues",
                     Me.OptionSwitch(IValidateFastaFile.SwitchOptions.AllowDashInResidues))
                    Me.OptionSwitch(IValidateFastaFile.SwitchOptions.CheckForDuplicateProteinNames) =
                     objSettingsFile.GetParam(XML_SECTION_OPTIONS, "CheckForDuplicateProteinNames",
                     Me.OptionSwitch(IValidateFastaFile.SwitchOptions.CheckForDuplicateProteinNames))
                    Me.OptionSwitch(IValidateFastaFile.SwitchOptions.CheckForDuplicateProteinSequences) =
                     objSettingsFile.GetParam(XML_SECTION_OPTIONS, "CheckForDuplicateProteinSequences",
                     Me.OptionSwitch(IValidateFastaFile.SwitchOptions.CheckForDuplicateProteinSequences))

                    Me.OptionSwitch(IValidateFastaFile.SwitchOptions.SaveProteinSequenceHashInfoFiles) =
                     objSettingsFile.GetParam(XML_SECTION_OPTIONS, "SaveProteinSequenceHashInfoFiles",
                     Me.OptionSwitch(IValidateFastaFile.SwitchOptions.SaveProteinSequenceHashInfoFiles))

                    Me.OptionSwitch(IValidateFastaFile.SwitchOptions.SaveBasicProteinHashInfoFile) =
                     objSettingsFile.GetParam(XML_SECTION_OPTIONS, "SaveBasicProteinHashInfoFile",
                     Me.OptionSwitch(IValidateFastaFile.SwitchOptions.SaveBasicProteinHashInfoFile))

                    Me.MaximumFileErrorsToTrack = objSettingsFile.GetParam(XML_SECTION_OPTIONS,
                     "MaximumFileErrorsToTrack", Me.MaximumFileErrorsToTrack)
                    Me.MinimumProteinNameLength = objSettingsFile.GetParam(XML_SECTION_OPTIONS,
                     "MinimumProteinNameLength", Me.MinimumProteinNameLength)
                    Me.MaximumProteinNameLength = objSettingsFile.GetParam(XML_SECTION_OPTIONS,
                     "MaximumProteinNameLength", Me.MaximumProteinNameLength)
                    Me.MaximumResiduesPerLine = objSettingsFile.GetParam(XML_SECTION_OPTIONS,
                     "MaximumResiduesPerLine", Me.MaximumResiduesPerLine)

                    Me.OptionSwitch(IValidateFastaFile.SwitchOptions.WarnBlankLinesBetweenProteins) =
                     objSettingsFile.GetParam(XML_SECTION_OPTIONS, "WarnBlankLinesBetweenProteins",
                     Me.OptionSwitch(IValidateFastaFile.SwitchOptions.WarnBlankLinesBetweenProteins))
                    Me.OptionSwitch(IValidateFastaFile.SwitchOptions.WarnLineStartsWithSpace) =
                     objSettingsFile.GetParam(XML_SECTION_OPTIONS, "WarnLineStartsWithSpace",
                     Me.OptionSwitch(IValidateFastaFile.SwitchOptions.WarnLineStartsWithSpace))

                    Me.OptionSwitch(IValidateFastaFile.SwitchOptions.OutputToStatsFile) =
                     objSettingsFile.GetParam(XML_SECTION_OPTIONS, "OutputToStatsFile",
                     Me.OptionSwitch(IValidateFastaFile.SwitchOptions.OutputToStatsFile))

                    Me.OptionSwitch(IValidateFastaFile.SwitchOptions.NormalizeFileLineEndCharacters) =
                     objSettingsFile.GetParam(XML_SECTION_OPTIONS, "NormalizeFileLineEndCharacters",
                     Me.OptionSwitch(IValidateFastaFile.SwitchOptions.NormalizeFileLineEndCharacters))


                    If Not objSettingsFile.SectionPresent(XML_SECTION_FIXED_FASTA_FILE_OPTIONS) Then
                        ' "ValidateFastaFixedFASTAFileOptions" section not present
                        ' Only read the settings for GenerateFixedFASTAFile and SplitOutMultipleRefsInProteinName

                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.GenerateFixedFASTAFile) =
                         objSettingsFile.GetParam(XML_SECTION_OPTIONS, "GenerateFixedFASTAFile",
                         Me.OptionSwitch(IValidateFastaFile.SwitchOptions.GenerateFixedFASTAFile))

                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.SplitOutMultipleRefsInProteinName) =
                         objSettingsFile.GetParam(XML_SECTION_OPTIONS, "SplitOutMultipleRefsInProteinName",
                         Me.OptionSwitch(IValidateFastaFile.SwitchOptions.SplitOutMultipleRefsInProteinName))

                    Else
                        ' "ValidateFastaFixedFASTAFileOptions" section is present

                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.GenerateFixedFASTAFile) =
                         objSettingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "GenerateFixedFASTAFile",
                         Me.OptionSwitch(IValidateFastaFile.SwitchOptions.GenerateFixedFASTAFile))

                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.SplitOutMultipleRefsInProteinName) =
                         objSettingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "SplitOutMultipleRefsInProteinName",
                         Me.OptionSwitch(IValidateFastaFile.SwitchOptions.SplitOutMultipleRefsInProteinName))

                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaRenameDuplicateNameProteins) =
                         objSettingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "RenameDuplicateNameProteins",
                         Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaRenameDuplicateNameProteins))

                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaKeepDuplicateNamedProteins) =
                         objSettingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "KeepDuplicateNamedProteins",
                         Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaKeepDuplicateNamedProteins))

                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaConsolidateDuplicateProteinSeqs) =
                         objSettingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ConsolidateDuplicateProteinSeqs",
                         Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaConsolidateDuplicateProteinSeqs))

                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaConsolidateDupsIgnoreILDiff) =
                         objSettingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ConsolidateDupsIgnoreILDiff",
                         Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaConsolidateDupsIgnoreILDiff))

                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaTruncateLongProteinNames) =
                         objSettingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "TruncateLongProteinNames",
                         Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaTruncateLongProteinNames))

                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaSplitOutMultipleRefsForKnownAccession) =
                         objSettingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "SplitOutMultipleRefsForKnownAccession",
                         Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaSplitOutMultipleRefsForKnownAccession))

                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaWrapLongResidueLines) =
                         objSettingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "WrapLongResidueLines",
                         Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaWrapLongResidueLines))

                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaRemoveInvalidResidues) =
                         objSettingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "RemoveInvalidResidues",
                         Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaRemoveInvalidResidues))

                        ' Look for the special character lists
                        ' If defined, then update the default values
                        strCharacterList = objSettingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "LongProteinNameSplitChars", String.Empty)
                        If Not strCharacterList Is Nothing AndAlso strCharacterList.Length > 0 Then
                            ' Update mFixedFastaOptions.LongProteinNameSplitChars with strCharacterList
                            Me.LongProteinNameSplitChars = strCharacterList
                        End If

                        strCharacterList = objSettingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ProteinNameInvalidCharsToRemove", String.Empty)
                        If Not strCharacterList Is Nothing AndAlso strCharacterList.Length > 0 Then
                            ' Update mFixedFastaOptions.ProteinNameInvalidCharsToRemove with strCharacterList
                            Me.ProteinNameInvalidCharsToRemove = strCharacterList
                        End If

                        strCharacterList = objSettingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ProteinNameFirstRefSepChars", String.Empty)
                        If Not strCharacterList Is Nothing AndAlso strCharacterList.Length > 0 Then
                            ' Update mProteinNameFirstRefSepChars
                            Me.ProteinNameFirstRefSepChars = strCharacterList.ToCharArray
                        End If

                        strCharacterList = objSettingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ProteinNameSubsequentRefSepChars", String.Empty)
                        If Not strCharacterList Is Nothing AndAlso strCharacterList.Length > 0 Then
                            ' Update mProteinNameSubsequentRefSepChars
                            Me.ProteinNameSubsequentRefSepChars = strCharacterList.ToCharArray
                        End If
                    End If

                    ' Read the custom rules
                    ' If all of the sections are missing, then use the default rules
                    blnCustomRulesLoaded = False

                    blnSuccess = ReadRulesFromParameterFile(objSettingsFile, XML_SECTION_FASTA_HEADER_LINE_RULES, mHeaderLineRules)
                    blnCustomRulesLoaded = blnCustomRulesLoaded Or blnSuccess

                    blnSuccess = ReadRulesFromParameterFile(objSettingsFile, XML_SECTION_FASTA_PROTEIN_NAME_RULES, mProteinNameRules)
                    blnCustomRulesLoaded = blnCustomRulesLoaded Or blnSuccess

                    blnSuccess = ReadRulesFromParameterFile(objSettingsFile, XML_SECTION_FASTA_PROTEIN_DESCRIPTION_RULES, mProteinDescriptionRules)
                    blnCustomRulesLoaded = blnCustomRulesLoaded Or blnSuccess

                    blnSuccess = ReadRulesFromParameterFile(objSettingsFile, XML_SECTION_FASTA_PROTEIN_SEQUENCE_RULES, mProteinSequenceRules)
                    blnCustomRulesLoaded = blnCustomRulesLoaded Or blnSuccess
                End If
            End If
        Catch ex As Exception
            If MyBase.ShowMessages Then
                ShowErrorMessage("Error in LoadParameterFileSettings: " & ex.Message)
            Else
                Throw New Exception("Error in LoadParameterFileSettings", ex)
            End If
            Return False
        End Try

        If Not blnCustomRulesLoaded Then
            SetDefaultRules()
        End If

        Return True

    End Function

    Public Function LookupMessageDescription(intErrorMessageCode As Integer) As String _
     Implements IValidateFastaFile.LookupMessageDescription
        Return Me.LookupMessageDescription(intErrorMessageCode, Nothing)
    End Function

    Public Function LookupMessageDescription(intErrorMessageCode As Integer, strExtraInfo As String) As String _
     Implements IValidateFastaFile.LookupMessageDescription

        Dim strMessage As String
        Dim blnMatchFound As Boolean

        Select Case intErrorMessageCode
            ' Error messages
            Case eMessageCodeConstants.ProteinNameIsTooLong
                strMessage = "Protein name is longer than the maximum allowed length of " & mMaximumProteinNameLength.ToString & " characters"
                'Case eMessageCodeConstants.ProteinNameContainsInvalidCharacters
                '    strMessage = "Protein name contains invalid characters"
                '    If Not blnSpecifiedInvalidProteinNameChars Then
                '        strMessage &= " (should contain letters, numbers, period, dash, underscore, colon, comma, or vertical bar)"
                '        blnSpecifiedInvalidProteinNameChars = True
                '    End If
            Case eMessageCodeConstants.LineStartsWithSpace
                strMessage = "Found a line starting with a space"
                'Case eMessageCodeConstants.RightArrowFollowedBySpace
                '    strMessage = "Space found directly after the > symbol"
                'Case eMessageCodeConstants.RightArrowFollowedByTab
                '    strMessage = "Tab character found directly after the > symbol"
                'Case eMessageCodeConstants.RightArrowButNoProteinName
                '    strMessage = "Line starts with > but does not contain a protein name"
            Case eMessageCodeConstants.BlankLineBetweenProteinNameAndResidues
                strMessage = "A blank line was found between the protein name and its residues"
            Case eMessageCodeConstants.BlankLineInMiddleOfResidues
                strMessage = "A blank line was found in the middle of the residue block for the protein"
            Case eMessageCodeConstants.ResiduesFoundWithoutProteinHeader
                strMessage = "Residues were found, but a protein header didn't precede them"
                'Case eMessageCodeConstants.ResiduesWithAsterisk
                '    strMessage = "An asterisk was found in the residues"
                'Case eMessageCodeConstants.ResiduesWithSpace
                '    strMessage = "A space was found in the residues"
                'Case eMessageCodeConstants.ResiduesWithTab
                '    strMessage = "A tab character was found in the residues"
                'Case eMessageCodeConstants.ResiduesWithInvalidCharacters
                '    strMessage = "Invalid residues found"
                '    If Not blnSpecifiedResidueChars Then
                '        If mAllowAsteriskInResidues Then
                '            strMessage &= " (should be any capital letter except J, plus *)"
                '        Else
                '            strMessage &= " (should be any capital letter except J)"
                '        End If
                '        blnSpecifiedResidueChars = True
                '    End If
            Case eMessageCodeConstants.ProteinEntriesNotFound
                strMessage = "File does not contain any protein entries"
            Case eMessageCodeConstants.FinalProteinEntryMissingResidues
                strMessage = "The last entry in the file is a protein header line, but there is no protein sequence line after it"
            Case eMessageCodeConstants.FileDoesNotEndWithLinefeed
                strMessage = "File does not end in a blank line; this is a problem for Sequest"
            Case eMessageCodeConstants.DuplicateProteinName
                strMessage = "Duplicate protein name found"

                ' Warning messages
            Case eMessageCodeConstants.ProteinNameIsTooShort
                strMessage = "Protein name is shorter than the minimum suggested length of " & mMinimumProteinNameLength.ToString & " characters"
                'Case eMessageCodeConstants.ProteinNameContainsComma
                '    strMessage = "Protein name contains a comma"
                'Case eMessageCodeConstants.ProteinNameContainsVerticalBars
                '    strMessage = "Protein name contains two or more vertical bars"
                'Case eMessageCodeConstants.ProteinNameContainsWarningCharacters
                '    strMessage = "Protein name contains undesirable characters"
                'Case eMessageCodeConstants.ProteinNameWithoutDescription
                '    strMessage = "Line contains a protein name, but not a description"
            Case eMessageCodeConstants.BlankLineBeforeProteinName
                strMessage = "Blank line found before the protein name; this is acceptable, but not preferred"
                'Case eMessageCodeConstants.ProteinNameAndDescriptionSeparatedByTab
                '    strMessage = "Protein name is separated from the protein description by a tab"
            Case eMessageCodeConstants.ResiduesLineTooLong
                strMessage = "Residues line is longer than the suggested maximum length of " & mMaximumResiduesPerLine.ToString & " characters"
                'Case eMessageCodeConstants.ProteinDescriptionWithTab
                '    strMessage = "Protein description contains a tab character"
                'Case eMessageCodeConstants.ProteinDescriptionWithQuotationMark
                '    strMessage = "Protein description contains a quotation mark"
                'Case eMessageCodeConstants.ProteinDescriptionWithEscapedSlash
                '    strMessage = "Protein description contains escaped slash: \/"
                'Case eMessageCodeConstants.ProteinDescriptionWithUndesirableCharacter
                '    strMessage = "Protein description contains undesirable characters"
                'Case eMessageCodeConstants.ResiduesLineTooLong
                '    strMessage = "Residues line is longer than the suggested maximum length of " & mMaximumResiduesPerLine.ToString & " characters"
                'Case eMessageCodeConstants.ResiduesLineContainsU
                '    strMessage = "Residues line contains U (selenocysteine); this residue is unsupported by Sequest"

            Case eMessageCodeConstants.DuplicateProteinSequence
                strMessage = "Duplicate protein sequences found"

            Case eMessageCodeConstants.RenamedProtein
                strMessage = "Renamed protein because duplicate name"

            Case eMessageCodeConstants.ProteinRemovedSinceDuplicateSequence
                strMessage = "Removed protein since duplicate sequence"

            Case eMessageCodeConstants.DuplicateProteinNameRetained
                strMessage = "Duplicate protein retained in fixed file"

            Case eMessageCodeConstants.UnspecifiedError
                strMessage = "Unspecified error"
            Case Else
                strMessage = "Unspecified error"

                ' Search the custom rules for the given code
                blnMatchFound = SearchRulesForID(mHeaderLineRules, intErrorMessageCode, strMessage)

                If Not blnMatchFound Then
                    blnMatchFound = SearchRulesForID(mProteinNameRules, intErrorMessageCode, strMessage)
                End If

                If Not blnMatchFound Then
                    blnMatchFound = SearchRulesForID(mProteinDescriptionRules, intErrorMessageCode, strMessage)
                End If

                If Not blnMatchFound Then
                    blnMatchFound = SearchRulesForID(mProteinSequenceRules, intErrorMessageCode, strMessage)
                End If

        End Select

        If Not String.IsNullOrWhiteSpace(strExtraInfo) Then
            strMessage &= " (" & strExtraInfo & ")"
        End If

        Return strMessage

    End Function

    Private Function LookupMessageType(EntryType As IValidateFastaFile.eMsgTypeConstants) As String _
     Implements IValidateFastaFile.LookupMessageTypeString
        Select Case EntryType
            Case IValidateFastaFile.eMsgTypeConstants.ErrorMsg
                Return "Error"
            Case IValidateFastaFile.eMsgTypeConstants.WarningMsg
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
    Protected Function SimpleProcessFile(strInputFilePath As String) As Boolean Implements IValidateFastaFile.ValidateFASTAFile
        Return Me.ProcessFile(strInputFilePath, Nothing, Nothing, False)
    End Function

    Public Overloads Overrides Function ProcessFile(
     strInputFilePath As String,
     strOutputFolderPath As String,
     strParameterFilePath As String,
     blnResetErrorCode As Boolean) As Boolean Implements IValidateFastaFile.ValidateFASTAFile

        'Returns True if success, False if failure

        Dim ioFile As FileInfo
        Dim swStatsOutFile As StreamWriter

        Dim strInputFilePathFull As String
        Dim strStatusMessage As String

        If blnResetErrorCode Then
            SetLocalErrorCode(IValidateFastaFile.eValidateFastaFileErrorCodes.NoError)
        End If

        If Not LoadParameterFileSettings(strParameterFilePath) Then
            strStatusMessage = "Parameter file load error: " & strParameterFilePath
            If MyBase.ShowMessages Then ShowErrorMessage(strStatusMessage)
            If MyBase.ErrorCode = eProcessFilesErrorCodes.NoError Then
                MyBase.SetBaseClassErrorCode(eProcessFilesErrorCodes.InvalidParameterFile)
            End If
            Return False
        End If

        Try
            If strInputFilePath Is Nothing OrElse strInputFilePath.Length = 0 Then
                Console.WriteLine("Input file name is empty")
                MyBase.SetBaseClassErrorCode(eProcessFilesErrorCodes.InvalidInputFilePath)
                Return False
            Else

                Console.WriteLine()
                ShowMessage("Parsing " & Path.GetFileName(strInputFilePath))

                If Not CleanupFilePaths(strInputFilePath, strOutputFolderPath) Then
                    MyBase.SetBaseClassErrorCode(eProcessFilesErrorCodes.FilePathError)
                    Return False
                Else
                    ' List of protein names to keep
                    ' Keys are protein names, values are the number of entries written to the fixed fasta file for the given protein name
                    Dim lstPreloadedProteinNamesToKeep As clsNestedStringIntList = Nothing

                    If Not String.IsNullOrEmpty(ExistingProteinHashFile) Then
                        Dim loadSuccess = LoadExistingProteinHashFile(ExistingProteinHashFile, lstPreloadedProteinNamesToKeep)
                        If Not loadSuccess Then
                            Return False
                        End If
                    End If

                    Try

                        ' Obtain the full path to the input file
                        ioFile = New FileInfo(strInputFilePath)
                        strInputFilePathFull = ioFile.FullName

                        Dim blnSuccess = AnalyzeFastaFile(strInputFilePathFull, lstPreloadedProteinNamesToKeep)

                        If blnSuccess Then
                            ReportResults(strOutputFolderPath, mOutputToStatsFile)
                            DeleteTempFiles()
                            Return True
                        Else
                            If mOutputToStatsFile Then
                                mStatsFilePath = ConstructStatsFilePath(strOutputFolderPath)
                                swStatsOutFile = New StreamWriter(mStatsFilePath, True)
                                swStatsOutFile.WriteLine(GetTimeStamp() & ControlChars.Tab &
                                                         "Error parsing " &
                                                         Path.GetFileName(strInputFilePath) & ": " & Me.GetErrorMessage())
                                swStatsOutFile.Close()
                            Else
                                ShowMessage("Error parsing " &
                                  Path.GetFileName(strInputFilePath) &
                                  ": " & Me.GetErrorMessage())
                            End If
                            Return False
                        End If

                    Catch ex As Exception
                        If MyBase.ShowMessages Then
                            ShowErrorMessage("Error calling AnalyzeFastaFile: " + ex.Message)
                            ShowExceptionStackTrace("ProcessFile (AnalyzeFastaFile)", ex)
                        Else
                            ShowExceptionStackTrace("ProcessFile (AnalyzeFastaFile)", ex)
                            Throw New Exception("Error calling AnalyzeFastaFile", ex)
                        End If
                        Return False
                    End Try
                End If
            End If
        Catch ex As Exception
            If MyBase.ShowMessages Then
                ShowErrorMessage("Error in ProcessFile:" & ex.Message)
                ShowExceptionStackTrace("ProcessFile", ex)
            Else
                ShowExceptionStackTrace("ProcessFile", ex)
                Throw New Exception("Error in ProcessFile", ex)
            End If
            Return False
        End Try

    End Function

    Private Sub PrependExtraTextToProteinDescription(strExtraProteinNameText As String, ByRef strProteinDescription As String)
        Static chExtraCharsToTrim As Char() = New Char() {"|"c, " "c}

        If Not strExtraProteinNameText Is Nothing AndAlso strExtraProteinNameText.Length > 0 Then
            ' If strExtraProteinNameText ends in a vertical bar and/or space, them remove them
            strExtraProteinNameText = strExtraProteinNameText.TrimEnd(chExtraCharsToTrim)

            If Not strProteinDescription Is Nothing AndAlso strProteinDescription.Length > 0 Then
                If strProteinDescription.Chars(0) = " "c OrElse strProteinDescription.Chars(0) = "|"c Then
                    strProteinDescription = strExtraProteinNameText & strProteinDescription
                Else
                    strProteinDescription = strExtraProteinNameText & " " & strProteinDescription
                End If
            Else
                strProteinDescription = String.Copy(strExtraProteinNameText)
            End If
        End If


    End Sub

    Private Sub ProcessResiduesForPreviousProtein(
      strProteinName As String,
      sbCurrentResidues As StringBuilder,
      lstProteinSequenceHashes As clsNestedStringDictionary(Of Integer),
      ByRef intProteinSequenceHashCount As Integer,
      ByRef oProteinSeqHashInfo() As clsProteinHashInfo,
      blnConsolidateDupsIgnoreILDiff As Boolean,
      swFixedFastaOut As TextWriter,
      intCurrentValidResidueLineLengthMax As Integer,
      swProteinSequenceHashBasic As TextWriter)

        Dim intWrapLength As Integer
        Dim intResidueCount As Integer

        Dim intIndex As Integer
        Dim intLength As Integer

        ' Check for and remove any asterisks at the end of the residues
        While sbCurrentResidues.Length > 0 AndAlso sbCurrentResidues.Chars(sbCurrentResidues.Length - 1) = "*"
            sbCurrentResidues.Remove(sbCurrentResidues.Length - 1, 1)
        End While

        If sbCurrentResidues.Length > 0 Then

            ' Remove any spaces from the residues

            If mCheckForDuplicateProteinSequences OrElse mSaveBasicProteinHashInfoFile Then
                ' Process the previous protein entry to store a hash of the protein sequence
                ProcessSequenceHashInfo(
                    strProteinName, sbCurrentResidues,
                    lstProteinSequenceHashes,
                    intProteinSequenceHashCount, oProteinSeqHashInfo,
                    blnConsolidateDupsIgnoreILDiff, swProteinSequenceHashBasic)
            End If

            If mGenerateFixedFastaFile AndAlso mFixedFastaOptions.WrapLongResidueLines Then
                ' Write out the residues
                ' Wrap the lines at intCurrentValidResidueLineLengthMax characters (but do not allow to be longer than mMaximumResiduesPerLine residues)

                intWrapLength = intCurrentValidResidueLineLengthMax
                If intWrapLength <= 0 OrElse intWrapLength > mMaximumResiduesPerLine Then
                    intWrapLength = mMaximumResiduesPerLine
                End If

                If intWrapLength < 10 Then
                    ' Do not allow intWrapLength to be less than 10
                    intWrapLength = 10
                End If

                intIndex = 0
                intResidueCount = sbCurrentResidues.Length
                Do While intIndex < sbCurrentResidues.Length
                    intLength = Math.Min(intWrapLength, intResidueCount - intIndex)
                    swFixedFastaOut.WriteLine(sbCurrentResidues.ToString(intIndex, intLength))
                    intIndex += intWrapLength
                Loop

            End If

            sbCurrentResidues.Length = 0

        End If

    End Sub

    Private Sub ProcessSequenceHashInfo(
      strProteinName As String,
      sbCurrentResidues As StringBuilder,
      lstProteinSequenceHashes As clsNestedStringDictionary(Of Integer),
      ByRef intProteinSequenceHashCount As Integer,
      ByRef oProteinSeqHashInfo() As clsProteinHashInfo,
      blnConsolidateDupsIgnoreILDiff As Boolean,
      swProteinSequenceHashBasic As TextWriter)

        Dim strComputedHash As String

        Try
            If sbCurrentResidues.Length > 0 Then

                ' Compute the hash value for sbCurrentResidues
                strComputedHash = ComputeProteinHash(sbCurrentResidues, blnConsolidateDupsIgnoreILDiff)

                If Not swProteinSequenceHashBasic Is Nothing Then

                    Dim dataValues = New List(Of String) From {
                        mProteinCount.ToString,
                        strProteinName,
                        sbCurrentResidues.Length.ToString,
                        strComputedHash}

                    swProteinSequenceHashBasic.WriteLine(FlattenList(dataValues))
                End If

                If mCheckForDuplicateProteinSequences AndAlso Not lstProteinSequenceHashes Is Nothing Then

                    ' See if lstProteinSequenceHashes contains strHash
                    Dim intSeqHashLookupPointer As Integer
                    If lstProteinSequenceHashes.TryGetValue(strComputedHash, intSeqHashLookupPointer) Then

                        ' Value exists; update the entry in oProteinSeqHashInfo
                        CachedSequenceHashInfoUpdate(oProteinSeqHashInfo(intSeqHashLookupPointer), strProteinName)

                    Else
                        ' Value not yet present; add it
                        CachedSequenceHashInfoUpdateAppend(
                            intProteinSequenceHashCount, oProteinSeqHashInfo,
                            strComputedHash, sbCurrentResidues, strProteinName)

                        lstProteinSequenceHashes.Add(strComputedHash, intProteinSequenceHashCount)
                        intProteinSequenceHashCount += 1

                    End If

                End If
            End If

        Catch ex As Exception
            ' Error caught; pass it up to the calling function
            ShowMessage(ex.Message)
            ShowExceptionStackTrace("ProcessSequenceHashInfo", ex)
            Throw
        End Try


    End Sub

    Private Sub CachedSequenceHashInfoUpdate(oProteinSeqHashInfo As clsProteinHashInfo, strProteinName As String)

        If oProteinSeqHashInfo.ProteinNameFirst = strProteinName Then
            oProteinSeqHashInfo.DuplicateProteinNameCount += 1
        Else
            oProteinSeqHashInfo.AddAdditionalProtein(strProteinName)
        End If
    End Sub

    Private Sub CachedSequenceHashInfoUpdateAppend(
      ByRef intProteinSequenceHashCount As Integer,
      ByRef oProteinSeqHashInfo() As clsProteinHashInfo,
      strComputedHash As String,
      sbCurrentResidues As StringBuilder,
      strProteinName As String)

        If intProteinSequenceHashCount >= oProteinSeqHashInfo.Length Then
            ' Need to reserve more space in oProteinSeqHashInfo
            If oProteinSeqHashInfo.Length < 1000000 Then
                ReDim Preserve oProteinSeqHashInfo(oProteinSeqHashInfo.Length * 2 - 1)
            Else
                ReDim Preserve oProteinSeqHashInfo(CInt(oProteinSeqHashInfo.Length * 1.2) - 1)
            End If

        End If

        Dim newProteinHashInfo = New clsProteinHashInfo(strComputedHash, sbCurrentResidues, strProteinName)
        oProteinSeqHashInfo(intProteinSequenceHashCount) = newProteinHashInfo

    End Sub

    Private Function ReadRulesFromParameterFile(
     objSettingsFile As XmlSettingsFileAccessor,
     strSectionName As String,
     ByRef udtRules() As udtRuleDefinitionType) As Boolean
        ' Returns True if the section named strSectionName is present and if it contains an item with keyName = "RuleCount"
        ' Note: even if RuleCount = 0, this function will return True

        Dim blnSuccess = False
        Dim intRuleCount As Integer
        Dim intRuleNumber As Integer

        Dim strRuleBase As String

        Dim udtNewRule As udtRuleDefinitionType

        intRuleCount = objSettingsFile.GetParam(strSectionName, XML_OPTION_ENTRY_RULE_COUNT, -1)

        If intRuleCount >= 0 Then
            ClearRulesDataStructure(udtRules)

            For intRuleNumber = 1 To intRuleCount
                strRuleBase = "Rule" & intRuleNumber.ToString

                udtNewRule.MatchRegEx = objSettingsFile.GetParam(strSectionName, strRuleBase & "MatchRegEx", String.Empty)

                If udtNewRule.MatchRegEx.Length > 0 Then
                    ' Only read the rule settings if MatchRegEx contains 1 or more characters

                    With udtNewRule
                        .MatchIndicatesProblem = objSettingsFile.GetParam(strSectionName, strRuleBase & "MatchIndicatesProblem", True)
                        .MessageWhenProblem = objSettingsFile.GetParam(strSectionName, strRuleBase & "MessageWhenProblem", "Error found with RegEx " & .MatchRegEx)
                        .Severity = objSettingsFile.GetParam(strSectionName, strRuleBase & "Severity", 3S)
                        .DisplayMatchAsExtraInfo = objSettingsFile.GetParam(strSectionName, strRuleBase & "DisplayMatchAsExtraInfo", False)

                        SetRule(udtRules, .MatchRegEx, .MatchIndicatesProblem, .MessageWhenProblem, .Severity, .DisplayMatchAsExtraInfo)

                    End With

                End If

            Next intRuleNumber

            blnSuccess = True
        End If

        Return blnSuccess

    End Function

    Private Sub RecordFastaFileError(
     intLineNumber As Integer,
     intCharIndex As Integer,
     strProteinName As String,
     intErrorMessageCode As Integer)

        RecordFastaFileError(intLineNumber, intCharIndex, strProteinName,
         intErrorMessageCode, String.Empty, String.Empty)
    End Sub

    Private Sub RecordFastaFileError(
     intLineNumber As Integer,
     intCharIndex As Integer,
     strProteinName As String,
     intErrorMessageCode As Integer,
     strExtraInfo As String,
     strContext As String)
        RecordFastaFileProblemWork(
         mFileErrorStats,
         mFileErrorCount,
         mFileErrors,
         intLineNumber,
         intCharIndex,
         strProteinName,
         intErrorMessageCode,
         strExtraInfo, strContext)
    End Sub

    Private Sub RecordFastaFileWarning(
     intLineNumber As Integer,
     intCharIndex As Integer,
     strProteinName As String,
     intWarningMessageCode As Integer)
        RecordFastaFileWarning(
         intLineNumber,
         intCharIndex,
         strProteinName,
         intWarningMessageCode,
         String.Empty, String.Empty)
    End Sub

    Private Sub RecordFastaFileWarning(
     intLineNumber As Integer,
     intCharIndex As Integer,
     strProteinName As String,
     intWarningMessageCode As Integer,
     strExtraInfo As String, strContext As String)

        RecordFastaFileProblemWork(mFileWarningStats, mFileWarningCount,
         mFileWarnings, intLineNumber, intCharIndex, strProteinName,
         intWarningMessageCode, strExtraInfo, strContext)
    End Sub

    Private Sub RecordFastaFileProblemWork(
     ByRef udtItemSummaryIndexed As udtItemSummaryIndexedType,
     ByRef intItemCountSpecified As Integer,
     ByRef udtItems() As IValidateFastaFile.udtMsgInfoType,
     intLineNumber As Integer,
     intCharIndex As Integer,
     strProteinName As String,
     intMessageCode As Integer,
     strExtraInfo As String,
     strContext As String)

        ' Note that intCharIndex is the index in the source string at which the error occurred
        ' When storing in .ColNumber, we add 1 to intCharIndex

        ' Lookup the index of the entry with intMessageCode in udtItemSummaryIndexed.ErrorStats
        ' Add it if not present

        Try
            Dim intItemIndex As Integer

            With udtItemSummaryIndexed
                If Not .MessageCodeToArrayIndex.TryGetValue(intMessageCode, intItemIndex) Then
                    If .ErrorStats.Length <= 0 Then
                        ReDim .ErrorStats(1)
                    ElseIf .ErrorStatsCount = .ErrorStats.Length Then
                        ReDim Preserve .ErrorStats(.ErrorStats.Length * 2 - 1)
                    End If
                    intItemIndex = .ErrorStatsCount
                    .ErrorStats(intItemIndex).MessageCode = intMessageCode
                    .MessageCodeToArrayIndex.Add(intMessageCode, intItemIndex)
                    .ErrorStatsCount += 1
                End If

            End With

            With udtItemSummaryIndexed.ErrorStats(intItemIndex)
                If .CountSpecified >= mMaximumFileErrorsToTrack Then
                    .CountUnspecified += 1
                Else
                    If udtItems.Length <= 0 Then
                        ' Initially reserve space for 10 errors
                        ReDim udtItems(10)
                    ElseIf intItemCountSpecified >= udtItems.Length Then
                        ' Double the amount of space reserved for errors
                        ReDim Preserve udtItems(udtItems.Length * 2 - 1)
                    End If

                    With udtItems(intItemCountSpecified)
                        .LineNumber = intLineNumber
                        .ColNumber = intCharIndex + 1
                        If strProteinName Is Nothing Then
                            .ProteinName = String.Empty
                        Else
                            .ProteinName = strProteinName
                        End If

                        .MessageCode = intMessageCode
                        If strExtraInfo Is Nothing Then
                            .ExtraInfo = String.Empty
                        Else
                            .ExtraInfo = strExtraInfo
                        End If

                        If strExtraInfo Is Nothing Then
                            .Context = String.Empty
                        Else
                            .Context = strContext
                        End If

                    End With
                    intItemCountSpecified += 1

                    .CountSpecified += 1
                End If

            End With

        Catch ex As Exception
            ' Ignore any errors that occur, but output the error to the console
            ShowMessage("Error in RecordFastaFileProblemWork: " & ex.Message)
            ShowExceptionStackTrace("RecordFastaFileProblemWork", ex)
        End Try

    End Sub

    Private Sub ReplaceXMLCodesWithText(strParameterFilePath As String)

        Dim strOutputFilePath As String
        Dim strTimeStamp As String
        Dim strLineIn As String

        Try
            ' Define the output file path
            strTimeStamp = GetTimeStamp().Replace(" ", "_").Replace(":", "_").Replace("/", "_")

            strOutputFilePath = strParameterFilePath & "_" & strTimeStamp & ".fixed"

            ' Open the input file and output file
            Using srInFile = New StreamReader(strParameterFilePath),
                swOutFile = New StreamWriter(New FileStream(strOutputFilePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))

                ' Parse each line in the file
                Do While Not srInFile.EndOfStream
                    strLineIn = srInFile.ReadLine()

                    If Not strLineIn Is Nothing Then
                        strLineIn = strLineIn.Replace("&gt;", ">").Replace("&lt;", "<")
                        swOutFile.WriteLine(strLineIn)
                    End If
                Loop

                ' Close the input and output files
            End Using

            ' Wait 100 msec
            Threading.Thread.Sleep(100)

            ' Delete the input file
            File.Delete(strParameterFilePath)

            ' Wait 250 msec
            Threading.Thread.Sleep(250)

            ' Rename the output file to the input file
            File.Move(strOutputFilePath, strParameterFilePath)

        Catch ex As Exception
            If MyBase.ShowMessages Then
                ShowErrorMessage("Error in ReplaceXMLCodesWithText:" & ex.Message)
                ShowExceptionStackTrace("ReplaceXMLCodesWithText", ex)
            Else
                Throw New Exception("Error in ReplaceXMLCodesWithText", ex)
            End If
        End Try

    End Sub

    Private Sub ReportMemoryUsage()

        If REPORT_DETAILED_MEMORY_USAGE Then
            Console.WriteLine(MEM_USAGE_PREFIX & mMemoryUsageLogger.GetMemoryUsageSummary())
        Else
            Console.WriteLine(MEM_USAGE_PREFIX & GetProcessMemoryUsageWithTimestamp())
        End If

    End Sub

    Private Sub ReportMemoryUsage(
      lstPreloadedProteinNamesToKeep As clsNestedStringIntList,
      lstProteinSequenceHashes As clsNestedStringDictionary(Of Integer),
      lstProteinNames As ICollection,
      oProteinSeqHashInfo() As clsProteinHashInfo)

        Console.WriteLine()
        ReportMemoryUsage()

        If Not lstPreloadedProteinNamesToKeep Is Nothing AndAlso lstPreloadedProteinNamesToKeep.Count > 0 Then
            Console.WriteLine(" lstPreloadedProteinNamesToKeep: {0,12:#,##0} records", lstPreloadedProteinNamesToKeep.Count)
        End If

        If lstProteinSequenceHashes.Count > 0 Then
            Console.WriteLine(" lstProteinSequenceHashes:  {0,12:#,##0} records", lstProteinSequenceHashes.Count)
            Console.WriteLine("   {0}", lstProteinSequenceHashes.GetSizeSummary())
        End If

        Console.WriteLine(" lstProteinNames:           {0,12:#,##0} records", lstProteinNames.Count)
        Console.WriteLine(" oProteinSeqHashInfo:       {0,12:#,##0} records", oProteinSeqHashInfo.Count)

    End Sub

    Private Sub ReportMemoryUsage(
       lstProteinNameFirst As clsNestedStringDictionary(Of Integer),
       lstProteinsWritten As clsNestedStringDictionary(Of Integer),
       lstDuplicateProteinList As clsNestedStringDictionary(Of String))

        Console.WriteLine()
        ReportMemoryUsage()
        Console.WriteLine(" lstProteinNameFirst:      {0,12:#,##0} records", lstProteinNameFirst.Count)
        Console.WriteLine("   {0}", lstProteinNameFirst.GetSizeSummary())
        Console.WriteLine(" lstProteinsWritten:       {0,12:#,##0} records", lstProteinsWritten.Count)
        Console.WriteLine("   {0}", lstProteinsWritten.GetSizeSummary())
        Console.WriteLine(" lstDuplicateProteinList:  {0,12:#,##0} records", lstDuplicateProteinList.Count)
        Console.WriteLine("   {0}", lstDuplicateProteinList.GetSizeSummary())

    End Sub

    Private Sub ReportResults(
     outputFolderPath As String,
     outputToStatsFile As Boolean)

        Dim srOutFile As StreamWriter = Nothing

        Dim iErrorInfoComparerClass As ErrorInfoComparerClass

        Dim strSourceFile As String
        Dim strProteinName As String

        Dim strSepChar As String

        Dim intIndex As Integer
        Dim intRetryCount As Integer

        Dim blnSuccess As Boolean
        Dim blnFileAlreadyExists As Boolean

        Try
            Try
                strSourceFile = Path.GetFileName(mFastaFilePath)
            Catch ex As Exception
                strSourceFile = "Unknown_filename_due_to_error.fasta"
            End Try

            If outputToStatsFile Then
                mStatsFilePath = ConstructStatsFilePath(outputFolderPath)
                blnFileAlreadyExists = File.Exists(mStatsFilePath)

                blnSuccess = False
                intRetryCount = 0

                Do While Not blnSuccess And intRetryCount < 5
                    Dim objOutStream As FileStream = Nothing
                    Try
                        objOutStream = New FileStream(mStatsFilePath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite)
                        srOutFile = New StreamWriter(objOutStream)
                        blnSuccess = True
                    Catch ex As Exception
                        ' Failed to open file, wait 1 second, then try again

                        srOutFile = Nothing
                        If Not objOutStream Is Nothing Then
                            objOutStream.Close()
                        End If

                        intRetryCount += 1
                        Threading.Thread.Sleep(1000)
                    End Try
                Loop

                If blnSuccess Then
                    strSepChar = ControlChars.Tab
                    If Not blnFileAlreadyExists Then
                        ' Write the header line
                        srOutFile.WriteLine(
                         "Date" & strSepChar &
                         "SourceFile" & strSepChar &
                         "MessageType" & strSepChar &
                         "LineNumber" & strSepChar &
                         "ColumnNumber" & strSepChar &
                         "Description_or_Protein" & strSepChar &
                         "Info" & strSepChar &
                         "Context")
                    End If
                Else
                    strSepChar = ", "
                    outputToStatsFile = False
                    SetLocalErrorCode(IValidateFastaFile.eValidateFastaFileErrorCodes.ErrorCreatingStatsFile)
                End If
            Else
                strSepChar = ", "
            End If

            ReportResultAddEntry(
             strSourceFile,
             IValidateFastaFile.eMsgTypeConstants.StatusMsg,
             "Full path to file",
             mFastaFilePath,
             String.Empty,
             outputToStatsFile,
             srOutFile,
             strSepChar)

            ReportResultAddEntry(
             strSourceFile,
             IValidateFastaFile.eMsgTypeConstants.StatusMsg,
             "Protein count",
             mProteinCount.ToString("#,##0"),
             String.Empty,
             outputToStatsFile,
             srOutFile,
             strSepChar)

            ReportResultAddEntry(
             strSourceFile,
             IValidateFastaFile.eMsgTypeConstants.StatusMsg,
             "Residue count",
             mResidueCount.ToString("#,##0"),
             String.Empty,
             outputToStatsFile,
             srOutFile,
             strSepChar)

            If mFileErrorCount > 0 Then
                ReportResultAddEntry(
                 strSourceFile,
                 IValidateFastaFile.eMsgTypeConstants.ErrorMsg,
                 "Error count",
                 Me.ErrorWarningCounts(
                  IValidateFastaFile.eMsgTypeConstants.ErrorMsg,
                  IValidateFastaFile.ErrorWarningCountTypes.Total).ToString,
                 String.Empty,
                 outputToStatsFile,
                 srOutFile, strSepChar)

                If mFileErrorCount > 1 Then
                    iErrorInfoComparerClass = New ErrorInfoComparerClass
                    Array.Sort(mFileErrors, 0, mFileErrorCount, iErrorInfoComparerClass)
                End If

                For intIndex = 0 To mFileErrorCount - 1
                    With mFileErrors(intIndex)
                        If .ProteinName Is Nothing OrElse .ProteinName.Length = 0 Then
                            strProteinName = "N/A"
                        Else
                            strProteinName = String.Copy(.ProteinName)
                        End If

                        Dim messageDescription = LookupMessageDescription(.MessageCode, .ExtraInfo)

                        ReportResultAddEntry(strSourceFile,
                           IValidateFastaFile.eMsgTypeConstants.ErrorMsg,
                           .LineNumber,
                           .ColNumber,
                           strProteinName,
                           messageDescription,
                           .Context,
                           outputToStatsFile,
                           srOutFile,
                           strSepChar)

                    End With
                Next intIndex
            End If

            If mFileWarningCount > 0 Then
                ReportResultAddEntry(
                 strSourceFile,
                 IValidateFastaFile.eMsgTypeConstants.WarningMsg,
                 "Warning count",
                 Me.ErrorWarningCounts(
                  IValidateFastaFile.eMsgTypeConstants.WarningMsg,
                  IValidateFastaFile.ErrorWarningCountTypes.Total).ToString,
                 String.Empty,
                 outputToStatsFile,
                 srOutFile,
                 strSepChar)

                If mFileWarningCount > 1 Then
                    iErrorInfoComparerClass = New ErrorInfoComparerClass
                    Array.Sort(mFileWarnings, 0, mFileWarningCount, iErrorInfoComparerClass)
                End If

                For intIndex = 0 To mFileWarningCount - 1
                    With mFileWarnings(intIndex)
                        If .ProteinName Is Nothing OrElse .ProteinName.Length = 0 Then
                            strProteinName = "N/A"
                        Else
                            strProteinName = String.Copy(.ProteinName)
                        End If

                        ReportResultAddEntry(strSourceFile,
                          IValidateFastaFile.eMsgTypeConstants.WarningMsg,
                          .LineNumber,
                          .ColNumber,
                          strProteinName,
                          LookupMessageDescription(.MessageCode, .ExtraInfo),
                          .Context,
                          outputToStatsFile,
                          srOutFile,
                          strSepChar)

                    End With
                Next intIndex
            End If

            Dim fiFastaFile = New FileInfo(mFastaFilePath)

            ' # Proteins, # Peptides, FileSizeKB
            ReportResultAddEntry(
              strSourceFile,
              IValidateFastaFile.eMsgTypeConstants.StatusMsg,
              "Summary line",
              mProteinCount.ToString() & " proteins, " & mResidueCount.ToString() & " residues, " & (fiFastaFile.Length / 1024.0).ToString("0") & " KB",
              String.Empty,
              outputToStatsFile,
              srOutFile,
              strSepChar)

            If outputToStatsFile AndAlso Not srOutFile Is Nothing Then
                srOutFile.Close()
            End If

        Catch ex As Exception
            If MyBase.ShowMessages Then
                ShowErrorMessage("Error in ReportResults:" & ex.Message)
                ShowExceptionStackTrace("ReportResults", ex)
            Else
                ShowExceptionStackTrace("ReportResults", ex)
                Throw New Exception("Error in ReportResults", ex)
            End If

        End Try

    End Sub

    Private Sub ReportResultAddEntry(
      strSourceFile As String,
      EntryType As IValidateFastaFile.eMsgTypeConstants,
      strDescriptionOrProteinName As String,
      strInfo As String,
      strContext As String,
      blnOutputToStatsFile As Boolean,
      srOutFile As TextWriter,
      strSepChar As String)

        ReportResultAddEntry(
         strSourceFile,
         EntryType, 0, 0,
         strDescriptionOrProteinName,
         strInfo,
         strContext,
         blnOutputToStatsFile,
         srOutFile, strSepChar)
    End Sub

    Private Sub ReportResultAddEntry(
      strSourceFile As String,
      EntryType As IValidateFastaFile.eMsgTypeConstants,
      intLineNumber As Integer,
      intColNumber As Integer,
      strDescriptionOrProteinName As String,
      strInfo As String,
      strContext As String,
      blnOutputToStatsFile As Boolean,
      srOutFile As TextWriter,
      strSepChar As String)

        Dim strMessage As String

        strMessage = strSourceFile & strSepChar &
         LookupMessageType(EntryType) & strSepChar &
         intLineNumber.ToString & strSepChar &
         intColNumber.ToString & strSepChar &
         strDescriptionOrProteinName & strSepChar &
         strInfo

        If Not strContext Is Nothing AndAlso strContext.Length > 0 Then
            strMessage &= strSepChar & strContext
        End If

        If blnOutputToStatsFile Then
            srOutFile.WriteLine(GetTimeStamp() & strSepChar & strMessage)
        Else
            Console.WriteLine(strMessage)
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

    Private Sub ResetItemSummaryStructure(ByRef udtItemSummary As udtItemSummaryIndexedType)
        With udtItemSummary
            .ErrorStatsCount = 0
            ReDim .ErrorStats(-1)
            If .MessageCodeToArrayIndex Is Nothing Then
                .MessageCodeToArrayIndex = New Dictionary(Of Integer, Integer)
            Else
                .MessageCodeToArrayIndex.Clear()
            End If
        End With

    End Sub

    Private Sub SaveRulesToParameterFile(objSettingsFile As XmlSettingsFileAccessor, strSectionName As String, udtRules As IList(Of udtRuleDefinitionType))

        Dim intRuleNumber As Integer
        Dim strRuleBase As String

        If udtRules Is Nothing OrElse udtRules.Count <= 0 Then
            objSettingsFile.SetParam(strSectionName, XML_OPTION_ENTRY_RULE_COUNT, 0)
        Else
            objSettingsFile.SetParam(strSectionName, XML_OPTION_ENTRY_RULE_COUNT, udtRules.Count)

            For intRuleNumber = 1 To udtRules.Count
                strRuleBase = "Rule" & intRuleNumber.ToString

                With udtRules(intRuleNumber - 1)
                    objSettingsFile.SetParam(strSectionName, strRuleBase & "MatchRegEx", .MatchRegEx)
                    objSettingsFile.SetParam(strSectionName, strRuleBase & "MatchIndicatesProblem", .MatchIndicatesProblem)
                    objSettingsFile.SetParam(strSectionName, strRuleBase & "MessageWhenProblem", .MessageWhenProblem)
                    objSettingsFile.SetParam(strSectionName, strRuleBase & "Severity", .Severity)
                    objSettingsFile.SetParam(strSectionName, strRuleBase & "DisplayMatchAsExtraInfo", .DisplayMatchAsExtraInfo)
                End With

            Next intRuleNumber
        End If

    End Sub

    Public Function SaveSettingsToParameterFile(strParameterFilePath As String) As Boolean Implements IValidateFastaFile.SaveParameterSettingsToParameterFile
        ' Save a model parameter file

        Dim srOutFile As StreamWriter
        Dim objSettingsFile As New XmlSettingsFileAccessor

        Try

            If strParameterFilePath Is Nothing OrElse strParameterFilePath.Length = 0 Then
                ' No parameter file specified; do not save the settings
                Return True
            End If

            If Not File.Exists(strParameterFilePath) Then
                ' Need to generate a blank XML settings file

                srOutFile = New StreamWriter(strParameterFilePath, False)

                srOutFile.WriteLine("<?xml version=""1.0"" encoding=""UTF-8""?>")
                srOutFile.WriteLine("<sections>")
                srOutFile.WriteLine("  <section name=""" & XML_SECTION_OPTIONS & """>")
                srOutFile.WriteLine("  </section>")
                srOutFile.WriteLine("</sections>")

                srOutFile.Close()
            End If

            objSettingsFile.LoadSettings(strParameterFilePath)

            ' Save the general settings

            objSettingsFile.SetParam(XML_SECTION_OPTIONS, "AddMissingLinefeedAtEOF", Me.OptionSwitch(IValidateFastaFile.SwitchOptions.AddMissingLinefeedatEOF))
            objSettingsFile.SetParam(XML_SECTION_OPTIONS, "AllowAsteriskInResidues", Me.OptionSwitch(IValidateFastaFile.SwitchOptions.AllowAsteriskInResidues))
            objSettingsFile.SetParam(XML_SECTION_OPTIONS, "AllowDashInResidues", Me.OptionSwitch(IValidateFastaFile.SwitchOptions.AllowDashInResidues))

            objSettingsFile.SetParam(XML_SECTION_OPTIONS, "CheckForDuplicateProteinNames", Me.OptionSwitch(IValidateFastaFile.SwitchOptions.CheckForDuplicateProteinNames))
            objSettingsFile.SetParam(XML_SECTION_OPTIONS, "CheckForDuplicateProteinSequences", Me.OptionSwitch(IValidateFastaFile.SwitchOptions.CheckForDuplicateProteinSequences))
            objSettingsFile.SetParam(XML_SECTION_OPTIONS, "SaveProteinSequenceHashInfoFiles", Me.OptionSwitch(IValidateFastaFile.SwitchOptions.SaveProteinSequenceHashInfoFiles))
            objSettingsFile.SetParam(XML_SECTION_OPTIONS, "SaveBasicProteinHashInfoFile", Me.OptionSwitch(IValidateFastaFile.SwitchOptions.SaveBasicProteinHashInfoFile))

            objSettingsFile.SetParam(XML_SECTION_OPTIONS, "MaximumFileErrorsToTrack", Me.MaximumFileErrorsToTrack)
            objSettingsFile.SetParam(XML_SECTION_OPTIONS, "MinimumProteinNameLength", Me.MinimumProteinNameLength)
            objSettingsFile.SetParam(XML_SECTION_OPTIONS, "MaximumProteinNameLength", Me.MaximumProteinNameLength)
            objSettingsFile.SetParam(XML_SECTION_OPTIONS, "MaximumResiduesPerLine", Me.MaximumResiduesPerLine)

            objSettingsFile.SetParam(XML_SECTION_OPTIONS, "WarnBlankLinesBetweenProteins", Me.OptionSwitch(IValidateFastaFile.SwitchOptions.WarnBlankLinesBetweenProteins))
            objSettingsFile.SetParam(XML_SECTION_OPTIONS, "WarnLineStartsWithSpace", Me.OptionSwitch(IValidateFastaFile.SwitchOptions.WarnLineStartsWithSpace))
            objSettingsFile.SetParam(XML_SECTION_OPTIONS, "OutputToStatsFile", Me.OptionSwitch(IValidateFastaFile.SwitchOptions.OutputToStatsFile))
            objSettingsFile.SetParam(XML_SECTION_OPTIONS, "NormalizeFileLineEndCharacters", Me.OptionSwitch(IValidateFastaFile.SwitchOptions.NormalizeFileLineEndCharacters))


            objSettingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "GenerateFixedFASTAFile", Me.OptionSwitch(IValidateFastaFile.SwitchOptions.GenerateFixedFASTAFile))
            objSettingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "SplitOutMultipleRefsInProteinName", Me.OptionSwitch(IValidateFastaFile.SwitchOptions.SplitOutMultipleRefsInProteinName))

            objSettingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "RenameDuplicateNameProteins", Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaRenameDuplicateNameProteins))
            objSettingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "KeepDuplicateNamedProteins", Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaKeepDuplicateNamedProteins))

            objSettingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ConsolidateDuplicateProteinSeqs", Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaConsolidateDuplicateProteinSeqs))
            objSettingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ConsolidateDupsIgnoreILDiff", Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaConsolidateDupsIgnoreILDiff))

            objSettingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "TruncateLongProteinNames", Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaTruncateLongProteinNames))
            objSettingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "SplitOutMultipleRefsForKnownAccession", Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaSplitOutMultipleRefsForKnownAccession))
            objSettingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "WrapLongResidueLines", Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaWrapLongResidueLines))
            objSettingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "RemoveInvalidResidues", Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaRemoveInvalidResidues))


            objSettingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "LongProteinNameSplitChars", Me.LongProteinNameSplitChars)
            objSettingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ProteinNameInvalidCharsToRemove", Me.ProteinNameInvalidCharsToRemove)
            objSettingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ProteinNameFirstRefSepChars", Me.ProteinNameFirstRefSepChars)
            objSettingsFile.SetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ProteinNameSubsequentRefSepChars", Me.ProteinNameSubsequentRefSepChars)


            ' Save the rules
            SaveRulesToParameterFile(objSettingsFile, XML_SECTION_FASTA_HEADER_LINE_RULES, mHeaderLineRules)
            SaveRulesToParameterFile(objSettingsFile, XML_SECTION_FASTA_PROTEIN_NAME_RULES, mProteinNameRules)
            SaveRulesToParameterFile(objSettingsFile, XML_SECTION_FASTA_PROTEIN_DESCRIPTION_RULES, mProteinDescriptionRules)
            SaveRulesToParameterFile(objSettingsFile, XML_SECTION_FASTA_PROTEIN_SEQUENCE_RULES, mProteinSequenceRules)

            ' Commit the new settings to disk
            objSettingsFile.SaveSettings()

            ' Need to re-open the parameter file and replace instances of "&gt;" with ">" and "&lt;" with "<"
            ReplaceXMLCodesWithText(strParameterFilePath)

        Catch ex As Exception
            If MyBase.ShowMessages Then
                ShowErrorMessage("Error in SaveSettingsToParameterFile:" & ex.Message)
            Else
                Throw New Exception("Error in SaveSettingsToParameterFile", ex)
            End If
            Return False
        End Try

        Return True

    End Function

    Private Function SearchRulesForID(
      udtRules As IList(Of udtRuleDefinitionType),
      intErrorMessageCode As Integer,
      <Out()> ByRef strMessage As String) As Boolean

        If Not udtRules Is Nothing Then
            For intIndex = 0 To udtRules.Count - 1
                If udtRules(intIndex).CustomRuleID = intErrorMessageCode Then
                    strMessage = udtRules(intIndex).MessageWhenProblem
                    Return True
                End If
            Next intIndex
        End If

        strMessage = Nothing
        Return False

    End Function

    ''' <summary>
    ''' Updates the validation rules using the current options
    ''' </summary>
    ''' <remarks>Call this function after setting new options using SetOptionSwitch</remarks>
    Public Sub SetDefaultRules() Implements IValidateFastaFile.SetDefaultRules

        Me.ClearAllRules()

        ' For the rules, severity level 1 to 4 is warning; severity 5 or higher is an error

        ' Header line errors
        Me.SetRule(IValidateFastaFile.RuleTypes.HeaderLine, "^>[ \t]*$", True, "Line starts with > but does not contain a protein name", DEFAULT_ERROR_SEVERITY)
        Me.SetRule(IValidateFastaFile.RuleTypes.HeaderLine, "^>[ \t].+", True, "Space or tab found directly after the > symbol", DEFAULT_ERROR_SEVERITY)

        ' Header line warnings
        Me.SetRule(IValidateFastaFile.RuleTypes.HeaderLine, "^>[^ \t]+[ \t]*$", True, MESSAGE_TEXT_PROTEIN_DESCRIPTION_MISSING, DEFAULT_WARNING_SEVERITY)
        Me.SetRule(IValidateFastaFile.RuleTypes.HeaderLine, "^>[^ \t]+\t", True, "Protein name is separated from the protein description by a tab", DEFAULT_WARNING_SEVERITY)

        ' Protein Name error characters
        Dim allowedChars = "A-Za-z0-9.\-_:,\|/()\[\]\=\+#"

        If mAllowAllSymbolsInProteinNames Then
            allowedChars &= "!@$%^&*<>?,\\"
        End If

        Dim allowedCharsMatchSpec = "[^" & allowedChars & "]"

        Me.SetRule(IValidateFastaFile.RuleTypes.ProteinName, allowedCharsMatchSpec, True, "Protein name contains invalid characters", DEFAULT_ERROR_SEVERITY, True)

        ' Protein name warnings

        ' Note that .*? changes .* from being greedy to being lazy
        Me.SetRule(IValidateFastaFile.RuleTypes.ProteinName, "[:|].*?[:|;].*?[:|;]", True, "Protein name contains 3 or more vertical bars", DEFAULT_WARNING_SEVERITY + 1, True)

        If Not mAllowAllSymbolsInProteinNames Then
            Me.SetRule(IValidateFastaFile.RuleTypes.ProteinName, "[/()\[\],]", True, "Protein name contains undesirable characters", DEFAULT_WARNING_SEVERITY, True)
        End If

        ' Protein description warnings
        Me.SetRule(IValidateFastaFile.RuleTypes.ProteinDescription, """", True, "Protein description contains a quotation mark", DEFAULT_WARNING_SEVERITY)
        Me.SetRule(IValidateFastaFile.RuleTypes.ProteinDescription, "\t", True, "Protein description contains a tab character", DEFAULT_WARNING_SEVERITY)
        Me.SetRule(IValidateFastaFile.RuleTypes.ProteinDescription, "\\/", True, "Protein description contains an escaped slash: \/", DEFAULT_WARNING_SEVERITY)
        Me.SetRule(IValidateFastaFile.RuleTypes.ProteinDescription, "[\x00-\x08\x0E-\x1F]", True, "Protein description contains an escape code character", DEFAULT_ERROR_SEVERITY)
        Me.SetRule(IValidateFastaFile.RuleTypes.ProteinDescription, ".{900,}", True, MESSAGE_TEXT_PROTEIN_DESCRIPTION_TOO_LONG, DEFAULT_WARNING_SEVERITY + 1, False)

        ' Protein sequence errors
        Me.SetRule(IValidateFastaFile.RuleTypes.ProteinSequence, "[ \t]", True, "A space or tab was found in the residues", DEFAULT_ERROR_SEVERITY)

        If Not mAllowAsteriskInResidues Then
            Me.SetRule(IValidateFastaFile.RuleTypes.ProteinSequence, "\*", True, MESSAGE_TEXT_ASTERISK_IN_RESIDUES, DEFAULT_ERROR_SEVERITY)
        End If

        If Not mAllowDashInResidues Then
            Me.SetRule(IValidateFastaFile.RuleTypes.ProteinSequence, "\-", True, MESSAGE_TEXT_DASH_IN_RESIDUES, DEFAULT_ERROR_SEVERITY)
        End If

        ' Note: we look for a space, tab, asterisk, and dash with separate rules (defined above)
        ' Thus they are "allowed" by this RegEx, even though we may flag them as warnings with a different RegEx
        ' We look for non-standard amino acids with warning rules (defined below)
        Me.SetRule(IValidateFastaFile.RuleTypes.ProteinSequence, "[^A-Z \t\*\-]", True, "Invalid residues found", DEFAULT_ERROR_SEVERITY, True)

        ' Protein residue warnings
        ' MSGF+ treats these residues as stop characters(meaning no identified peptide will ever contain B, J, O, U, X, or Z)

        ' SEQUEST uses mass 114.53494 for B (average of N and D)
        Me.SetRule(IValidateFastaFile.RuleTypes.ProteinSequence, "B", True, "Residues line contains B (non-standard amino acid for N or D)", DEFAULT_WARNING_SEVERITY - 1)

        ' Unsupported by SEQUEST
        Me.SetRule(IValidateFastaFile.RuleTypes.ProteinSequence, "J", True, "Residues line contains J (non-standard amino acid)", DEFAULT_WARNING_SEVERITY - 1)

        ' SEQUEST uses mass 114.07931 for O
        Me.SetRule(IValidateFastaFile.RuleTypes.ProteinSequence, "O", True, "Residues line contains O (non-standard amino acid, ornithine)", DEFAULT_WARNING_SEVERITY - 1)

        ' Unsupported by SEQUEST
        Me.SetRule(IValidateFastaFile.RuleTypes.ProteinSequence, "U", True, "Residues line contains U (non-standard amino acid, selenocysteine)", DEFAULT_WARNING_SEVERITY)

        ' SEQUEST uses mass 113.08406 for X (same as L and I)
        Me.SetRule(IValidateFastaFile.RuleTypes.ProteinSequence, "X", True, "Residues line contains X (non-standard amino acid for L or I)", DEFAULT_WARNING_SEVERITY - 1)

        ' SEQUEST uses mass 128.55059 for Z (average of Q and E)
        Me.SetRule(IValidateFastaFile.RuleTypes.ProteinSequence, "Z", True, "Residues line contains Z (non-standard amino acid for Q or E)", DEFAULT_WARNING_SEVERITY - 1)

    End Sub

    Private Sub SetLocalErrorCode(eNewErrorCode As IValidateFastaFile.eValidateFastaFileErrorCodes)
        SetLocalErrorCode(eNewErrorCode, False)
    End Sub

    Private Sub SetLocalErrorCode(
      eNewErrorCode As IValidateFastaFile.eValidateFastaFileErrorCodes,
      blnLeaveExistingErrorCodeUnchanged As Boolean)

        If blnLeaveExistingErrorCodeUnchanged AndAlso mLocalErrorCode <> IValidateFastaFile.eValidateFastaFileErrorCodes.NoError Then
            ' An error code is already defined; do not change it
        Else
            mLocalErrorCode = eNewErrorCode

            If eNewErrorCode = IValidateFastaFile.eValidateFastaFileErrorCodes.NoError Then
                If MyBase.ErrorCode = eProcessFilesErrorCodes.LocalizedError Then
                    MyBase.SetBaseClassErrorCode(eProcessFilesErrorCodes.NoError)
                End If
            Else
                MyBase.SetBaseClassErrorCode(eProcessFilesErrorCodes.LocalizedError)
            End If
        End If

    End Sub

    Private Sub SetRule(
      ruleType As IValidateFastaFile.RuleTypes,
      regexToMatch As String,
      doesMatchIndicateProblem As Boolean,
      problemReturnMessage As String,
      severityLevel As Short) Implements IValidateFastaFile.SetRule

        Me.SetRule(
         ruleType, regexToMatch,
         doesMatchIndicateProblem,
         problemReturnMessage,
         severityLevel, False)

    End Sub

    Private Sub SetRule(
      ruleType As IValidateFastaFile.RuleTypes,
      regexToMatch As String,
      doesMatchIndicateProblem As Boolean,
      problemReturnMessage As String,
      severityLevel As Short,
      displayMatchAsExtraInfo As Boolean) Implements IValidateFastaFile.SetRule

        Select Case ruleType
            Case IValidateFastaFile.RuleTypes.HeaderLine
                SetRule(Me.mHeaderLineRules, regexToMatch, doesMatchIndicateProblem, problemReturnMessage, severityLevel, displayMatchAsExtraInfo)
            Case IValidateFastaFile.RuleTypes.ProteinDescription
                SetRule(Me.mProteinDescriptionRules, regexToMatch, doesMatchIndicateProblem, problemReturnMessage, severityLevel, displayMatchAsExtraInfo)
            Case IValidateFastaFile.RuleTypes.ProteinName
                SetRule(Me.mProteinNameRules, regexToMatch, doesMatchIndicateProblem, problemReturnMessage, severityLevel, displayMatchAsExtraInfo)
            Case IValidateFastaFile.RuleTypes.ProteinSequence
                SetRule(Me.mProteinSequenceRules, regexToMatch, doesMatchIndicateProblem, problemReturnMessage, severityLevel, displayMatchAsExtraInfo)
        End Select

    End Sub

    Private Sub SetRule(
       ByRef udtRules() As udtRuleDefinitionType,
       strMatchRegEx As String,
       blnMatchIndicatesProblem As Boolean,
       strMessageWhenProblem As String,
       bytSeverity As Short,
       blnDisplayMatchAsExtraInfo As Boolean)

        If udtRules Is Nothing OrElse udtRules.Length = 0 Then
            ReDim udtRules(0)
        Else
            ReDim Preserve udtRules(udtRules.Length)
        End If

        With udtRules(udtRules.Length - 1)
            .MatchRegEx = strMatchRegEx
            .MatchIndicatesProblem = blnMatchIndicatesProblem
            .MessageWhenProblem = strMessageWhenProblem
            .Severity = bytSeverity
            .DisplayMatchAsExtraInfo = blnDisplayMatchAsExtraInfo
            .CustomRuleID = mMasterCustomRuleID
        End With

        mMasterCustomRuleID += 1

    End Sub

    Private Sub ShowExceptionStackTrace(callingFunction As String, ex As Exception)

        Console.WriteLine()

        Dim stackTraceInfo = clsStackTraceFormatter.GetExceptionStackTraceMultiLine(ex)
        If (String.IsNullOrEmpty(callingFunction)) Then
            Console.WriteLine(stackTraceInfo)
        Else
            Console.WriteLine(callingFunction & ": " & stackTraceInfo)
        End If

    End Sub

    Private Function SortFile(fiProteinHashFile As FileInfo, sortColumnIndex As Integer, sortedFilePath As String) As Boolean

        Dim sortUtility = New FlexibleFileSortUtility.TextFileSorter

        mLastSortUtilityProgress = DateTime.UtcNow
        mSortUtilityErrorMessage = String.Empty

        sortUtility.WorkingDirectoryPath = fiProteinHashFile.Directory.FullName
        sortUtility.HasHeaderLine = True
        sortUtility.ColumnDelimiter = ControlChars.Tab
        sortUtility.MaxFileSizeMBForInMemorySort = 250
        sortUtility.ChunkSizeMB = 250
        sortUtility.SortColumn = sortColumnIndex + 1
        sortUtility.SortColumnIsNumeric = False

        ' The sort utility uses CompareOrdinal (StringComparison.Ordinal)
        sortUtility.IgnoreCase = False

        AddHandler sortUtility.ProgressChanged, AddressOf mSortUtility_ProgressChanged
        AddHandler sortUtility.ErrorEvent, AddressOf mSortUtility_ErrorEvent
        AddHandler sortUtility.WarningEvent, AddressOf mSortUtility_WarningEvent
        AddHandler sortUtility.MessageEvent, AddressOf mSortUtility_MessageEvent

        Dim success = sortUtility.SortFile(fiProteinHashFile.FullName, sortedFilePath)

        If success Then
            Console.WriteLine()
            Return True
        End If

        If String.IsNullOrWhiteSpace(mSortUtilityErrorMessage) Then
            ShowErrorMessage("Unknown error sorting " & fiProteinHashFile.Name)
        Else
            ShowErrorMessage("Sort error: " & mSortUtilityErrorMessage)
        End If

        Console.WriteLine()
        Return False

    End Function

    Private Sub SplitFastaProteinHeaderLine(
      strHeaderLine As String,
      <Out> ByRef strProteinName As String,
      <Out> ByRef strProteinDescription As String,
      <Out> ByRef intDescriptionStartIndex As Integer)

        strProteinDescription = String.Empty
        intDescriptionStartIndex = 0

        ' Make sure the protein name and description are valid
        ' Find the first space and/or tab
        Dim intSpaceIndex = GetBestSpaceIndex(strHeaderLine)

        ' At this point, intSpaceIndex will contain the location of the space or tab separating the protein name and description
        ' However, if the space or tab is directly after the > sign, then we cannot continue (if this is the case, then intSpaceIndex will be 1)
        If intSpaceIndex > 1 Then

            strProteinName = strHeaderLine.Substring(1, intSpaceIndex - 1)
            strProteinDescription = strHeaderLine.Substring(intSpaceIndex + 1)
            intDescriptionStartIndex = intSpaceIndex

        Else
            ' Line does not contain a description
            If intSpaceIndex <= 0 Then
                If strHeaderLine.Trim.Length <= 1 Then
                    strProteinName = String.Empty
                Else
                    ' The line contains a protein name, but not a description
                    strProteinName = strHeaderLine.Substring(1)
                End If
            Else
                ' Space or tab found directly after the > symbol
                strProteinName = String.Empty
            End If
        End If

    End Sub

    Private Function VerifyLinefeedAtEOF(strInputFilePath As String, blnAddCrLfIfMissing As Boolean) As Boolean
        Dim intByte As Integer
        Dim bytOneByte As Byte

        Dim blnNeedToAddCrLf As Boolean
        Dim blnSuccess As Boolean

        Try
            ' Open the input file and validate that the final characters are CrLf, simply CR, or simply LF
            Using fsInFile = New FileStream(strInputFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)

                If fsInFile.Length > 2 Then
                    fsInFile.Seek(-1, SeekOrigin.End)

                    intByte = fsInFile.ReadByte()

                    If intByte = 10 Or intByte = 13 Then
                        ' File ends in a linefeed or carriage return character; that's good
                        blnNeedToAddCrLf = False
                    Else
                        blnNeedToAddCrLf = True
                    End If
                End If

                If blnNeedToAddCrLf Then
                    If blnAddCrLfIfMissing Then
                        ShowMessage("Appending CrLf return to: " & Path.GetFileName(strInputFilePath))
                        bytOneByte = CType(13, Byte)
                        fsInFile.WriteByte(bytOneByte)

                        bytOneByte = CType(10, Byte)
                        fsInFile.WriteByte(bytOneByte)
                    End If
                End If

            End Using

            blnSuccess = True

        Catch ex As Exception
            SetLocalErrorCode(IValidateFastaFile.eValidateFastaFileErrorCodes.ErrorVerifyingLinefeedAtEOF)
            blnSuccess = False
        End Try

        Return blnSuccess

    End Function

    Private Sub WriteCachedProtein(
      strCachedProteinName As String,
      strCachedProteinDescription As String,
      swConsolidatedFastaOut As TextWriter,
      oProteinSeqHashInfo As IList(Of clsProteinHashInfo),
      sbCachedProteinResidueLines As StringBuilder,
      sbCachedProteinResidues As StringBuilder,
      blnConsolidateDuplicateProteinSeqsInFasta As Boolean,
      blnConsolidateDupsIgnoreILDiff As Boolean,
      lstProteinNameFirst As clsNestedStringDictionary(Of Integer),
      lstDuplicateProteinList As clsNestedStringDictionary(Of String),
      intLineCount As Integer,
      lstProteinsWritten As clsNestedStringDictionary(Of Integer))

        Static reAdditionalProtein As Regex = New Regex("(.+)-[a-z]\d*", RegexOptions.Compiled)

        Dim strMasterProteinName As String = String.Empty
        Dim strMasterProteinInfo As String

        Dim strProteinHash As String
        Dim strLineOut As String = String.Empty
        Dim reMatch As Match

        Dim blnKeepProtein As Boolean
        Dim blnSkipDupProtein As Boolean

        Dim lstAdditionalProteinNames = New List(Of String)

        Dim intSeqIndex As Integer
        If lstProteinNameFirst.TryGetValue(strCachedProteinName, intSeqIndex) Then
            ' strCachedProteinName was found in lstProteinNameFirst

            Dim cachedSeqIndex As Integer
            If lstProteinsWritten.TryGetValue(strCachedProteinName, cachedSeqIndex) Then
                If mFixedFastaOptions.KeepDuplicateNamedProteinsUnlessMatchingSequence Then
                    ' Keep this protein if its sequence hash differs from the first protein with this name
                    strProteinHash = ComputeProteinHash(sbCachedProteinResidues, blnConsolidateDupsIgnoreILDiff)
                    If oProteinSeqHashInfo(intSeqIndex).SequenceHash <> strProteinHash Then
                        RecordFastaFileWarning(intLineCount, 1, strCachedProteinName, eMessageCodeConstants.DuplicateProteinNameRetained)
                        blnKeepProtein = True
                    Else
                        blnKeepProtein = False
                    End If
                Else
                    blnKeepProtein = False
                End If
            Else
                blnKeepProtein = True
                lstProteinsWritten.Add(strCachedProteinName, intSeqIndex)
            End If

            If blnKeepProtein AndAlso intSeqIndex >= 0 Then
                If oProteinSeqHashInfo(intSeqIndex).AdditionalProteins.Count > 0 Then
                    ' The protein has duplicate proteins
                    ' Construct a list of the duplicate protein names

                    lstAdditionalProteinNames.Clear()
                    For Each strAdditionalProtein As String In oProteinSeqHashInfo(intSeqIndex).AdditionalProteins
                        ' Add the additional protein name if it is not of the form "BaseName-b", "BaseName-c", etc.
                        blnSkipDupProtein = False

                        If strAdditionalProtein Is Nothing Then
                            blnSkipDupProtein = True
                        Else
                            If strAdditionalProtein.ToLower() = strCachedProteinName.ToLower() Then
                                ' Names match; do not add to the list
                                blnSkipDupProtein = True
                            Else
                                ' Check whether strAdditionalProtein looks like one of the following
                                ' ProteinX-b
                                ' ProteinX-a2
                                ' ProteinX-d3
                                reMatch = reAdditionalProtein.Match(strAdditionalProtein)

                                If reMatch.Success Then

                                    If strCachedProteinName.ToLower() = reMatch.Groups(1).Value.ToLower() Then
                                        ' Base names match; do not add to the list
                                        ' For example, ProteinX and ProteinX-b
                                        blnSkipDupProtein = True
                                    End If
                                End If

                            End If
                        End If

                        If Not blnSkipDupProtein Then
                            lstAdditionalProteinNames.Add(strAdditionalProtein)
                        End If
                    Next

                    If lstAdditionalProteinNames.Count > 0 AndAlso blnConsolidateDuplicateProteinSeqsInFasta Then
                        ' Append the duplicate protein names to the description
                        ' However, do not let the description get over 7995 characters in length
                        Dim updatedDescription = strCachedProteinDescription & "; Duplicate proteins: " & FlattenArray(lstAdditionalProteinNames, ","c)
                        If updatedDescription.Length > MAX_PROTEIN_DESCRIPTION_LENGTH Then
                            updatedDescription = updatedDescription.Substring(0, MAX_PROTEIN_DESCRIPTION_LENGTH - 3) & "..."
                        End If

                        strLineOut = ConstructFastaHeaderLine(strCachedProteinName, updatedDescription)
                    End If
                End If
            End If
        Else
            blnKeepProtein = False
            mFixedFastaStats.DuplicateSequenceProteinsSkipped += 1

            If Not lstDuplicateProteinList.TryGetValue(strCachedProteinName, strMasterProteinName) Then
                strMasterProteinInfo = "same as ??"
            Else
                strMasterProteinInfo = "same as " & strMasterProteinName
            End If
            RecordFastaFileWarning(intLineCount, 0, strCachedProteinName, eMessageCodeConstants.ProteinRemovedSinceDuplicateSequence, strMasterProteinInfo, String.Empty)
        End If

        If blnKeepProtein Then
            If String.IsNullOrEmpty(strLineOut) Then
                strLineOut = ConstructFastaHeaderLine(strCachedProteinName, strCachedProteinDescription)
            End If
            swConsolidatedFastaOut.WriteLine(strLineOut)
            swConsolidatedFastaOut.Write(sbCachedProteinResidueLines.ToString())
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

    Private Sub mSortUtility_ErrorEvent(sender As Object, e As FlexibleFileSortUtility.MessageEventArgs)
        mSortUtilityErrorMessage = e.Message
        ShowErrorMessage(e.Message)
    End Sub

    Private Sub mSortUtility_MessageEvent(sender As Object, e As FlexibleFileSortUtility.MessageEventArgs)
        ' The FlexibleFileSortUtility DLL already displays these messages at the console; no need to repeat them
    End Sub

    Private Sub mSortUtility_ProgressChanged(sender As Object, e As FlexibleFileSortUtility.ProgressChangedEventArgs)
        If ShowMessages AndAlso DateTime.UtcNow.Subtract(mLastSortUtilityProgress).TotalSeconds >= 15 Then
            mLastSortUtilityProgress = DateTime.UtcNow
            Console.WriteLine(e.taskDescription & ": " & e.percentComplete.ToString("0.0") & "% complete")
        End If
    End Sub

    Private Sub mSortUtility_WarningEvent(sender As Object, e As FlexibleFileSortUtility.MessageEventArgs)
        ShowMessage("Sort tool warning: " & e.Message)
    End Sub
#End Region

    ' IComparer class to allow comparison of udtMsgInfoType items
    Private Class ErrorInfoComparerClass
        Implements IComparer

        Public Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare

            Dim udtErrorInfo1, udtErrorInfo2 As IValidateFastaFile.udtMsgInfoType

            udtErrorInfo1 = CType(x, IValidateFastaFile.udtMsgInfoType)
            udtErrorInfo2 = CType(y, IValidateFastaFile.udtMsgInfoType)

            If udtErrorInfo1.MessageCode > udtErrorInfo2.MessageCode Then
                Return 1
            ElseIf udtErrorInfo1.MessageCode < udtErrorInfo2.MessageCode Then
                Return -1
            Else
                If udtErrorInfo1.LineNumber > udtErrorInfo2.LineNumber Then
                    Return 1
                ElseIf udtErrorInfo1.LineNumber < udtErrorInfo2.LineNumber Then
                    Return -1
                Else
                    If udtErrorInfo1.ColNumber > udtErrorInfo2.ColNumber Then
                        Return 1
                    ElseIf udtErrorInfo1.ColNumber < udtErrorInfo2.ColNumber Then
                        Return -1
                    Else
                        Return 0
                    End If
                End If
            End If

        End Function
    End Class

End Class
