Option Strict On

' This class will read a protein fasta file and validate its contents
' 
' -------------------------------------------------------------------------------
' Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
' Program started March 21, 2005

' E-mail: matthew.monroe@pnl.gov or matt@alchemistmatt.com
' Website: http://ncrr.pnl.gov/ or http://www.sysbio.org/resources/staff/
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

Public Class clsValidateFastaFile
    Inherits clsProcessFilesBaseClass
    Implements IValidateFastaFile

    Public Sub New()
		MyBase.mFileDate = "February 2, 2012"
        InitializeLocalVariables()
    End Sub

    Public Sub New(ByVal ParameterFilePath As String)
        Me.New()
        Me.LoadParameterFileSettings(ParameterFilePath)
    End Sub


#Region "Constants and Enums"
    Protected Const DEFAULT_MINIMUM_PROTEIN_NAME_LENGTH As Integer = 3
	Public Const DEFAULT_MAXIMUM_PROTEIN_NAME_LENGTH As Integer = 34
    Protected Const DEFAULT_MAXIMUM_RESIDUES_PER_LINE As Integer = 120

    Public Const DEFAULT_PROTEIN_LINE_START_CHAR As Char = ">"c
    Public Const DEFAULT_LONG_PROTEIN_NAME_SPLIT_CHAR As Char = "|"c
    Public Const DEFAULT_PROTEIN_NAME_FIRST_REF_SEP_CHARS As String = ":|"
    Public Const DEFAULT_PROTEIN_NAME_SUBSEQUENT_REF_SEP_CHARS As String = ":|;"

    Private Const INVALID_PROTEIN_NAME_CHAR_REPLACEMENT As Char = "_"c

    Private CUSTOM_RULE_ID_START As Integer = 1000
    Private DEFAULT_CONTEXT_LENGTH As Integer = 13

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
    End Enum

#End Region

#Region "Structures"

    Protected Structure udtErrorStatsType
        Public MessageCode As Integer               ' Note: Custom rules start with message code CUSTOM_RULE_ID_START
        Public CountSpecified As Integer
        Public CountUnspecified As Integer
    End Structure

    Private Structure udtItemSummaryIndexedType
        Public ErrorStatsCount As Integer
        Public ErrorStats() As udtErrorStatsType        ' Note: This array ranges from 0 to .ErrorStatsCount since it is Dimmed with extra space
        Public htMessageCodeToArrayIndex As Hashtable
    End Structure


    Public Structure udtRuleDefinitionType
        Public MatchRegEx As String
        Public MatchIndicatesProblem As Boolean     ' True means text matching the RegEx means a problem; false means if text doesn't match the RegEx, then that means a problem
        Public MessageWhenProblem As String         ' Message to display if a problem is present
        Public Severity As Short                    ' 0 is lowest severity, 9 is highest severity; value >= 5 means error
        Public DisplayMatchAsExtraInfo As Boolean   ' If true, then the matching text is stored as the context info
        Public CustomRuleID As Integer              ' This value is auto-assigned
    End Structure

    Protected Structure udtRuleDefinitionExtendedType
        Public RuleDefinition As udtRuleDefinitionType
        Public reRule As System.Text.RegularExpressions.Regex
        Public Valid As Boolean
    End Structure

    Protected Structure udtProteinHashInfoType
        Public SequenceHash As String
        Public SequenceLength As Integer
        Public SequenceStart As String          ' The first 20 residues of the protein sequence
        Public ProteinNameFirst As String
        Public AdditionalProteins() As String   ' Will contain .ProteinCount-1 items since .ProteinNameFirst is not stored here
        Public ProteinCount As Integer          ' 1 if only one protein matches this sequence
    End Structure

    Protected Structure udtFixedFastaOptionsType
        Public SplitOutMultipleRefsInProteinName As Boolean
        Public SplitOutMultipleRefsForKnownAccession As Boolean
        Public LongProteinNameSplitChars As Char()
        Public ProteinNameInvalidCharsToRemove As Char()
        Public RenameProteinsWithDuplicateNames As Boolean
        Public ConsolidateProteinsWithDuplicateSeqs As Boolean
        Public ConsolidateDupsIgnoreILDiff As Boolean
        Public TruncateLongProteinNames As Boolean
        Public WrapLongResidueLines As Boolean
        Public RemoveInvalidResidues As Boolean
    End Structure

    Protected Structure udtFixedFastaStatsType
        Public TruncatedProteinNameCount As Integer
        Public UpdatedResidueLines As Integer
        Public ProteinNamesInvalidCharsReplaced As Integer
        Public ProteinNamesMultipleRefsRemoved As Integer
        Public DuplicateNameProteinsSkipped As Integer
        Public DuplicateNameProteinsRenamed As Integer
        Public DuplicateSequenceProteinsSkipped As Integer
    End Structure

    Protected Structure udtProteinNameTruncationRegex
        Public reMatchIPI As System.Text.RegularExpressions.Regex
        Public reMatchGI As System.Text.RegularExpressions.Regex
        Public reMatchJGI As System.Text.RegularExpressions.Regex
        Public reMatchJGIBaseAndID As System.Text.RegularExpressions.Regex
        Public reMatchGeneric As System.Text.RegularExpressions.Regex
        Public reMatchDoubleBarOrColonAndBar As System.Text.RegularExpressions.Regex
    End Structure

#End Region

#Region "Classwide Variables"

    Protected mFastaFilePath As String
    Protected mLineCount As Integer
    Protected mProteinCount As Integer
    Protected mResidueCount As Long

    Protected mFixedFastaStats As udtFixedFastaStatsType

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

    Protected mAddMissingLinefeedAtEOF As Boolean
    Protected mCheckForDuplicateProteinNames As Boolean
    Protected mCheckForDuplicateProteinSequences As Boolean     ' This will be set to True if mSaveProteinSequenceHashInfoFiles = True or if mFixedFastaConsolidateProteinsWithDuplicateSeqs = True

    Protected mMaximumFileErrorsToTrack As Integer        ' This is the maximum # of errors per type to track
    Protected mMinimumProteinNameLength As Integer
    Protected mMaximumProteinNameLength As Integer
    Protected mMaximumResiduesPerLine As Integer

    Protected mFixedFastaOptions As udtFixedFastaOptionsType     ' Used if mGenerateFixedFastaFile = True

    Protected mOutputToStatsFile As Boolean
    Protected mStatsFilePath As String
    Protected mGenerateFixedFastaFile As Boolean
    Protected mSaveProteinSequenceHashInfoFiles As Boolean

    ' When true, then creates a text file that will contain the protein name and sequence hash for each protein; 
    '  this option will not store protein names and/or hashes in memory, and is thus useful for processing 
    '  huge .Fasta files to determine duplicate proteins
    Protected mSaveBasicProteinHashInfoFile As Boolean

    Protected mProteinLineStartChar As Char

    Protected mAllowAsteriskInResidues As Boolean
    Protected mAllowDashInResidues As Boolean
    Protected mWarnBlankLinesBetweenProteins As Boolean
    Protected mWarnLineStartsWithSpace As Boolean
    Protected mNormalizeFileLineEndCharacters As Boolean

    Private mLocalErrorCode As IValidateFastaFile.eValidateFastaFileErrorCodes
#End Region

#Region "Properties"

    Public Property OptionSwitch(ByVal SwitchName As IValidateFastaFile.SwitchOptions) As Boolean _
        Implements IValidateFastaFile.OptionSwitches
        Get
            Return Me.GetOptionSwitchValue(SwitchName)
        End Get
        Set(ByVal Value As Boolean)
            Me.SetOptionSwitch(SwitchName, Value)
        End Set
    End Property

    Public Sub SetOptionSwitch( _
        ByVal SwitchName As IValidateFastaFile.SwitchOptions, _
        ByVal State As Boolean)

        Select Case SwitchName
            Case IValidateFastaFile.SwitchOptions.AddMissingLinefeedatEOF
                Me.mAddMissingLinefeedAtEOF = State
            Case IValidateFastaFile.SwitchOptions.AllowAsteriskInResidues
                Me.mAllowAsteriskInResidues = State
            Case IValidateFastaFile.SwitchOptions.CheckForDuplicateProteinNames
                Me.mCheckForDuplicateProteinNames = State
            Case IValidateFastaFile.SwitchOptions.GenerateFixedFASTAFile
                Me.mGenerateFixedFastaFile = State
            Case IValidateFastaFile.SwitchOptions.OutputToStatsFile
                Me.mOutputToStatsFile = State
            Case IValidateFastaFile.SwitchOptions.SplitOutMultipleRefsInProteinName
                Me.mFixedFastaOptions.SplitOutMultipleRefsInProteinName = State
            Case IValidateFastaFile.SwitchOptions.WarnBlankLinesBetweenProteins
                Me.mWarnBlankLinesBetweenProteins = State
            Case IValidateFastaFile.SwitchOptions.WarnLineStartsWithSpace
                Me.mWarnLineStartsWithSpace = State
            Case IValidateFastaFile.SwitchOptions.NormalizeFileLineEndCharacters
                Me.mNormalizeFileLineEndCharacters = State
            Case IValidateFastaFile.SwitchOptions.CheckForDuplicateProteinSequences
                Me.mCheckForDuplicateProteinSequences = State
            Case IValidateFastaFile.SwitchOptions.FixedFastaRenameDuplicateNameProteins
                Me.mFixedFastaOptions.RenameProteinsWithDuplicateNames = State
            Case IValidateFastaFile.SwitchOptions.SaveProteinSequenceHashInfoFiles
                Me.mSaveProteinSequenceHashInfoFiles = State
            Case IValidateFastaFile.SwitchOptions.FixedFastaConsolidateDuplicateProteinSeqs
                Me.mFixedFastaOptions.ConsolidateProteinsWithDuplicateSeqs = State
            Case IValidateFastaFile.SwitchOptions.FixedFastaConsolidateDupsIgnoreILDiff
                Me.mFixedFastaOptions.ConsolidateDupsIgnoreILDiff = State
            Case IValidateFastaFile.SwitchOptions.FixedFastaTruncateLongProteinNames
                Me.mFixedFastaOptions.TruncateLongProteinNames = State
            Case IValidateFastaFile.SwitchOptions.FixedFastaSplitOutMultipleRefsForKnownAccession
                Me.mFixedFastaOptions.SplitOutMultipleRefsForKnownAccession = State
            Case IValidateFastaFile.SwitchOptions.FixedFastaWrapLongResidueLines
                Me.mFixedFastaOptions.WrapLongResidueLines = State
            Case IValidateFastaFile.SwitchOptions.FixedFastaRemoveInvalidResidues
                Me.mFixedFastaOptions.RemoveInvalidResidues = State
            Case IValidateFastaFile.SwitchOptions.SaveBasicProteinHashInfoFile
                Me.mSaveBasicProteinHashInfoFile = State
            Case IValidateFastaFile.SwitchOptions.AllowDashInResidues
                Me.mAllowDashInResidues = State
        End Select

    End Sub

    Public Function GetOptionSwitchValue( _
        ByVal SwitchName As IValidateFastaFile.SwitchOptions) As Boolean

        Select Case SwitchName
            Case IValidateFastaFile.SwitchOptions.AddMissingLinefeedatEOF
                Return Me.mAddMissingLinefeedAtEOF
            Case IValidateFastaFile.SwitchOptions.AllowAsteriskInResidues
                Return Me.mAllowAsteriskInResidues
            Case IValidateFastaFile.SwitchOptions.CheckForDuplicateProteinNames
                Return Me.mCheckForDuplicateProteinNames
            Case IValidateFastaFile.SwitchOptions.GenerateFixedFASTAFile
                Return Me.mGenerateFixedFastaFile
            Case IValidateFastaFile.SwitchOptions.OutputToStatsFile
                Return Me.mOutputToStatsFile
            Case IValidateFastaFile.SwitchOptions.SplitOutMultipleRefsInProteinName
                Return Me.mFixedFastaOptions.SplitOutMultipleRefsInProteinName
            Case IValidateFastaFile.SwitchOptions.WarnBlankLinesBetweenProteins
                Return Me.mWarnBlankLinesBetweenProteins
            Case IValidateFastaFile.SwitchOptions.WarnLineStartsWithSpace
                Return Me.mWarnLineStartsWithSpace
            Case IValidateFastaFile.SwitchOptions.NormalizeFileLineEndCharacters
                Return Me.mNormalizeFileLineEndCharacters
            Case IValidateFastaFile.SwitchOptions.CheckForDuplicateProteinSequences
                Return Me.mCheckForDuplicateProteinSequences
            Case IValidateFastaFile.SwitchOptions.FixedFastaRenameDuplicateNameProteins
                Return Me.mFixedFastaOptions.RenameProteinsWithDuplicateNames
            Case IValidateFastaFile.SwitchOptions.SaveProteinSequenceHashInfoFiles
                Return Me.mSaveProteinSequenceHashInfoFiles
            Case IValidateFastaFile.SwitchOptions.FixedFastaConsolidateDuplicateProteinSeqs
                Return Me.mFixedFastaOptions.ConsolidateProteinsWithDuplicateSeqs
            Case IValidateFastaFile.SwitchOptions.FixedFastaConsolidateDupsIgnoreILDiff
                Return Me.mFixedFastaOptions.ConsolidateDupsIgnoreILDiff
            Case IValidateFastaFile.SwitchOptions.FixedFastaTruncateLongProteinNames
                Return Me.mFixedFastaOptions.TruncateLongProteinNames
            Case IValidateFastaFile.SwitchOptions.FixedFastaSplitOutMultipleRefsForKnownAccession
                Return Me.mFixedFastaOptions.SplitOutMultipleRefsForKnownAccession
            Case IValidateFastaFile.SwitchOptions.FixedFastaWrapLongResidueLines
                Return Me.mFixedFastaOptions.WrapLongResidueLines
            Case IValidateFastaFile.SwitchOptions.FixedFastaRemoveInvalidResidues
                Return Me.mFixedFastaOptions.RemoveInvalidResidues
            Case IValidateFastaFile.SwitchOptions.SaveBasicProteinHashInfoFile
                Return Me.mSaveBasicProteinHashInfoFile
            Case IValidateFastaFile.SwitchOptions.AllowDashInResidues
                Return Me.mAllowDashInResidues
        End Select

    End Function

    Public ReadOnly Property ErrorWarningCounts( _
    ByVal messageType As IValidateFastaFile.eMsgTypeConstants, _
    ByVal CountType As IValidateFastaFile.ErrorWarningCountTypes) As Integer

        Get
            Dim tmpValue As Integer
            Select Case CountType
                Case IValidateFastaFile.ErrorWarningCountTypes.Total
                    Select Case messageType
                        Case IValidateFastaFile.eMsgTypeConstants.ErrorMsg
                            tmpValue = Me.mFileErrorCount + Me.ComputeTotalUnspecifiedCount(Me.mFileErrorStats)
                        Case IValidateFastaFile.eMsgTypeConstants.WarningMsg
                            tmpValue = Me.mFileWarningCount + Me.ComputeTotalUnspecifiedCount(Me.mFileWarningStats)
                        Case IValidateFastaFile.eMsgTypeConstants.StatusMsg
                            tmpValue = 0
                    End Select
                Case IValidateFastaFile.ErrorWarningCountTypes.Unspecified
                    Select Case messageType
                        Case IValidateFastaFile.eMsgTypeConstants.ErrorMsg
                            tmpValue = Me.ComputeTotalUnspecifiedCount(Me.mFileErrorStats)
                        Case IValidateFastaFile.eMsgTypeConstants.WarningMsg
                            tmpValue = Me.ComputeTotalSpecifiedCount(Me.mFileWarningStats)
                        Case IValidateFastaFile.eMsgTypeConstants.StatusMsg
                            tmpValue = 0
                    End Select
                Case IValidateFastaFile.ErrorWarningCountTypes.Specified
                    Select Case messageType
                        Case IValidateFastaFile.eMsgTypeConstants.ErrorMsg
                            tmpValue = Me.mFileErrorCount
                        Case IValidateFastaFile.eMsgTypeConstants.WarningMsg
                            tmpValue = Me.mFileWarningCount
                        Case IValidateFastaFile.eMsgTypeConstants.StatusMsg
                            tmpValue = 0
                    End Select
            End Select

            Return tmpValue
        End Get
    End Property

    'Public ReadOnly Property FileErrorCountSpecified() As Integer
    'Public ReadOnly Property FileErrorCountUnspecified() As Integer
    'Public ReadOnly Property FileErrorCountTotal() As Integer
    'Public ReadOnly Property FileWarningCountTotal() As Integer

    Public ReadOnly Property FixedFASTAFileStats( _
        ByVal ValueType As IValidateFastaFile.FixedFASTAFileValues) As Integer _
        Implements IValidateFastaFile.FixedFASTAFileStats

        Get
            Dim tmpValue As Integer
            Select Case ValueType
                Case IValidateFastaFile.FixedFASTAFileValues.DuplicateProteinNamesSkippedCount
                    tmpValue = Me.mFixedFastaStats.DuplicateNameProteinsSkipped
                Case IValidateFastaFile.FixedFASTAFileValues.ProteinNamesInvalidCharsReplaced
                    tmpValue = Me.mFixedFastaStats.ProteinNamesInvalidCharsReplaced
                Case IValidateFastaFile.FixedFASTAFileValues.ProteinNamesMultipleRefsRemoved
                    tmpValue = Me.mFixedFastaStats.ProteinNamesMultipleRefsRemoved
                Case IValidateFastaFile.FixedFASTAFileValues.TruncatedProteinNameCount
                    tmpValue = Me.mFixedFastaStats.TruncatedProteinNameCount
                Case IValidateFastaFile.FixedFASTAFileValues.UpdatedResidueLines
                    tmpValue = Me.mFixedFastaStats.UpdatedResidueLines
                Case IValidateFastaFile.FixedFASTAFileValues.DuplicateProteinNamesRenamedCount
                    tmpValue = Me.mFixedFastaStats.DuplicateNameProteinsRenamed
                Case IValidateFastaFile.FixedFASTAFileValues.DuplicateProteinSeqsSkippedCount
                    tmpValue = Me.mFixedFastaStats.DuplicateSequenceProteinsSkipped
            End Select
            Return tmpValue

        End Get
    End Property

    'Public ReadOnly Property FixedFastaDuplicateProteinNamesSkippedCount() As Integer
    'Protected ReadOnly Property FixedFastaProteinNamesInvalidCharsReplaced() As Integer
    'Public ReadOnly Property FixedFastaProteinNamesMultipleRefsRemoved() As Integer
    'Public ReadOnly Property FixedFastaTruncatedProteinNameCount() As Integer
    'Public ReadOnly Property FixedFastaUpdatedResidueLines() As Integer

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

    Public ReadOnly Property ErrorMessageTextByIndex( _
        ByVal index As Integer, _
        ByVal valueSeparator As String) As String _
            Implements IValidateFastaFile.ErrorMessageTextByIndex
        Get
            Return Me.GetFileErrorTextByIndex(index, valueSeparator)
        End Get
    End Property

    Public ReadOnly Property WarningMessageTextByIndex( _
        ByVal index As Integer, _
        ByVal valueSeparator As String) As String _
            Implements IValidateFastaFile.WarningMessageTextByIndex
        Get
            Return Me.GetFileWarningTextByIndex(index, valueSeparator)
        End Get
    End Property

    Public ReadOnly Property ErrorsByIndex(ByVal errorIndex As Integer) As IValidateFastaFile.udtMsgInfoType _
        Implements IValidateFastaFile.FileErrorByIndex
        Get
            Return (Me.GetFileErrorByIndex(errorIndex))
        End Get
    End Property

    Public ReadOnly Property WarningsByIndex(ByVal warningIndex As Integer) As IValidateFastaFile.udtMsgInfoType _
        Implements IValidateFastaFile.FileWarningByIndex
        Get
            Return Me.GetFileWarningByIndex(warningIndex)
        End Get
    End Property

    Public Property MaximumFileErrorsToTrack() As Integer _
        Implements IValidateFastaFile.MaximumFileErrorsToTrack
        Get
            Return mMaximumFileErrorsToTrack
        End Get
        Set(ByVal Value As Integer)
            If Value < 1 Then Value = 1
            mMaximumFileErrorsToTrack = Value
        End Set
    End Property

    Public Property MaximumProteinNameLength() As Integer _
        Implements IValidateFastaFile.MaximumProteinNameLength
        Get
            Return mMaximumProteinNameLength
        End Get
        Set(ByVal Value As Integer)
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
        Set(ByVal Value As Integer)
            If Value < 1 Then Value = DEFAULT_MINIMUM_PROTEIN_NAME_LENGTH
            mMinimumProteinNameLength = Value
        End Set
    End Property

    Public Property MaximumResiduesPerLine() As Integer _
        Implements IValidateFastaFile.MaximumResiduesPerLine
        Get
            Return mMaximumResiduesPerLine
        End Get
        Set(ByVal Value As Integer)
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
        Set(ByVal Value As Char)
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
        Set(ByVal Value As String)
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
        Set(ByVal Value As String)
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
        Set(ByVal Value As String)
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
        Set(ByVal Value As String)
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
            Return Me.GetFileWarnings
        End Get
    End Property

    Public ReadOnly Property FileErrorList() As IValidateFastaFile.udtMsgInfoType() _
        Implements IValidateFastaFile.FileErrorList
        Get
            Return Me.GetFileErrors()
        End Get
    End Property


    Public Shadows Property ShowMessages() As Boolean Implements IValidateFastaFile.ShowMessages
        Get
            Return MyBase.ShowMessages
        End Get
        Set(ByVal Value As Boolean)
            MyBase.ShowMessages = Value
        End Set
    End Property

#End Region

    Protected Event ProgressUpdated( _
        ByVal taskDescription As String, _
        ByVal percentComplete As Single) Implements IValidateFastaFile.ProgressChanged

    Protected Event ProgressCompleted() Implements IValidateFastaFile.ProgressCompleted

    Protected Event WroteLineEndNormalizedFASTA(ByVal newFilePath As String) Implements IValidateFastaFile.WroteLineEndNormalizedFASTA

    Protected Sub OnProgressUpdate( _
        ByVal taskDescription As String, _
        ByVal percentComplete As Single) Handles MyBase.ProgressChanged

        RaiseEvent ProgressUpdated(taskDescription, percentComplete)

    End Sub

    Protected Sub OnProgressComplete() Handles MyBase.ProgressComplete
        RaiseEvent ProgressCompleted()
    End Sub

    Protected Sub OnWroteLineEndNormalizedFASTA(ByVal newFilePath As String)
        RaiseEvent WroteLineEndNormalizedFASTA(newFilePath)
    End Sub

    Protected Function AnalyzeFastaFile(ByVal strFastaFilePath As String) As Boolean
        ' This function assumes strFastaFilePath exists
        ' Returns True if the file was successfully analyzed (even if errors were found)

        Dim fsInFile As System.IO.Stream
        Dim srFastaInFile As System.IO.StreamReader
        Dim swFixedFastaOut As System.IO.StreamWriter
        Dim swProteinSequenceHashBasic As System.IO.StreamWriter

        Dim lngBytesRead As Long
        Dim intTerminatorSize As Integer
        Dim sngPercentComplete As Single

        Dim strFastaFilePathOut As String = String.Empty
        Dim strBasicProteinHashInfoFilePath As String = String.Empty

        Dim strLineIn As String
        Dim strResiduesClean As String
        Dim strNonLetterResidueSpec As String

        Dim sbCurrentResidues As System.Text.StringBuilder

        ' Note: This value is updated only if the line length is < mMaximumResiduesPerLine
        Dim intCurrentValidResidueLineLengthMax As Integer

        Dim intIndex As Integer
        Dim intNewResidueCount As Integer

        Dim strProteinName As String

        Dim blnSuccess As Boolean
        Dim blnProteinHeaderFound As Boolean
        Dim blnProcessingResidueBlock As Boolean
        Dim blnProcessingDuplicateOrInvalidProtein As Boolean

        Dim blnConsolidatingProteinsWithDuplicateSeqs As Boolean
        Dim blnConsolidateDupsIgnoreILDiff As Boolean

        Dim blnBlankLineProcessed As Boolean

        ' Note that text is stored in the string dictionary lowercase, and searches are thus case-insensitive
        Dim htProteinNames As System.Collections.Specialized.StringDictionary

        ' This hashtable provides a quick lookup for existing protein hashes
        Dim htProteinSequenceHashes As Hashtable

        ' This array tracks protein hash details
        Dim intProteinSequenceHashCount As Integer
        Dim udtProteinSeqHashInfo() As udtProteinHashInfoType

        Dim udtHeaderLineRuleDetails() As udtRuleDefinitionExtendedType
        Dim udtProteinNameRuleDetails() As udtRuleDefinitionExtendedType
        Dim udtProteinDescriptionRuleDetails() As udtRuleDefinitionExtendedType
        Dim udtProteinSequenceRuleDetails() As udtRuleDefinitionExtendedType

        Dim reProteinNameTruncation As udtProteinNameTruncationRegex

        Dim reNonLetterResidues As System.Text.RegularExpressions.Regex

        Try
            ' Reset the data structures and variables
            ResetStructures()

            If mNormalizeFileLineEndCharacters Then
                mFastaFilePath = Me.NormalizeFileLineEndings( _
                    strFastaFilePath, _
                    "CRLF_" & System.IO.Path.GetFileName(strFastaFilePath), _
                    IValidateFastaFile.eLineEndingCharacters.CRLF)

                If mFastaFilePath <> strFastaFilePath Then
                    strFastaFilePath = String.Copy(mFastaFilePath)
                    Me.OnWroteLineEndNormalizedFASTA(strFastaFilePath)
                End If
            Else
                mFastaFilePath = String.Copy(strFastaFilePath)
            End If

            Me.OnProgressUpdate("Parsing " & System.IO.Path.GetFileName(mFastaFilePath), 0)

            blnSuccess = False
            blnProteinHeaderFound = False
            blnProcessingResidueBlock = False
            blnBlankLineProcessed = False

            strProteinName = String.Empty
            sbCurrentResidues = New System.Text.StringBuilder

            ' Initialize the RegEx objects

            With reProteinNameTruncation
                ' Note that each of these RegEx tests contain two groups with captured text:

                ' The following will extract IPI:IPI00048500.11 from IPI:IPI00048500.11|ref|23848934
                .reMatchIPI = _
                    New System.Text.RegularExpressions.Regex("^(IPI:IPI[\w.]{2,})\|(.+)", _
                        Text.RegularExpressions.RegexOptions.Singleline Or Text.RegularExpressions.RegexOptions.Compiled)

                ' The following will extract gi|169602219 from gi|169602219|ref|XP_001794531.1|
                .reMatchGI = _
                    New System.Text.RegularExpressions.Regex("^(gi\|\d+)\|(.+)", _
                        Text.RegularExpressions.RegexOptions.Singleline Or Text.RegularExpressions.RegexOptions.Compiled)

                ' The following will extract jgi|Batde5|906240 from jgi|Batde5|90624|GP3.061830
                .reMatchJGI = _
                    New System.Text.RegularExpressions.Regex("^(jgi\|[^|]+\|[^|]+)\|(.+)", _
                        Text.RegularExpressions.RegexOptions.Singleline Or Text.RegularExpressions.RegexOptions.Compiled)

                ' The following will extract bob|234384 from  bob|234384|ref|483293
                '                         or bob|845832 from  bob|845832;ref|384923
                .reMatchGeneric = _
                    New System.Text.RegularExpressions.Regex("^(\w{2,}[" & _
                    CharArrayToString(mProteinNameFirstRefSepChars) & "][\w\d._]{2,})[" & _
                    CharArrayToString(mProteinNameSubsequentRefSepChars) & "](.+)", _
                    Text.RegularExpressions.RegexOptions.Singleline Or Text.RegularExpressions.RegexOptions.Compiled)
            End With

            With reProteinNameTruncation
                ' The following matches jgi|Batde5|23435 ; it requires that there be a number after the second bar
                .reMatchJGIBaseAndID = _
                    New System.Text.RegularExpressions.Regex("^jgi\|[^|]+\|\d+", _
                            Text.RegularExpressions.RegexOptions.Singleline Or Text.RegularExpressions.RegexOptions.Compiled)

                ' Note that this RegEx contains a group with captured text:
                .reMatchDoubleBarOrColonAndBar = _
                    New System.Text.RegularExpressions.Regex("[" & _
                        CharArrayToString(mProteinNameFirstRefSepChars) & "][^" & _
                        CharArrayToString(mProteinNameSubsequentRefSepChars) & "]*([" & _
                        CharArrayToString(mProteinNameSubsequentRefSepChars) & "])", _
                        Text.RegularExpressions.RegexOptions.Singleline Or Text.RegularExpressions.RegexOptions.Compiled)
            End With

            ' Non-letter characters in residues
            strNonLetterResidueSpec = "A-Z"
            If mAllowAsteriskInResidues Then strNonLetterResidueSpec &= "*"
            If mAllowDashInResidues Then strNonLetterResidueSpec &= "-"

            reNonLetterResidues = _
                New System.Text.RegularExpressions.Regex("[^" & strNonLetterResidueSpec & "]", _
                Text.RegularExpressions.RegexOptions.Singleline Or Text.RegularExpressions.RegexOptions.Compiled)

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
            intTerminatorSize = DetermineLineTerminatorSize(strFastaFilePath)

            ' Open the input file
            fsInFile = New System.IO.FileStream( _
               strFastaFilePath, _
               System.IO.FileMode.Open, _
               System.IO.FileAccess.Read, _
               IO.FileShare.Read)

            srFastaInFile = New System.IO.StreamReader(fsInFile)

            ' Optionally, open the output fasta file
            If mGenerateFixedFastaFile Then
                Try
                    strFastaFilePathOut = _
                        System.IO.Path.Combine(System.IO.Path.GetDirectoryName(strFastaFilePath), _
                        System.IO.Path.GetFileNameWithoutExtension(strFastaFilePath) & "_new.fasta")
                    swFixedFastaOut = New System.IO.StreamWriter(strFastaFilePathOut, False)
                Catch ex As Exception
                    ' Error opening output file
                    RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError, _
                        "Error creating output file " & strFastaFilePathOut & ": " & ex.message, String.Empty)
                End Try
            End If

            ' Optionally, open the Sequence Hash file
            If mSaveBasicProteinHashInfoFile Then
                Try
                    strBasicProteinHashInfoFilePath = _
                        System.IO.Path.Combine(System.IO.Path.GetDirectoryName(strFastaFilePath), _
                        System.IO.Path.GetFileNameWithoutExtension(strFastaFilePath) & "_ProteinHashes.txt")
                    swProteinSequenceHashBasic = New System.IO.StreamWriter(strBasicProteinHashInfoFilePath, False)

                    swProteinSequenceHashBasic.WriteLine("Protein_ID" & ControlChars.Tab & _
                                                         "Protein_Name" & ControlChars.Tab & _
                                                         "Sequence_Length" & ControlChars.Tab & _
                                                         "Sequence_Hash")

                Catch ex As Exception
                    ' Error opening output file
                    RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError, _
                        "Error creating output file " & strBasicProteinHashInfoFilePath & ": " & ex.message, String.Empty)
                End Try

            End If

            If mGenerateFixedFastaFile And mFixedFastaOptions.RenameProteinsWithDuplicateNames Then
                ' Make sure mCheckForDuplicateProteinNames is enabled
                mCheckForDuplicateProteinNames = True
            End If

            ' Optionally, initialize htProteinNames
            If mCheckForDuplicateProteinNames Then
                htProteinNames = New System.Collections.Specialized.StringDictionary
            End If

            ' Optionally, initialize the protein sequence hash objects
            If mSaveProteinSequenceHashInfoFiles Then
                mCheckForDuplicateProteinSequences = True
            End If

            If mGenerateFixedFastaFile And mFixedFastaOptions.ConsolidateProteinsWithDuplicateSeqs Then
                mCheckForDuplicateProteinSequences = True
                mSaveProteinSequenceHashInfoFiles = True
                blnConsolidatingProteinsWithDuplicateSeqs = True
                blnConsolidateDupsIgnoreILDiff = mFixedFastaOptions.ConsolidateDupsIgnoreILDiff
            Else
                blnConsolidatingProteinsWithDuplicateSeqs = False
                blnConsolidateDupsIgnoreILDiff = False
            End If

            If mCheckForDuplicateProteinSequences Then
                htProteinSequenceHashes = New Hashtable
                intProteinSequenceHashCount = 0
                ReDim udtProteinSeqHashInfo(99)
            End If

            ' Parse each line in the file
            lngBytesRead = 0

            Do While srFastaInFile.Peek() >= 0
                strLineIn = srFastaInFile.ReadLine()
                lngBytesRead += strLineIn.Length + intTerminatorSize

                If mLineCount Mod 50 = 0 Then
                    If MyBase.AbortProcessing Then Exit Do

                    sngPercentComplete = CType(lngBytesRead / CType(srFastaInFile.BaseStream.Length, Single) * 100.0, Single)
                    If blnConsolidatingProteinsWithDuplicateSeqs Then
                        ' Bump the % complete down so that 100% complete in this routine will equate to 75% complete
                        ' The remaining 25% will occur in ConsolidateDuplicateProteinSeqsInFasta
                        sngPercentComplete = sngPercentComplete * 3 / 4
                    End If

                    MyBase.UpdateProgress("Validating FASTA File (" & Math.Round(sngPercentComplete, 0) & "% Done)", sngPercentComplete)
                End If

                mLineCount += 1

                If Not strLineIn Is Nothing Then
                    If strLineIn.Trim.Length = 0 Then
                        ' We typically only want blank lines at the end of the fasta file or between two protein entries
                        blnBlankLineProcessed = True
                    Else
                        If strLineIn.Chars(0) = " "c Then
                            If mWarnLineStartsWithSpace Then
                                RecordFastaFileError(mLineCount, 0, String.Empty, _
                                    eMessageCodeConstants.LineStartsWithSpace, String.Empty, ExtractContext(strLineIn, 0))
                            End If
                        End If

                        ' Note: Only trim the start of the line; do not trim the end of the line since Sequest incorrectly notates the peptide terminal state if a residue has a space after it
                        strLineIn = strLineIn.TrimStart

                        If strLineIn.Chars(0) = mProteinLineStartChar Then
                            ' Protein entry

                            If sbCurrentResidues.Length > 0 Then
                                ProcessResiduesForPreviousProtein( _
                                                        strProteinName, sbCurrentResidues, _
                                                        htProteinSequenceHashes, intProteinSequenceHashCount, udtProteinSeqHashInfo, _
                                                        blnConsolidateDupsIgnoreILDiff, _
                                                        swFixedFastaOut, intCurrentValidResidueLineLengthMax, _
                                                        swProteinSequenceHashBasic)

                                intCurrentValidResidueLineLengthMax = 0
                            End If

                            ' Now process this protein entry
                            mProteinCount += 1
                            blnProteinHeaderFound = True
                            blnProcessingResidueBlock = False
                            blnProcessingDuplicateOrInvalidProtein = False

                            strProteinName = String.Empty

                            AnalyzeFastaProcessProteinHeader( _
                                swFixedFastaOut, _
                                strLineIn, _
                                strProteinName, _
                                blnProcessingDuplicateOrInvalidProtein, _
                                htProteinNames, _
                                udtHeaderLineRuleDetails, _
                                udtProteinNameRuleDetails, _
                                udtProteinDescriptionRuleDetails, _
                                udtProteinSequenceRuleDetails, _
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


                            intNewResidueCount = strLineIn.Length
                            mResidueCount += intNewResidueCount

                            ' Check the line length; raise a warning if longer than suggested
                            If intNewResidueCount > mMaximumResiduesPerLine Then
                                RecordFastaFileWarning(mLineCount, 0, strProteinName, eMessageCodeConstants.ResiduesLineTooLong, intNewResidueCount.ToString, String.Empty)
                            End If

                            ' Test the protein sequence rules
                            EvaluateRules(udtProteinSequenceRuleDetails, strProteinName, strLineIn, 0, strLineIn, 5)

                            If mGenerateFixedFastaFile OrElse mCheckForDuplicateProteinSequences OrElse mSaveBasicProteinHashInfoFile Then
                                If mFixedFastaOptions.RemoveInvalidResidues Then
                                    ' Auto-fix residues to remove any non-letter characters (spaces, asterisks, etc.)
                                    strResiduesClean = reNonLetterResidues.Replace(strLineIn, String.Empty)
                                Else
                                    strResiduesClean = String.Copy(strLineIn)
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
                    End If
                End If
            Loop

            If sbCurrentResidues.Length > 0 Then
                ProcessResiduesForPreviousProtein( _
                                        strProteinName, sbCurrentResidues, _
                                        htProteinSequenceHashes, intProteinSequenceHashCount, udtProteinSeqHashInfo, _
                                        blnConsolidateDupsIgnoreILDiff, _
                                        swFixedFastaOut, intCurrentValidResidueLineLengthMax, _
                                        swProteinSequenceHashBasic)
            End If

            If mCheckForDuplicateProteinSequences Then
                ' Step through udtProteinSeqHashInfo and look for duplicate sequences
                For intIndex = 0 To intProteinSequenceHashCount - 1
                    If udtProteinSeqHashInfo(intIndex).ProteinCount > 1 Then
                        With udtProteinSeqHashInfo(intIndex)
                            RecordFastaFileWarning(mLineCount, 0, .ProteinNameFirst, eMessageCodeConstants.DuplicateProteinSequence, _
                                       .ProteinNameFirst & ", " & FlattenArray(.AdditionalProteins, .ProteinCount - 1, ","c), .SequenceStart)
                        End With
                    End If
                Next intIndex
            End If

            ' Close the input and output files
            If Not srFastaInFile Is Nothing Then
                srFastaInFile.Close()
            End If

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
                System.Threading.Thread.Sleep(100)

                If Not VerifyLinefeedAtEOF(strFastaFilePath, mAddMissingLinefeedAtEOF) Then
                    RecordFastaFileError(mLineCount, 0, String.Empty, eMessageCodeConstants.FileDoesNotEndWithLinefeed)
                End If
            End If

            If mSaveProteinSequenceHashInfoFiles Then
                sngPercentComplete = 98
                If blnConsolidatingProteinsWithDuplicateSeqs Then
                    sngPercentComplete = sngPercentComplete * 3 / 4
                End If
                MyBase.UpdateProgress("Validating FASTA File (" & Math.Round(sngPercentComplete, 0) & "% Done)", sngPercentComplete)

                blnSuccess = AnalysisFastaSaveHashInfo( _
                                    strFastaFilePath, _
                                    intProteinSequenceHashCount, _
                                    udtProteinSeqHashInfo, _
                                    blnConsolidatingProteinsWithDuplicateSeqs, _
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
                MsgBox("Error in AnalyzeFastaFile:" & ControlChars.NewLine & ex.Message, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
            Else
                Throw New System.Exception("Error in AnalyzeFastaFile", ex)
            End If
            blnSuccess = False
        Finally
            ' These close statements will typically be redundant,
            ' However, if an exception occurs, then they will be needed to close the files

            If Not srFastaInFile Is Nothing Then
                srFastaInFile.Close()
            End If

            If Not swFixedFastaOut Is Nothing Then
                swFixedFastaOut.Close()
            End If

            If Not swProteinSequenceHashBasic Is Nothing Then
                swProteinSequenceHashBasic.Close()
            End If

        End Try

        Return blnSuccess

    End Function

    Protected Sub AnalyzeFastaProcessProteinHeader( _
                                ByRef swFixedFastaOut As System.IO.StreamWriter, _
                                ByRef strLineIn As String, _
                                ByRef strProteinName As String, _
                                ByRef blnProcessingDuplicateOrInvalidProtein As Boolean, _
                                ByRef htProteinNames As System.Collections.Specialized.StringDictionary, _
                                ByRef udtHeaderLineRuleDetails() As udtRuleDefinitionExtendedType, _
                                ByRef udtProteinNameRuleDetails() As udtRuleDefinitionExtendedType, _
                                ByRef udtProteinDescriptionRuleDetails() As udtRuleDefinitionExtendedType, _
                                ByRef udtProteinSequenceRuleDetails() As udtRuleDefinitionExtendedType, _
                                ByRef reProteinNameTruncation As udtProteinNameTruncationRegex)

        Dim intDescriptionStartIndex As Integer
        Dim intCharIndex As Integer
        Dim intMinCharIndex As Integer

        Dim strNewProteinName As String
        Dim strExtraProteinNameText As String
        Dim strProteinDescription As String = String.Empty

        Dim reMatch As System.Text.RegularExpressions.Match

        Dim blnDuplicateName As Boolean
        Dim chLetterToAppend As Char
        Dim intNumberToAppend As Integer

        Dim blnMultipleRefsSplitOutFromKnownAccession As Boolean = False

        Dim blnProteinNameTooLong As Boolean
        Dim chInvalidChar As Char

        Try
            SplitFastaProteinHeaderLine(strLineIn, strProteinName, strProteinDescription, intDescriptionStartIndex)

            If strProteinName.Length = 0 Then
                blnProcessingDuplicateOrInvalidProtein = True
            Else
                blnProcessingDuplicateOrInvalidProtein = False
            End If

            ' Test the header line rules
            EvaluateRules(udtHeaderLineRuleDetails, strProteinName, strLineIn, 0, strLineIn, DEFAULT_CONTEXT_LENGTH)

            If strProteinDescription.Length > 0 Then
                ' Test the protein description rules

                EvaluateRules(udtProteinDescriptionRuleDetails, strProteinName, strProteinDescription, _
                    intDescriptionStartIndex, strLineIn, DEFAULT_CONTEXT_LENGTH)
            End If

            If strProteinName.Length > 0 Then

                ' Check for protein names that are too long or too short
                If strProteinName.Length < mMinimumProteinNameLength Then
                    RecordFastaFileWarning(mLineCount, 1, strProteinName, _
                        eMessageCodeConstants.ProteinNameIsTooShort, strProteinName.Length.ToString, String.Empty)
                ElseIf strProteinName.Length > mMaximumProteinNameLength Then
                    RecordFastaFileError(mLineCount, 1, strProteinName, _
                        eMessageCodeConstants.ProteinNameIsTooLong, strProteinName.Length.ToString, String.Empty)
                End If

                ' Test the protein name rules
                EvaluateRules(udtProteinNameRuleDetails, strProteinName, strProteinName, 1, strLineIn, DEFAULT_CONTEXT_LENGTH)

                If mGenerateFixedFastaFile Then
                    ' Auto-fix potential errors in the protein name

                    ' Possibly truncate to mMaximumProteinNameLength characters
                    If strProteinName.Length > mMaximumProteinNameLength Then
                        blnProteinNameTooLong = True
                    Else
                        blnProteinNameTooLong = False
                    End If

                    If mFixedFastaOptions.SplitOutMultipleRefsForKnownAccession OrElse _
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
                End If

                ' Optionally, check for duplicate protein names
                If mCheckForDuplicateProteinNames Then
                    blnDuplicateName = htProteinNames.ContainsKey(strProteinName.ToLower)

                    If blnDuplicateName AndAlso _
                       mGenerateFixedFastaFile AndAlso _
                       mFixedFastaOptions.RenameProteinsWithDuplicateNames Then

                        chLetterToAppend = "b"c
                        intNumberToAppend = 0
                        Do
                            strNewProteinName = strProteinName & "-"c & chLetterToAppend
                            If intNumberToAppend > 0 Then
                                strNewProteinName &= intNumberToAppend.ToString
                            End If

                            blnDuplicateName = htProteinNames.ContainsKey(strNewProteinName.ToLower)

                            If blnDuplicateName Then
                                ' Increment chLetterToAppend to the next letter and then try again to rename the protein
                                If chLetterToAppend = "z"c Then
                                    ' We've reached "z"
                                    ' Change back to "a" but increment intNumberToAppend
                                    chLetterToAppend = "a"c
                                    intNumberToAppend += 1
                                Else
                                    ' chLetterToAppend = Chr(Asc(chLetterToAppend) + 1)
                                    chLetterToAppend = System.Convert.ToChar(System.Convert.ToInt32(chLetterToAppend) + 1)
                                End If
                            End If
                        Loop While blnDuplicateName

                        RecordFastaFileWarning(mLineCount, 1, strProteinName, eMessageCodeConstants.RenamedProtein, "--> " & strNewProteinName, String.Empty)

                        strProteinName = String.Copy(strNewProteinName)
                        mFixedFastaStats.DuplicateNameProteinsRenamed += 1

                    End If

                    If blnDuplicateName Then
                        RecordFastaFileError(mLineCount, 1, strProteinName, eMessageCodeConstants.DuplicateProteinName)
                        blnProcessingDuplicateOrInvalidProtein = True
                        mFixedFastaStats.DuplicateNameProteinsSkipped += 1
                    Else
                        htProteinNames.Add(strProteinName.ToLower, Nothing)
                        blnProcessingDuplicateOrInvalidProtein = False
                    End If
                End If

                If Not swFixedFastaOut Is Nothing AndAlso Not blnProcessingDuplicateOrInvalidProtein Then
                    swFixedFastaOut.WriteLine(ConstructFastaHeaderLine(strProteinName.Trim, strProteinDescription.Trim))
                End If
            End If


        Catch ex As Exception
            RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError, _
                "Error parsing protein header line '" & strLineIn & "': " & ex.Message, String.Empty)

        End Try


    End Sub

    Private Function AnalysisFastaSaveHashInfo( _
                    ByVal strFastaFilePath As String, _
                    ByVal intProteinSequenceHashCount As Integer, _
                    ByRef udtProteinSeqHashInfo() As udtProteinHashInfoType, _
                    ByVal blnConsolidateDuplicateProteinSeqsInFasta As Boolean, _
                    ByVal strFastaFilePathOut As String) As Boolean

        Dim swUniqueProteinSeqsOut As System.IO.StreamWriter
        Dim swDuplicateProteinMapping As System.IO.StreamWriter

        Dim strUniqueProteinSeqsFileOut As String = String.Empty
        Dim strDuplicateProteinMappingFileOut As String = String.Empty
        Dim strLineOut As String

        Dim intIndex As Integer
        Dim intDuplicateIndex As Integer

        Dim strDuplicateProtein As String

        Dim blnDuplicateProteinSeqsFound As Boolean
        Dim blnSuccess As Boolean

        blnDuplicateProteinSeqsFound = False
        blnSuccess = False

        Try
            strUniqueProteinSeqsFileOut = _
                System.IO.Path.Combine(System.IO.Path.GetDirectoryName(strFastaFilePath), _
                System.IO.Path.GetFileNameWithoutExtension(strFastaFilePath) & "_UniqueProteinSeqs.txt")

            ' Create swUniqueProteinSeqsOut
            swUniqueProteinSeqsOut = New System.IO.StreamWriter(strUniqueProteinSeqsFileOut, False)
        Catch ex As Exception
            ' Error opening output file
            RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError, _
                "Error creating output file " & strUniqueProteinSeqsFileOut & ": " & ex.message, String.Empty)

            Return False
        End Try

        Try
            ' Define the path to the protein mapping file, but don't create it yet; just delete it if it exists
            ' We'll only create it if two or more proteins have the same protein sequence
            strDuplicateProteinMappingFileOut = _
                                   System.IO.Path.Combine(System.IO.Path.GetDirectoryName(strFastaFilePath), _
                                   System.IO.Path.GetFileNameWithoutExtension(strFastaFilePath) & "_UniqueProteinSeqDuplicates.txt")                    ' Look for strDuplicateProteinMappingFileOut and erase it if it exists

            If System.IO.File.Exists(strDuplicateProteinMappingFileOut) Then
                System.IO.File.Delete(strDuplicateProteinMappingFileOut)
            End If
        Catch ex As Exception
            ' Error deleting output file
            RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError, _
                "Error deleting output file " & strDuplicateProteinMappingFileOut & ": " & ex.message, String.Empty)

            Return False
        End Try

        Try

            strLineOut = "Sequence_Index" & ControlChars.Tab & _
                         "Protein_Name_First" & ControlChars.Tab & _
                         "Sequence_Length" & ControlChars.Tab & _
                         "Sequence_Hash" & ControlChars.Tab & _
                         "Protein_Count" & ControlChars.Tab & _
                         "Duplicate_Proteins"

            swUniqueProteinSeqsOut.WriteLine(strLineOut)

            For intIndex = 0 To intProteinSequenceHashCount - 1
                With udtProteinSeqHashInfo(intIndex)
                    strLineOut = (intIndex + 1).ToString & ControlChars.Tab & _
                                 .ProteinNameFirst & ControlChars.Tab & _
                                 .SequenceLength & ControlChars.Tab & _
                                 .SequenceHash & ControlChars.Tab & _
                                 .ProteinCount & ControlChars.Tab & _
                                 FlattenArray(.AdditionalProteins, .ProteinCount - 1, ","c)

                    swUniqueProteinSeqsOut.WriteLine(strLineOut)

                    If .ProteinCount > 1 Then
                        blnDuplicateProteinSeqsFound = True

                        If swDuplicateProteinMapping Is Nothing Then
                            ' Need to create swDuplicateProteinMapping
                            swDuplicateProteinMapping = New System.IO.StreamWriter(strDuplicateProteinMappingFileOut, False)

                            strLineOut = "Sequence_Index" & ControlChars.Tab & _
                                         "Protein_Name_First" & ControlChars.Tab & _
                                         "Sequence_Length" & ControlChars.Tab & _
                                         "Duplicate_Protein"

                            swDuplicateProteinMapping.WriteLine(strLineOut)
                        End If

                        For intDuplicateIndex = 0 To .ProteinCount - 2
                            If Not .AdditionalProteins(intDuplicateIndex) Is Nothing Then
                                strDuplicateProtein = .AdditionalProteins(intDuplicateIndex).Trim
                                If strDuplicateProtein.Length > 0 Then
                                    strLineOut = (intIndex + 1).ToString & ControlChars.Tab & _
                                                 .ProteinNameFirst & ControlChars.Tab & _
                                                 .SequenceLength & ControlChars.Tab & _
                                                 strDuplicateProtein
                                    swDuplicateProteinMapping.WriteLine(strLineOut)
                                End If
                            End If
                        Next intDuplicateIndex
                    End If
                End With

            Next intIndex
            swUniqueProteinSeqsOut.Close()
            If Not swDuplicateProteinMapping Is Nothing Then swDuplicateProteinMapping.Close()

            blnSuccess = True

        Catch ex As Exception
            ' Error writing results
            RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError, _
                "Error writing results to " & strUniqueProteinSeqsFileOut & " or " & strDuplicateProteinMappingFileOut & ": " & ex.message, String.Empty)

            blnSuccess = False
        End Try

        If blnSuccess And intProteinSequenceHashCount > 0 And blnDuplicateProteinSeqsFound Then
            If blnConsolidateDuplicateProteinSeqsInFasta Then
                blnSuccess = ConsolidateDuplicateProteinSeqsInFasta(strFastaFilePathOut, intProteinSequenceHashCount, udtProteinSeqHashInfo)
            End If
        End If

        Return blnSuccess

    End Function

    Private Function CharArrayToString(ByVal chCharArray() As Char) As String
        Return CStr(chCharArray)
    End Function

    Protected Shadows Sub AbortProcessingNow() Implements IValidateFastaFile.AbortProcessingNow
        MyBase.AbortProcessingNow()
    End Sub

    Protected Sub ClearAllRules() Implements IValidateFastaFile.ClearAllRules
        Me.ClearRules(IValidateFastaFile.RuleTypes.HeaderLine)
        Me.ClearRules(IValidateFastaFile.RuleTypes.ProteinDescription)
        Me.ClearRules(IValidateFastaFile.RuleTypes.ProteinName)
        Me.ClearRules(IValidateFastaFile.RuleTypes.ProteinSequence)

        mMasterCustomRuleID = CUSTOM_RULE_ID_START
    End Sub

    Protected Sub ClearRules(ByVal ruleType As IValidateFastaFile.RuleTypes) Implements IValidateFastaFile.ClearRules
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

    Private Function ComputeTotalSpecifiedCount(ByVal udtErrorStats As udtItemSummaryIndexedType) As Integer
        Dim intTotal As Integer
        Dim intIndex As Integer

        intTotal = 0
        For intIndex = 0 To udtErrorStats.ErrorStatsCount - 1
            intTotal += udtErrorStats.ErrorStats(intIndex).CountSpecified
        Next intIndex

        Return intTotal

    End Function

    Private Function ComputeTotalUnspecifiedCount(ByVal udtErrorStats As udtItemSummaryIndexedType) As Integer
        Dim intTotal As Integer
        Dim intIndex As Integer

        intTotal = 0
        For intIndex = 0 To udtErrorStats.ErrorStatsCount - 1
            intTotal += udtErrorStats.ErrorStats(intIndex).CountUnspecified
        Next intIndex

        Return intTotal

    End Function

    Protected Function ConsolidateDuplicateProteinSeqsInFasta( _
                    ByVal strFixedFastaFilePath As String, _
                    ByVal intProteinSequenceHashCount As Integer, _
                    ByRef udtProteinSeqHashInfo() As udtProteinHashInfoType) As Boolean

        Dim fsInFile As System.IO.Stream
        Dim srFastaInFile As System.IO.StreamReader
        Dim swConsolidatedFastaOut As System.IO.StreamWriter

        Dim lngBytesRead As Long
        Dim intTerminatorSize As Integer
        Dim sngPercentComplete As Single
        Dim intLineCount As Integer

        Dim strFixedFastaFilePathTemp As String = String.Empty
        Dim strLineIn As String

        Dim strProteinName As String = String.Empty
        Dim strProteinDescription As String = String.Empty
        Dim strMasterProteinInfo As String

        ' This hashtable contains the protein names that we will keep, the hash values are the index values pointing into udtProteinSeqHashInfo
        Dim htProteinNameFirstList As Hashtable

        ' This hashtable contains the names of dupliate proteins; the hash values are the protein names of the master protein that has the same sequence
        Dim htDuplicateProteinList As Hashtable
        Dim objValue As Object

        Dim intDescriptionStartIndex As Integer
        Dim intIndex As Integer
        Dim intDupIndex As Integer
        Dim intDuplicateNameSkipCount As Integer

        Dim blnKeepProtein As Boolean
        Dim blnSuccess As Boolean
        blnSuccess = False

        Dim blnSkipDupProtein As Boolean

        Dim intAdditionalProteinCount As Integer
        Dim strAdditionalProteinNames() As String

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

            If System.IO.File.Exists(strFixedFastaFilePathTemp) Then
                System.IO.File.Delete(strFixedFastaFilePathTemp)
            End If

            System.IO.File.Move(strFixedFastaFilePath, strFixedFastaFilePathTemp)
        Catch ex As Exception
            RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError, _
                "Error renaming " & strFixedFastaFilePath & " to " & strFixedFastaFilePathTemp & ": " & ex.message, String.Empty)

            Return False
        End Try

        Try
            ' Open the file and read, at most, the first 100,000 characters to see if it contains CrLf or just Lf
            intTerminatorSize = DetermineLineTerminatorSize(strFixedFastaFilePathTemp)

            ' Open the Fixed fasta file
            fsInFile = New System.IO.FileStream( _
               strFixedFastaFilePathTemp, _
               System.IO.FileMode.Open, _
               System.IO.FileAccess.Read, _
               IO.FileShare.Read)

            srFastaInFile = New System.IO.StreamReader(fsInFile)

        Catch ex As Exception
            RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError, _
                "Error opening " & strFixedFastaFilePathTemp & ": " & ex.message, String.Empty)

            Return False
        End Try

        Try
            ' Create the new fasta file
            swConsolidatedFastaOut = New System.IO.StreamWriter(strFixedFastaFilePath, False)
        Catch ex As Exception
            RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError, _
                "Error creating consolidated fasta output file " & strFixedFastaFilePath & ": " & ex.message, String.Empty)
        End Try

        Try
            ' Populate htProteinNameFirstList with the protein names in udtProteinSeqHashInfo().ProteinNameFirst
            htProteinNameFirstList = New Hashtable

            ' Populate htDuplicateProteinList with the protein names in udtProteinSeqHashInfo().AdditionalProteins 
            htDuplicateProteinList = New Hashtable

            For intIndex = 0 To intProteinSequenceHashCount - 1
                With udtProteinSeqHashInfo(intIndex)
                    strProteinName = .ProteinNameFirst.ToLower

                    objValue = htProteinNameFirstList.Item(strProteinName)
                    If objValue Is Nothing Then
                        htProteinNameFirstList.Add(strProteinName, intIndex)
                    Else
                        ' strProteinName is already present in htProteinNameFirstList
                        ' The fixed fasta file will only actually contain the first occurrence of strProteinName, so we can effectively ignore this entry

                        intDuplicateNameSkipCount += 1
                    End If

                    If .ProteinCount > 1 Then
                        For intDupIndex = 0 To .ProteinCount - 2
                            strProteinName = .AdditionalProteins(intDupIndex).ToLower
                            objValue = htDuplicateProteinList.Item(strProteinName)
                            If objValue Is Nothing Then
                                htDuplicateProteinList.Add(.AdditionalProteins(intDupIndex).ToLower, .ProteinNameFirst)
                            End If
                        Next
                    End If
                End With
            Next intIndex

            ' Reserve space for 10 entries in strAdditionalProteinNames
            intAdditionalProteinCount = 0
            ReDim strAdditionalProteinNames(9)

            ' Parse each line in the file
            intLineCount = 0
            lngBytesRead = 0
            mFixedFastaStats.DuplicateSequenceProteinsSkipped = 0

            Do While srFastaInFile.Peek() >= 0
                strLineIn = srFastaInFile.ReadLine()
                lngBytesRead += strLineIn.Length + intTerminatorSize

                If intLineCount Mod 50 = 0 Then
                    If MyBase.AbortProcessing Then Exit Do

                    sngPercentComplete = 75 + CType(lngBytesRead / CType(srFastaInFile.BaseStream.Length, Single) * 100.0, Single) / 4
                    MyBase.UpdateProgress("Consolidating duplicate proteins to create a new FASTA File (" & Math.Round(sngPercentComplete, 0) & "% Done)", sngPercentComplete)
                End If

                intLineCount += 1

                If Not strLineIn Is Nothing Then
                    If strLineIn.Trim.Length > 0 Then
                        ' Note: Trim the start of the line (however, since this is a fixed fasta file it should not start with a space)
                        strLineIn = strLineIn.TrimStart

                        If strLineIn.Chars(0) = mProteinLineStartChar Then
                            ' Protein entry
                            ' Extract the protein name and description

                            SplitFastaProteinHeaderLine(strLineIn, strProteinName, strProteinDescription, intDescriptionStartIndex)

                            objValue = htProteinNameFirstList.Item(strProteinName.ToLower)
                            If Not objValue Is Nothing Then
                                ' strProteinName was found in htProteinNameFirstList

                                blnKeepProtein = True
                                intIndex = CInt(objValue)

                                If intIndex >= 0 Then
                                    If udtProteinSeqHashInfo(intIndex).ProteinCount > 1 Then
                                        ' The protein has duplicate proteins
                                        ' Construct a list of the duplicate protein names

                                        intAdditionalProteinCount = 0

                                        For intDupIndex = 0 To udtProteinSeqHashInfo(intIndex).ProteinCount - 2
                                            ' Add the additional protein name if it is not of the form "BaseName-b", "BaseName-c", etc.
                                            blnSkipDupProtein = False

                                            If udtProteinSeqHashInfo(intIndex).AdditionalProteins(intDupIndex) Is Nothing Then
                                                blnSkipDupProtein = True
                                            Else
                                                If udtProteinSeqHashInfo(intIndex).AdditionalProteins(intDupIndex).ToLower = strProteinName.ToLower Then
                                                    ' Names match; do not add to the list
                                                    blnSkipDupProtein = True
                                                Else
                                                    With udtProteinSeqHashInfo(intIndex).AdditionalProteins(intDupIndex)
                                                        If .Length > 2 AndAlso .Chars(.Length - 2) = "-"c Then
                                                            If Char.IsLetter(.Chars(.Length - 1)) Then
                                                                If strProteinName.ToLower = udtProteinSeqHashInfo(intIndex).AdditionalProteins(intDupIndex).Substring(0, .Length - 2).ToLower Then
                                                                    ' Base names match; do not add to the list
                                                                    ' For example, ProteinX and ProteinX-b
                                                                    blnSkipDupProtein = True
                                                                End If
                                                            End If
                                                        End If
                                                    End With
                                                End If
                                            End If

                                            If Not blnSkipDupProtein Then
                                                If intAdditionalProteinCount >= strAdditionalProteinNames.Length Then
                                                    ReDim Preserve strAdditionalProteinNames(strAdditionalProteinNames.Length * 2 - 1)
                                                End If

                                                strAdditionalProteinNames(intAdditionalProteinCount) = String.Copy(udtProteinSeqHashInfo(intIndex).AdditionalProteins(intDupIndex))
                                                intAdditionalProteinCount += 1
                                            End If
                                        Next intDupIndex

                                        If intAdditionalProteinCount > 0 Then
                                            ' Append the duplicate protein names to the description
                                            strLineIn = ConstructFastaHeaderLine( _
                                                                strProteinName, _
                                                                strProteinDescription & "; Duplicate proteins: " & FlattenArray(strAdditionalProteinNames, intAdditionalProteinCount, ","c))
                                        End If
                                    End If
                                End If
                            Else
                                blnKeepProtein = False
                                mFixedFastaStats.DuplicateSequenceProteinsSkipped += 1

                                objValue = htDuplicateProteinList.Item(strProteinName.ToLower)
                                If objValue Is Nothing Then
                                    strMasterProteinInfo = "same as ??"
                                Else
                                    strMasterProteinInfo = "same as " & CStr(objValue)
                                End If
                                RecordFastaFileWarning(intLineCount, 0, strProteinName, eMessageCodeConstants.ProteinRemovedSinceDuplicateSequence, strMasterProteinInfo, String.Empty)
                            End If

                            If blnKeepProtein Then
                                swConsolidatedFastaOut.WriteLine(strLineIn)
                            End If

                        Else
                            ' Protein residues

                            If blnKeepProtein Then
                                swConsolidatedFastaOut.WriteLine(strLineIn)
                            End If
                        End If
                    End If
                End If
            Loop

            blnSuccess = True

        Catch ex As Exception
            RecordFastaFileError(0, 0, String.Empty, eMessageCodeConstants.UnspecifiedError, _
                "Error writing to consolidated fasta file " & strFixedFastaFilePath & ": " & ex.message, String.Empty)

            Return False
        Finally
            Try
                If Not srFastaInFile Is Nothing Then srFastaInFile.Close()
                If Not swConsolidatedFastaOut Is Nothing Then swConsolidatedFastaOut.Close()

                System.Threading.Thread.Sleep(100)

                System.IO.File.Delete(strFixedFastaFilePathTemp)
            Catch ex As Exception
                ' Ignore errors here
            End Try
        End Try

        Return blnSuccess

    End Function

    Protected Function ConstructFastaHeaderLine(ByRef strProteinName As String, ByRef strProteinDescription As String) As String

        If strProteinName Is Nothing Then strProteinName = "????"

        If strProteinDescription Is Nothing OrElse strProteinDescription.Length = 0 Then
            Return mProteinLineStartChar & strProteinName
        Else
            Return mProteinLineStartChar & strProteinName & " " & strProteinDescription
        End If

    End Function

    Protected Function ConstructStatsFilePath(ByVal strOutputFolderPath As String) As String

        Dim strStatsFilePath As String = String.Empty

        Try
            ' Record the current time in dtNow
			strStatsFilePath = "FastaFileStats_" & System.DateTime.Now.ToString("yyyy-MM-dd") & ".txt"

            If Not strOutputFolderPath Is Nothing AndAlso strOutputFolderPath.Length > 0 Then
                strStatsFilePath = System.IO.Path.Combine(strOutputFolderPath, strStatsFilePath)
            End If
        Catch ex As Exception
            If strStatsFilePath Is Nothing OrElse strStatsFilePath.Length = 0 Then
                strStatsFilePath = "FastaFileStats.txt"
            End If
        End Try

        Return strStatsFilePath

    End Function

    Private Function DetermineLineTerminatorSize(ByVal strInputFilePath As String) As Integer
        Dim endCharType As IValidateFastaFile.eLineEndingCharacters = _
            Me.DetermineLineTerminatorType(strInputFilePath)

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

    End Function

    Private Function DetermineLineTerminatorType(ByVal strInputFilePath As String) As IValidateFastaFile.eLineEndingCharacters
        Dim fsInFile As System.IO.FileStream
        Dim intByte As Integer

        Dim intTerminatorSize As Integer = 2
        Dim endCharacterType As IValidateFastaFile.eLineEndingCharacters

        Try
            ' Open the input file and look for the first carriage return or line feed
            fsInFile = New System.IO.FileStream( _
                strInputFilePath, _
                System.IO.FileMode.Open, _
                System.IO.FileAccess.Read, _
                IO.FileShare.Read)

            Do While fsInFile.Position < fsInFile.Length AndAlso fsInFile.Position < 100000

                intByte = fsInFile.ReadByte()

                If intByte = 10 Then
                    ' Found linefeed
                    If fsInFile.Position < fsInFile.Length Then
                        intByte = fsInFile.ReadByte()
                        If intByte = 13 Then
                            ' LfCr
                            intTerminatorSize = 2
                            endCharacterType = IValidateFastaFile.eLineEndingCharacters.LFCR
                        Else
                            ' Lf only
                            intTerminatorSize = 1
                            endCharacterType = IValidateFastaFile.eLineEndingCharacters.LF
                        End If
                    Else
                        intTerminatorSize = 1
                        endCharacterType = IValidateFastaFile.eLineEndingCharacters.LF
                    End If
                    Exit Do
                ElseIf intByte = 13 Then
                    ' Found carriage return
                    If fsInFile.Position < fsInFile.Length Then
                        intByte = fsInFile.ReadByte()
                        If intByte = 10 Then
                            ' CrLf
                            intTerminatorSize = 2
                            endCharacterType = IValidateFastaFile.eLineEndingCharacters.CRLF
                        Else
                            ' Cr only
                            intTerminatorSize = 1
                            endCharacterType = IValidateFastaFile.eLineEndingCharacters.CR
                        End If
                    Else
                        intTerminatorSize = 1
                        endCharacterType = IValidateFastaFile.eLineEndingCharacters.CR
                    End If
                    Exit Do
                End If

            Loop

        Catch ex As Exception
            SetLocalErrorCode(IValidateFastaFile.eValidateFastaFileErrorCodes.ErrorVerifyingLinefeedAtEOF)
        Finally
            If Not fsInFile Is Nothing Then
                fsInFile.Close()
            End If
        End Try

        Return endCharacterType

    End Function

    ' Unused function
    ''Protected Function ExtractListItem(ByVal strList As String, ByVal intItem As Integer) As String
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

    Protected Function NormalizeFileLineEndings( _
        ByVal pathOfFileToFix As String, _
        ByVal newFileName As String, _
        ByVal desiredLineEndCharacterType As IValidateFastaFile.eLineEndingCharacters) As String

        Dim newEndChar As String = ControlChars.CrLf

        Dim endCharType As IValidateFastaFile.eLineEndingCharacters = _
            Me.DetermineLineTerminatorType(pathOfFileToFix)

        Dim fi As System.IO.FileInfo
        Dim tr As System.IO.TextReader
        Dim s As String
        Dim origEndCharCount As Integer

        Dim fileLength As Long
        Dim currFilePos As Long
        Dim lineCount As Long

        Dim sw As System.IO.StreamWriter


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

            If Not System.IO.Path.IsPathRooted(newFileName) Then
                newFileName = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(pathOfFileToFix), System.IO.Path.GetFileName(newFileName))
            End If

            fi = New System.IO.FileInfo(pathOfFileToFix)
            fileLength = fi.Length

            tr = fi.OpenText

            sw = New System.IO.StreamWriter(newFileName)

            Me.OnProgressUpdate("Normalizing Line Endings...", 0.0)

            s = tr.ReadLine
            Do While Not s Is Nothing
                sw.Write(s)
                sw.Write(newEndChar)

                currFilePos = s.Length + origEndCharCount
                lineCount += 1

                If lineCount Mod 1000 = 0 Then
                    Me.OnProgressUpdate("Normalizing Line Endings (" & _
                        Math.Round(CDbl(currFilePos / fileLength * 100), 1).ToString & _
                        " % complete", CSng(currFilePos / fileLength * 100))
                End If

                s = tr.ReadLine
            Loop
            tr.Close()
            sw.Close()
            fi = Nothing

            Return newFileName
        Else
            Return pathOfFileToFix
        End If
    End Function

    Public Overrides Function GetDefaultExtensionsToParse() As String() Implements IValidateFastaFile.GetDefaultExtensionsToParse
        Dim strExtensionsToParse(0) As String

        strExtensionsToParse(0) = ".fasta"

        Return strExtensionsToParse

    End Function

    Private Sub EvaluateRules( _
        ByRef udtRuleDetails() As udtRuleDefinitionExtendedType, _
        ByVal strProteinName As String, _
        ByVal strTextToTest As String, _
        ByVal intTestTextOffsetInLine As Integer, _
        ByVal strEntireLine As String, _
        ByVal intContextLength As Integer)

        Dim intIndex As Integer
        Dim reMatch As System.Text.RegularExpressions.Match
        Dim strExtraInfo As String
        Dim intCharIndexOfMatch As Integer

        For intIndex = 0 To udtRuleDetails.Length - 1
            With udtRuleDetails(intIndex)

                reMatch = .reRule.Match(strTextToTest)

                If (.RuleDefinition.MatchIndicatesProblem And reMatch.Success) OrElse _
                    Not (.RuleDefinition.MatchIndicatesProblem And Not reMatch.Success) Then

                    If .RuleDefinition.DisplayMatchAsExtraInfo Then
                        strExtraInfo = reMatch.ToString
                    Else
                        strExtraInfo = String.Empty
                    End If

                    intCharIndexOfMatch = intTestTextOffsetInLine + reMatch.Index
                    If .RuleDefinition.Severity >= 5 Then
                        RecordFastaFileError(mLineCount, intCharIndexOfMatch, strProteinName, _
                            .RuleDefinition.CustomRuleID, strExtraInfo, _
                            ExtractContext(strEntireLine, intCharIndexOfMatch, intContextLength))
                    Else
                        RecordFastaFileWarning(mLineCount, intCharIndexOfMatch, strProteinName, _
                            .RuleDefinition.CustomRuleID, strExtraInfo, _
                            ExtractContext(strEntireLine, intCharIndexOfMatch, intContextLength))
                    End If

                End If

            End With
        Next intIndex

    End Sub

    Private Function ExtractContext(ByVal strText As String, ByVal intStartIndex As Integer) As String
        Return ExtractContext(strText, intStartIndex, DEFAULT_CONTEXT_LENGTH)
    End Function

    Private Function ExtractContext(ByVal strText As String, ByVal intStartIndex As Integer, ByVal intContextLength As Integer) As String
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

    Private Function FlattenAdditionalProteinList(ByRef udtProteinSeqHashEntry As udtProteinHashInfoType, ByVal chSepChar As Char) As String
        Return FlattenArray(udtProteinSeqHashEntry.AdditionalProteins, udtProteinSeqHashEntry.ProteinCount - 1, chSepChar)
    End Function

    Private Function FlattenArray(ByRef strArray() As String) As String
        Return FlattenArray(strArray, ControlChars.Tab)
    End Function

    Private Function FlattenArray(ByRef strArray() As String, ByVal chSepChar As Char) As String
        If strArray Is Nothing Then
            Return String.Empty
        Else
            Return FlattenArray(strArray, strArray.Length, chSepChar)
        End If
    End Function

    Private Function FlattenArray(ByRef strArray() As String, ByVal intDataCount As Integer, ByVal chSepChar As Char) As String
        Dim intIndex As Integer
        Dim strResult As String

        If strArray Is Nothing Then
            Return String.Empty
        ElseIf strArray.Length = 0 OrElse intDataCount <= 0 Then
            Return String.Empty
        Else
            If intDataCount > strArray.Length Then
                intDataCount = strArray.Length
            End If

            strResult = strArray(0)
            If strResult Is Nothing Then strResult = String.Empty

            For intIndex = 1 To intDataCount - 1
                If strArray(intIndex) Is Nothing Then
                    strResult &= chSepChar
                Else
                    strResult &= chSepChar & strArray(intIndex)
                End If
            Next intIndex
            Return strResult
        End If
    End Function

    Public Overrides Function GetErrorMessage() As String Implements IValidateFastaFile.GetCurrentErrorMessage
        ' Returns "" if no error

        Dim strErrorMessage As String

        If MyBase.ErrorCode = clsProcessFilesBaseClass.eProcessFilesErrorCodes.LocalizedError Or _
           MyBase.ErrorCode = clsProcessFilesBaseClass.eProcessFilesErrorCodes.NoError Then
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

    Protected Function GetFileErrorTextByIndex(ByVal intFileErrorIndex As Integer, ByVal strSepChar As String) As String

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

                Return LookupMessageType(IValidateFastaFile.eMsgTypeConstants.ErrorMsg) & strSepChar & _
                    "Line " & .LineNumber.ToString & strSepChar & _
                    "Col " & .ColNumber.ToString & strSepChar & _
                    strProteinName & strSepChar & _
                    LookupMessageDescription(.MessageCode, .ExtraInfo) & strSepChar & .Context

            End With
        End If

    End Function

    Protected Function GetFileErrorByIndex(ByVal intFileErrorIndex As Integer) As IValidateFastaFile.udtMsgInfoType

        If mFileErrorCount <= 0 Or intFileErrorIndex < 0 Or intFileErrorIndex >= mFileErrorCount Then
            Return New IValidateFastaFile.udtMsgInfoType
        Else
            Return mFileErrors(intFileErrorIndex)
        End If

    End Function

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


    Protected Function GetFileWarningTextByIndex(ByVal intFileWarningIndex As Integer, ByVal strSepChar As String) As String
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

                Return LookupMessageType(IValidateFastaFile.eMsgTypeConstants.WarningMsg) & strSepChar & _
                    "Line " & .LineNumber.ToString & strSepChar & _
                    "Col " & .ColNumber.ToString & strSepChar & _
                    strProteinName & strSepChar & _
                    LookupMessageDescription(.MessageCode, .ExtraInfo) & _
                    strSepChar & .Context

            End With
        End If

    End Function

    Protected Function GetFileWarningByIndex(ByVal intFileWarningIndex As Integer) As IValidateFastaFile.udtMsgInfoType

        If mFileWarningCount <= 0 Or intFileWarningIndex < 0 Or intFileWarningIndex >= mFileWarningCount Then
            Return New IValidateFastaFile.udtMsgInfoType
        Else
            Return mFileWarnings(intFileWarningIndex)
        End If

    End Function

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

    Private Function GetTimeStamp() As String
		' Record the current time
		Return System.DateTime.Now.ToShortDateString & " " & System.DateTime.Now.ToLongTimeString
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

    End Sub

    Private Sub InitializeRuleDetails( _
        ByRef udtRuleDefs() As udtRuleDefinitionType, _
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
                        .reRule = New System.Text.RegularExpressions.Regex( _
                            .RuleDefinition.MatchRegEx, _
                            System.Text.RegularExpressions.RegexOptions.Singleline Or _
                            System.Text.RegularExpressions.RegexOptions.Compiled)
                        .Valid = True
                    End With
                Catch ex As Exception
                    ' Ignore the error, but mark .Valid = false
                    udtRuleDetails(intIndex).Valid = False
                End Try
            Next intIndex
        End If

    End Sub

    Public Function LoadParameterFileSettings( _
        ByVal strParameterFilePath As String) As Boolean Implements IValidateFastaFile.LoadParameterFileSettings

        Dim objSettingsFile As New XmlSettingsFileAccessor

        Dim blnCustomRulesLoaded As Boolean
        Dim blnSuccess As Boolean

        Dim strCharacterList As String

        Try

            If strParameterFilePath Is Nothing OrElse strParameterFilePath.Length = 0 Then
                ' No parameter file specified; nothing to load
                Return True
            End If

            If Not System.IO.File.Exists(strParameterFilePath) Then
                ' See if strParameterFilePath points to a file in the same directory as the application
                strParameterFilePath = System.IO.Path.Combine( _
                    System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), _
                        System.IO.Path.GetFileName(strParameterFilePath))
                If Not System.IO.File.Exists(strParameterFilePath) Then
                    MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.ParameterFileNotFound)
                    Return False
                End If
            End If

            If objSettingsFile.LoadSettings(strParameterFilePath) Then
                If Not objSettingsFile.SectionPresent(XML_SECTION_OPTIONS) Then
                    If MyBase.ShowMessages Then
                        MsgBox("The node '<section name=""" & XML_SECTION_OPTIONS & """> was not found in the parameter file: " & _
                            strParameterFilePath, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Invalid File")
                    End If
                    MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.InvalidParameterFile)
                    Return False
                Else
                    ' Read customized settings

                    Me.OptionSwitch(IValidateFastaFile.SwitchOptions.AddMissingLinefeedatEOF) = _
                        objSettingsFile.GetParam(XML_SECTION_OPTIONS, "AddMissingLinefeedAtEOF", _
                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.AddMissingLinefeedatEOF))
                    Me.OptionSwitch(IValidateFastaFile.SwitchOptions.AllowAsteriskInResidues) = _
                        objSettingsFile.GetParam(XML_SECTION_OPTIONS, "AllowAsteriskInResidues", _
                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.AllowAsteriskInResidues))
                    Me.OptionSwitch(IValidateFastaFile.SwitchOptions.AllowDashInResidues) = _
                        objSettingsFile.GetParam(XML_SECTION_OPTIONS, "AllowDashInResidues", _
                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.AllowDashInResidues))
                    Me.OptionSwitch(IValidateFastaFile.SwitchOptions.CheckForDuplicateProteinNames) = _
                        objSettingsFile.GetParam(XML_SECTION_OPTIONS, "CheckForDuplicateProteinNames", _
                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.CheckForDuplicateProteinNames))
                    Me.OptionSwitch(IValidateFastaFile.SwitchOptions.CheckForDuplicateProteinSequences) = _
                        objSettingsFile.GetParam(XML_SECTION_OPTIONS, "CheckForDuplicateProteinSequences", _
                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.CheckForDuplicateProteinSequences))

                    Me.OptionSwitch(IValidateFastaFile.SwitchOptions.SaveProteinSequenceHashInfoFiles) = _
                        objSettingsFile.GetParam(XML_SECTION_OPTIONS, "SaveProteinSequenceHashInfoFiles", _
                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.SaveProteinSequenceHashInfoFiles))

                    Me.OptionSwitch(IValidateFastaFile.SwitchOptions.SaveBasicProteinHashInfoFile) = _
                        objSettingsFile.GetParam(XML_SECTION_OPTIONS, "SaveBasicProteinHashInfoFile", _
                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.SaveBasicProteinHashInfoFile))

                    Me.MaximumFileErrorsToTrack = objSettingsFile.GetParam(XML_SECTION_OPTIONS, _
                        "MaximumFileErrorsToTrack", Me.MaximumFileErrorsToTrack)
                    Me.MinimumProteinNameLength = objSettingsFile.GetParam(XML_SECTION_OPTIONS, _
                        "MinimumProteinNameLength", Me.MinimumProteinNameLength)
                    Me.MaximumProteinNameLength = objSettingsFile.GetParam(XML_SECTION_OPTIONS, _
                        "MaximumProteinNameLength", Me.MaximumProteinNameLength)
                    Me.MaximumResiduesPerLine = objSettingsFile.GetParam(XML_SECTION_OPTIONS, _
                        "MaximumResiduesPerLine", Me.MaximumResiduesPerLine)

                    Me.OptionSwitch(IValidateFastaFile.SwitchOptions.WarnBlankLinesBetweenProteins) = _
                        objSettingsFile.GetParam(XML_SECTION_OPTIONS, "WarnBlankLinesBetweenProteins", _
                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.WarnBlankLinesBetweenProteins))
                    Me.OptionSwitch(IValidateFastaFile.SwitchOptions.WarnLineStartsWithSpace) = _
                        objSettingsFile.GetParam(XML_SECTION_OPTIONS, "WarnLineStartsWithSpace", _
                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.WarnLineStartsWithSpace))

                    Me.OptionSwitch(IValidateFastaFile.SwitchOptions.OutputToStatsFile) = _
                        objSettingsFile.GetParam(XML_SECTION_OPTIONS, "OutputToStatsFile", _
                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.OutputToStatsFile))

                    Me.OptionSwitch(IValidateFastaFile.SwitchOptions.NormalizeFileLineEndCharacters) = _
                        objSettingsFile.GetParam(XML_SECTION_OPTIONS, "NormalizeFileLineEndCharacters", _
                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.NormalizeFileLineEndCharacters))


                    If Not objSettingsFile.SectionPresent(XML_SECTION_FIXED_FASTA_FILE_OPTIONS) Then
                        ' "ValidateFastaFixedFASTAFileOptions" section not present
                        ' Only read the settings for GenerateFixedFASTAFile and SplitOutMultipleRefsInProteinName

                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.GenerateFixedFASTAFile) = _
                            objSettingsFile.GetParam(XML_SECTION_OPTIONS, "GenerateFixedFASTAFile", _
                            Me.OptionSwitch(IValidateFastaFile.SwitchOptions.GenerateFixedFASTAFile))

                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.SplitOutMultipleRefsInProteinName) = _
                            objSettingsFile.GetParam(XML_SECTION_OPTIONS, "SplitOutMultipleRefsInProteinName", _
                            Me.OptionSwitch(IValidateFastaFile.SwitchOptions.SplitOutMultipleRefsInProteinName))

                    Else
                        ' "ValidateFastaFixedFASTAFileOptions" section is present

                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.GenerateFixedFASTAFile) = _
                            objSettingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "GenerateFixedFASTAFile", _
                            Me.OptionSwitch(IValidateFastaFile.SwitchOptions.GenerateFixedFASTAFile))

                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.SplitOutMultipleRefsInProteinName) = _
                            objSettingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "SplitOutMultipleRefsInProteinName", _
                            Me.OptionSwitch(IValidateFastaFile.SwitchOptions.SplitOutMultipleRefsInProteinName))

                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaRenameDuplicateNameProteins) = _
                            objSettingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "RenameDuplicateNameProteins", _
                            Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaRenameDuplicateNameProteins))

                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaConsolidateDuplicateProteinSeqs) = _
                            objSettingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ConsolidateDuplicateProteinSeqs", _
                            Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaConsolidateDuplicateProteinSeqs))

                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaConsolidateDupsIgnoreILDiff) = _
                            objSettingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "ConsolidateDupsIgnoreILDiff", _
                            Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaConsolidateDupsIgnoreILDiff))

                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaTruncateLongProteinNames) = _
                            objSettingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "TruncateLongProteinNames", _
                            Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaTruncateLongProteinNames))

                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaSplitOutMultipleRefsForKnownAccession) = _
                            objSettingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "SplitOutMultipleRefsForKnownAccession", _
                            Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaSplitOutMultipleRefsForKnownAccession))

                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaWrapLongResidueLines) = _
                            objSettingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "WrapLongResidueLines", _
                            Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaWrapLongResidueLines))

                        Me.OptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaRemoveInvalidResidues) = _
                            objSettingsFile.GetParam(XML_SECTION_FIXED_FASTA_FILE_OPTIONS, "RemoveInvalidResidues", _
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
                MsgBox("Error in LoadParameterFileSettings:" & ControlChars.NewLine & ex.Message, _
                    MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
            Else
                Throw New System.Exception("Error in LoadParameterFileSettings", ex)
            End If
            Return False
        End Try

        If Not blnCustomRulesLoaded Then
            SetDefaultRules()
        End If

        Return True

    End Function

    Public Function LookupMessageDescription(ByVal intErrorMessageCode As Integer) As String _
        Implements IValidateFastaFile.LookupMessageDescription
        Return Me.LookupMessageDescription(intErrorMessageCode, Nothing)
    End Function

    Public Function LookupMessageDescription(ByVal intErrorMessageCode As Integer, ByVal strExtraInfo As String) As String _
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

        If Not strExtraInfo Is Nothing AndAlso strExtraInfo.Length > 0 Then
            strMessage &= " (" & strExtraInfo & ")"
        End If

        Return strMessage

    End Function

    Protected Function LookupMessageType(ByVal EntryType As IValidateFastaFile.eMsgTypeConstants) As String _
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

    Protected Function SimpleProcessFile(ByVal strInputFilePath As String) As Boolean Implements IValidateFastaFile.ValidateFASTAFile
        ' Note that .ProcessFile returns True if a file is successfully processed (even if errors are found)
        Return Me.ProcessFile(strInputFilePath, Nothing, Nothing, False)
    End Function

    Public Overloads Overrides Function ProcessFile( _
        ByVal strInputFilePath As String, _
        ByVal strOutputFolderPath As String, _
        ByVal strParameterFilePath As String, _
        ByVal blnResetErrorCode As Boolean) As Boolean Implements IValidateFastaFile.ValidateFASTAFile

        'Returns True if success, False if failure

        Dim ioFile As System.IO.FileInfo
        Dim swStatsOutFile As System.IO.StreamWriter

        Dim strInputFilePathFull As String
        Dim strStatusMessage As String

        Dim blnSuccess As Boolean

        If blnResetErrorCode Then
            SetLocalErrorCode(IValidateFastaFile.eValidateFastaFileErrorCodes.NoError)
        End If

        If Not LoadParameterFileSettings(strParameterFilePath) Then
            strStatusMessage = "Parameter file load error: " & strParameterFilePath
            If MyBase.ShowMessages Then MsgBox(strStatusMessage, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
            Console.WriteLine(strStatusMessage)
            If MyBase.ErrorCode = clsProcessFilesBaseClass.eProcessFilesErrorCodes.NoError Then
                MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.InvalidParameterFile)
            End If
            Return False
        End If

        Try
            If strInputFilePath Is Nothing OrElse strInputFilePath.Length = 0 Then
                Console.WriteLine("Input file name is empty")
                MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.InvalidInputFilePath)
            Else

                Console.WriteLine()
                Console.WriteLine("Parsing " & System.IO.Path.GetFileName(strInputFilePath))

                If Not CleanupFilePaths(strInputFilePath, strOutputFolderPath) Then
                    MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.FilePathError)
                Else
                    Try

                        ' Obtain the full path to the input file
                        ioFile = New System.IO.FileInfo(strInputFilePath)
                        strInputFilePathFull = ioFile.FullName

                        blnSuccess = AnalyzeFastaFile(strInputFilePathFull)

                        If blnSuccess Then
                            ReportResults(strOutputFolderPath, mOutputToStatsFile)
                        Else
                            If mOutputToStatsFile Then
                                mStatsFilePath = ConstructStatsFilePath(strOutputFolderPath)
                                swStatsOutFile = New System.IO.StreamWriter(mStatsFilePath, True)
                                swStatsOutFile.WriteLine(GetTimeStamp() & ControlChars.Tab & "Error parsing " & _
                                    System.IO.Path.GetFileName(strInputFilePath) & ": " & Me.GetErrorMessage())
                                swStatsOutFile.Close()
                                swStatsOutFile = Nothing
                            Else
                                Console.WriteLine("Error parsing " & _
                                    System.IO.Path.GetFileName(strInputFilePath) & _
                                    ": " & Me.GetErrorMessage())
                            End If
                        End If
                    Catch ex As Exception
                        If MyBase.ShowMessages Then
                            MsgBox("Error calling AnalyzeFastaFile" & ControlChars.NewLine & ex.Message, _
                                MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
                        Else
                            Throw New System.Exception("Error calling AnalyzeFastaFile", ex)
                        End If
                    End Try
                End If
            End If
        Catch ex As Exception
            If MyBase.ShowMessages Then
                MsgBox("Error in ProcessFile:" & ControlChars.NewLine & ex.Message, _
                    MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
            Else
                Throw New System.Exception("Error in ProcessFile", ex)
            End If
        End Try

        Return blnSuccess

    End Function

    Private Sub PrependExtraTextToProteinDescription(ByVal strExtraProteinNameText As String, ByRef strProteinDescription As String)
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


    Private Sub ProcessResiduesForPreviousProtein( _
                        ByVal strProteinName As String, _
                        ByRef sbCurrentResidues As Text.StringBuilder, _
                        ByRef htProteinSequenceHashes As Hashtable, _
                        ByRef intProteinSequenceHashCount As Integer, _
                        ByRef udtProteinSeqHashInfo() As udtProteinHashInfoType, _
                        ByVal blnConsolidateDupsIgnoreILDiff As Boolean, _
                        ByRef swFixedFastaOut As System.IO.StreamWriter, _
                        ByVal intCurrentValidResidueLineLengthMax As Integer, _
                        ByRef swProteinSequenceHashBasic As System.IO.StreamWriter)

        Dim intWrapLength As Integer
        Dim intResidueCount As Integer

        Dim intIndex As Integer
        Dim intLength As Integer

        If sbCurrentResidues.Length > 0 Then
            If mCheckForDuplicateProteinSequences OrElse mSaveBasicProteinHashInfoFile Then
                ' Process the previous protein entry to store a hash of the protein sequence
                ProcessSequenceHashInfo(strProteinName, sbCurrentResidues, htProteinSequenceHashes, intProteinSequenceHashCount, udtProteinSeqHashInfo, blnConsolidateDupsIgnoreILDiff, swProteinSequenceHashBasic)
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

    Private Sub ProcessSequenceHashInfo( _
                        ByVal strProteinName As String, _
                        ByRef sbCurrentResidues As Text.StringBuilder, _
                        ByRef htProteinSequenceHashes As Hashtable, _
                        ByRef intProteinSequenceHashCount As Integer, _
                        ByRef udtProteinSeqHashInfo() As udtProteinHashInfoType, _
                        ByVal blnConsolidateDupsIgnoreILDiff As Boolean, _
                        ByRef swProteinSequenceHashBasic As System.IO.StreamWriter)

        Static objHashGenerator As clsHashGenerator

        Dim objHashtableValue As Object
        Dim strComputedHash As String

        If objHashGenerator Is Nothing Then
            objHashGenerator = New clsHashGenerator
        End If

        Try
            If sbCurrentResidues.Length > 0 Then
                ' Compute the hash value for sbCurrentResidues
                If blnConsolidateDupsIgnoreILDiff Then
                    strComputedHash = objHashGenerator.GenerateHash(sbCurrentResidues.ToString.Replace("L"c, "I"c))
                Else
                    strComputedHash = objHashGenerator.GenerateHash(sbCurrentResidues.ToString)
                End If

                If Not swProteinSequenceHashBasic Is Nothing Then
                    swProteinSequenceHashBasic.WriteLine(mProteinCount.ToString & ControlChars.Tab & _
                                                         strProteinName & ControlChars.Tab & _
                                                         sbCurrentResidues.Length.ToString & ControlChars.Tab & _
                                                         strComputedHash)
                End If

                If mCheckForDuplicateProteinSequences AndAlso Not htProteinSequenceHashes Is Nothing Then
                    ' See if htProteinSequenceHashes contains strHash
                    objHashtableValue = htProteinSequenceHashes(strComputedHash)
                    If Not objHashtableValue Is Nothing Then
                        ' Value exists; update the entry in udtProteinSeqHashInfo
                        With udtProteinSeqHashInfo(CInt(objHashtableValue))
                            If .ProteinCount >= .AdditionalProteins.Length Then
                                ReDim Preserve .AdditionalProteins(.AdditionalProteins.Length * 2 - 1)
                            End If
                            .AdditionalProteins(.ProteinCount - 1) = String.Copy(strProteinName)
                            .ProteinCount += 1
                        End With
                    Else
                        ' Value not yet present; add it

                        If intProteinSequenceHashCount >= udtProteinSeqHashInfo.Length Then
                            ' Need to reserve more space in udtProteinSeqHashInfo
                            ReDim Preserve udtProteinSeqHashInfo(udtProteinSeqHashInfo.Length * 2 - 1)
                        End If

                        With udtProteinSeqHashInfo(intProteinSequenceHashCount)
                            .ProteinCount = 1
                            .ProteinNameFirst = String.Copy(strProteinName)
                            ReDim .AdditionalProteins(4)
                            .SequenceHash = String.Copy(strComputedHash)
                            .SequenceLength = sbCurrentResidues.Length
                            .SequenceStart = sbCurrentResidues.ToString.Substring(0, Math.Min(sbCurrentResidues.Length, 20))
                        End With

                        htProteinSequenceHashes.Add(strComputedHash, intProteinSequenceHashCount)
                        intProteinSequenceHashCount += 1
                    End If

                End If

            End If
        Catch ex As Exception
            'Error caught; pass it up to the calling function
            Throw ex
        End Try


    End Sub

    Private Function ReadRulesFromParameterFile( _
        ByRef objSettingsFile As XmlSettingsFileAccessor, _
        ByVal strSectionName As String, _
        ByRef udtRules() As udtRuleDefinitionType) As Boolean
        ' Returns True if the section named strSectionName is present and if it contains an item with keyName = "RuleCount"
        ' Note: even if RuleCount = 0, this function will return True

        Dim blnSuccess As Boolean = False
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

    Private Sub RecordFastaFileError( _
        ByVal intLineNumber As Integer, _
        ByVal intCharIndex As Integer, _
        ByVal strProteinName As String, _
        ByVal intErrorMessageCode As Integer)

        RecordFastaFileError(intLineNumber, intCharIndex, strProteinName, _
            intErrorMessageCode, String.Empty, String.Empty)
    End Sub

    Private Sub RecordFastaFileError( _
        ByVal intLineNumber As Integer, _
        ByVal intCharIndex As Integer, _
        ByVal strProteinName As String, _
        ByVal intErrorMessageCode As Integer, _
        ByVal strExtraInfo As String, _
        ByVal strContext As String)
        RecordFastaFileProblemWork( _
            mFileErrorStats, _
            mFileErrorCount, _
            mFileErrors, _
            intLineNumber, _
            intCharIndex, _
            strProteinName, _
            intErrorMessageCode, _
            strExtraInfo, strContext)
    End Sub

    Private Sub RecordFastaFileWarning( _
        ByVal intLineNumber As Integer, _
        ByVal intCharIndex As Integer, _
        ByVal strProteinName As String, _
        ByVal intWarningMessageCode As Integer)
        RecordFastaFileWarning( _
            intLineNumber, _
            intCharIndex, _
            strProteinName, _
            intWarningMessageCode, _
            String.Empty, String.Empty)
    End Sub

    Private Sub RecordFastaFileWarning( _
        ByVal intLineNumber As Integer, _
        ByVal intCharIndex As Integer, _
        ByVal strProteinName As String, _
        ByVal intWarningMessageCode As Integer, _
        ByVal strExtraInfo As String, ByVal strContext As String)

        RecordFastaFileProblemWork(mFileWarningStats, mFileWarningCount, _
            mFileWarnings, intLineNumber, intCharIndex, strProteinName, _
            intWarningMessageCode, strExtraInfo, strContext)
    End Sub

    Private Sub RecordFastaFileProblemWork( _
        ByRef udtItemSummaryIndexed As udtItemSummaryIndexedType, _
        ByRef intItemCountSpecified As Integer, _
        ByRef udtItems() As IValidateFastaFile.udtMsgInfoType, _
        ByVal intLineNumber As Integer, _
        ByVal intCharIndex As Integer, _
        ByVal strProteinName As String, _
        ByVal intMessageCode As Integer, _
        ByVal strExtraInfo As String, _
        ByVal strContext As String)

        ' Note that intCharIndex is the index in the source string at which the error occurred
        ' When storing in .ColNumber, we add 1 to intCharIndex

        ' Lookup the index of the entry with intMessageCode in udtItemSummaryIndexed.ErrorStats
        ' Add it if not present

        Dim objItemIndex As Object
        Dim intItemIndex As Integer

        Try
            With udtItemSummaryIndexed
                intItemIndex = -1
                If .htMessageCodeToArrayIndex.Count > 0 Then

                    objItemIndex = .htMessageCodeToArrayIndex(intMessageCode)
                    If Not objItemIndex Is Nothing Then
                        intItemIndex = CInt(objItemIndex)
                    End If
                End If

                If intItemIndex < 0 Then
                    If .ErrorStats.Length <= 0 Then
                        ReDim .ErrorStats(1)
                    ElseIf .ErrorStatsCount = .ErrorStats.Length Then
                        ReDim Preserve .ErrorStats(.ErrorStats.Length * 2 - 1)
                    End If
                    intItemIndex = .ErrorStatsCount
                    .ErrorStats(intItemIndex).MessageCode = intMessageCode
                    .htMessageCodeToArrayIndex.Add(intMessageCode, intItemIndex)
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
            Console.WriteLine("Error in RecordFastaFileProblemWork: " & ex.Message)
        End Try

    End Sub

    Private Sub ReplaceXMLCodesWithText(ByVal strParameterFilePath As String)

        Dim srInFile As System.IO.StreamReader
        Dim srOutFile As System.IO.StreamWriter

        Dim strOutputFilePath As String
        Dim strTimeStamp As String
        Dim strLineIn As String

        Try
            ' Define the output file path
            strTimeStamp = GetTimeStamp().Replace(" ", "_").Replace(":", "_").Replace("/", "_")

            strOutputFilePath = strParameterFilePath & "_" & strTimeStamp & ".fixed"

            ' Open the input file
            srInFile = New System.IO.StreamReader(strParameterFilePath)

            ' Open the output file
            srOutFile = New System.IO.StreamWriter(strOutputFilePath, False)

            ' Parse each line in the file
            Do While srInFile.Peek() >= 0
                strLineIn = srInFile.ReadLine()

                If Not strLineIn Is Nothing Then
                    strLineIn = strLineIn.Replace("&gt;", ">").Replace("&lt;", "<")
                    srOutFile.WriteLine(strLineIn)
                End If
            Loop

            ' Close the input and output files
            srInFile.Close()
            srOutFile.Close()

            ' Wait 100 msec
            System.Threading.Thread.Sleep(100)

            ' Delete the input file
            System.IO.File.Delete(strParameterFilePath)

            ' Wait 250 msec
            System.Threading.Thread.Sleep(250)

            ' Rename the output file to the input file
            System.IO.File.Move(strOutputFilePath, strParameterFilePath)

        Catch ex As Exception
            If MyBase.ShowMessages Then
                MsgBox("Error in ReplaceXMLCodesWithText:" & ControlChars.NewLine & ex.Message, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
            Else
                Throw New System.Exception("Error in ReplaceXMLCodesWithText", ex)
            End If
        Finally
            If Not srInFile Is Nothing Then
                srInFile.Close()
            End If

            If Not srOutFile Is Nothing Then
                srOutFile.Close()
            End If
        End Try

    End Sub

    Protected Sub ReportResults( _
        ByVal strOutputFolderPath As String, _
        ByVal blnOutputToStatsFile As Boolean)

        Dim objOutStream As System.IO.FileStream
        Dim srOutFile As System.IO.StreamWriter

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
                strSourceFile = System.IO.Path.GetFileName(mFastaFilePath)
            Catch ex As Exception
                strSourceFile = "Unknown_filename_due_to_error.fasta"
            End Try

            If blnOutputToStatsFile Then
                mStatsFilePath = ConstructStatsFilePath(strOutputFolderPath)
                blnFileAlreadyExists = System.IO.File.Exists(mStatsFilePath)

                blnSuccess = False
                intRetryCount = 0

                Do While Not blnSuccess And intRetryCount < 5
                    Try
                        objOutStream = New System.IO.FileStream(mStatsFilePath, IO.FileMode.Append, IO.FileAccess.Write, IO.FileShare.Read)
                        srOutFile = New System.IO.StreamWriter(objOutStream)
                        blnSuccess = True
                    Catch ex As Exception
                        ' Failed to open file, wait 1 second, then try again

                        srOutFile = Nothing
                        If Not objOutStream Is Nothing Then
                            objOutStream.Close()
                            objOutStream = Nothing
                        End If

                        intRetryCount += 1
                        System.Threading.Thread.Sleep(1000)
                    End Try
                Loop

                If blnSuccess Then
                    strSepChar = ControlChars.Tab
                    If Not blnFileAlreadyExists Then
                        ' Write the header line
                        srOutFile.WriteLine( _
                            "Date" & strSepChar & _
                            "SourceFile" & strSepChar & _
                            "MessageType" & strSepChar & _
                            "LineNumber" & strSepChar & _
                            "ColumnNumber" & strSepChar & _
                            "Description_or_Protein" & strSepChar & _
                            "Info" & strSepChar & _
                            "Context")
                    End If
                Else
                    strSepChar = ", "
                    blnOutputToStatsFile = False
                    SetLocalErrorCode(IValidateFastaFile.eValidateFastaFileErrorCodes.ErrorCreatingStatsFile)
                End If
            Else
                strSepChar = ", "
            End If

            ReportResultAddEntry( _
                strSourceFile, _
                IValidateFastaFile.eMsgTypeConstants.StatusMsg, _
                "Full path to file", _
                mFastaFilePath, _
                String.Empty, _
                blnOutputToStatsFile, _
                srOutFile, _
                strSepChar)

            ReportResultAddEntry( _
                strSourceFile, _
                IValidateFastaFile.eMsgTypeConstants.StatusMsg, _
                "Protein count", _
                mProteinCount.ToString, _
                String.Empty, _
                blnOutputToStatsFile, _
                srOutFile, _
                strSepChar)

            ReportResultAddEntry( _
                strSourceFile, _
                IValidateFastaFile.eMsgTypeConstants.StatusMsg, _
                "Residue count", _
                mResidueCount.ToString, _
                String.Empty, _
                blnOutputToStatsFile, _
                srOutFile, _
                strSepChar)

            If mFileErrorCount > 0 Then
                ReportResultAddEntry( _
                    strSourceFile, _
                    IValidateFastaFile.eMsgTypeConstants.ErrorMsg, _
                    "Error count", _
                    Me.ErrorWarningCounts( _
                        IValidateFastaFile.eMsgTypeConstants.ErrorMsg, _
                        IValidateFastaFile.ErrorWarningCountTypes.Total).ToString, _
                    String.Empty, _
                    blnOutputToStatsFile, _
                    srOutFile, strSepChar)

                If mFileErrorCount > 1 Then
                    iErrorInfoComparerClass = New ErrorInfoComparerClass
                    Array.Sort(mFileErrors, 0, mFileErrorCount, iErrorInfoComparerClass)
                    iErrorInfoComparerClass = Nothing
                End If

                For intIndex = 0 To mFileErrorCount - 1
                    With mFileErrors(intIndex)
                        If .ProteinName Is Nothing OrElse .ProteinName.Length = 0 Then
                            strProteinName = "N/A"
                        Else
                            strProteinName = String.Copy(.ProteinName)
                        End If

                        ReportResultAddEntry(strSourceFile, _
                                             IValidateFastaFile.eMsgTypeConstants.ErrorMsg, _
                                             .LineNumber, _
                                             .ColNumber, _
                                             strProteinName, _
                                             LookupMessageDescription(.MessageCode, .ExtraInfo), _
                                             .Context, _
                                             blnOutputToStatsFile, _
                                             srOutFile, _
                                             strSepChar)

                    End With
                Next intIndex
            End If

            If mFileWarningCount > 0 Then
                ReportResultAddEntry( _
                    strSourceFile, _
                    IValidateFastaFile.eMsgTypeConstants.WarningMsg, _
                    "Warning count", _
                    Me.ErrorWarningCounts( _
                        IValidateFastaFile.eMsgTypeConstants.WarningMsg, _
                        IValidateFastaFile.ErrorWarningCountTypes.Total).ToString, _
                    String.Empty, _
                    blnOutputToStatsFile, _
                    srOutFile, _
                    strSepChar)

                If mFileWarningCount > 1 Then
                    iErrorInfoComparerClass = New ErrorInfoComparerClass
                    Array.Sort(mFileWarnings, 0, mFileWarningCount, iErrorInfoComparerClass)
                    iErrorInfoComparerClass = Nothing
                End If

                For intIndex = 0 To mFileWarningCount - 1
                    With mFileWarnings(intIndex)
                        If .ProteinName Is Nothing OrElse .ProteinName.Length = 0 Then
                            strProteinName = "N/A"
                        Else
                            strProteinName = String.Copy(.ProteinName)
                        End If

                        ReportResultAddEntry(strSourceFile, _
                                            IValidateFastaFile.eMsgTypeConstants.WarningMsg, _
                                            .LineNumber, _
                                            .ColNumber, _
                                            strProteinName, _
                                            LookupMessageDescription(.MessageCode, .ExtraInfo), _
                                            .Context, _
                                            blnOutputToStatsFile, _
                                            srOutFile, _
                                            strSepChar)

                    End With
                Next intIndex
            End If

            If blnOutputToStatsFile Then
                srOutFile.Close()
                srOutFile = Nothing
            End If

        Catch ex As Exception
            If MyBase.ShowMessages Then
                MsgBox("Error in ReportResults:" & ControlChars.NewLine & ex.Message, _
                    MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
            Else
                Throw New System.Exception("Error in ReportResults", ex)
            End If

        End Try

    End Sub

    Private Sub ReportResultAddEntry( _
        ByVal strSourceFile As String, _
        ByVal EntryType As IValidateFastaFile.eMsgTypeConstants, _
        ByVal strDescriptionOrProteinName As String, _
        ByVal strInfo As String, _
        ByVal strContext As String, _
        ByVal blnOutputToStatsFile As Boolean, _
        ByVal srOutFile As System.IO.StreamWriter, _
        ByVal strSepChar As String)

        ReportResultAddEntry( _
            strSourceFile, _
            EntryType, 0, 0, _
            strDescriptionOrProteinName, _
            strInfo, _
            strContext, _
            blnOutputToStatsFile, _
            srOutFile, strSepChar)
    End Sub

    Private Sub ReportResultAddEntry( _
        ByVal strSourceFile As String, _
        ByVal EntryType As IValidateFastaFile.eMsgTypeConstants, _
        ByVal intLineNumber As Integer, _
        ByVal intColNumber As Integer, _
        ByVal strDescriptionOrProteinName As String, _
        ByVal strInfo As String, _
        ByVal strContext As String, _
        ByVal blnOutputToStatsFile As Boolean, _
        ByVal srOutFile As System.IO.StreamWriter, _
        ByVal strSepChar As String)

        Dim strMessage As String

        strMessage = strSourceFile & strSepChar & _
            LookupMessageType(EntryType) & strSepChar & _
            intLineNumber.ToString & strSepChar & _
            intColNumber.ToString & strSepChar & _
            strDescriptionOrProteinName & strSepChar & _
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
            If .htMessageCodeToArrayIndex Is Nothing Then
                .htMessageCodeToArrayIndex = New Hashtable
            Else
                .htMessageCodeToArrayIndex.Clear()
            End If
        End With

    End Sub

    Private Sub SaveRulesToParameterFile(ByRef objSettingsFile As XmlSettingsFileAccessor, ByVal strSectionName As String, ByRef udtRules() As udtRuleDefinitionType)

        Dim intRuleNumber As Integer
        Dim strRuleBase As String

        If udtRules Is Nothing OrElse udtRules.Length <= 0 Then
            objSettingsFile.SetParam(strSectionName, XML_OPTION_ENTRY_RULE_COUNT, 0)
        Else
            objSettingsFile.SetParam(strSectionName, XML_OPTION_ENTRY_RULE_COUNT, udtRules.Length)

            For intRuleNumber = 1 To udtRules.Length
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

    Public Function SaveSettingsToParameterFile(ByVal strParameterFilePath As String) As Boolean Implements IValidateFastaFile.SaveParameterSettingsToParameterFile
        ' Save a model parameter file

        Dim srOutFile As System.IO.StreamWriter
        Dim objSettingsFile As New XmlSettingsFileAccessor

        Try

            If strParameterFilePath Is Nothing OrElse strParameterFilePath.Length = 0 Then
                ' No parameter file specified; do not save the settings
                Return True
            End If

            If Not System.IO.File.Exists(strParameterFilePath) Then
                ' Need to generate a blank XML settings file

                srOutFile = New System.IO.StreamWriter(strParameterFilePath, False)

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
                MsgBox("Error in SaveSettingsToParameterFile:" & ControlChars.NewLine & ex.Message, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
            Else
                Throw New System.Exception("Error in SaveSettingsToParameterFile", ex)
            End If
            Return False
        End Try

        Return True

    End Function

    Private Function SearchRulesForID(ByRef udtRules() As udtRuleDefinitionType, ByVal intErrorMessageCode As Integer, ByRef strMessage As String) As Boolean

        Dim intIndex As Integer
        Dim blnMatchFound As Boolean = False

        If Not udtRules Is Nothing Then
            For intIndex = 0 To udtRules.Length - 1
                If udtRules(intIndex).CustomRuleID = intErrorMessageCode Then
                    strMessage = udtRules(intIndex).MessageWhenProblem
                    blnMatchFound = True
                    Exit For
                End If
            Next intIndex
        End If

        Return blnMatchFound

    End Function

    Public Sub SetDefaultRules() Implements IValidateFastaFile.SetDefaultRules

        Me.ClearAllRules()

        ' Header line errors
        Me.SetRule(IValidateFastaFile.RuleTypes.HeaderLine, "^>[ \t]*$", True, "Line starts with > but does not contain a protein name", 7)
        Me.SetRule(IValidateFastaFile.RuleTypes.HeaderLine, "^>[ \t].+", True, "Space or tab found directly after the > symbol", 7)

        ' Header line warnings
        Me.SetRule(IValidateFastaFile.RuleTypes.HeaderLine, "^>[^ \t]+[ \t]*$", True, MESSAGE_TEXT_PROTEIN_DESCRIPTION_MISSING, 3)
        Me.SetRule(IValidateFastaFile.RuleTypes.HeaderLine, "^>[^ \t]+\t", True, "Protein name is separated from the protein description by a tab", 3)

        ' Protein Name error characters
        Me.SetRule(IValidateFastaFile.RuleTypes.ProteinName, "[^A-Za-z0-9.\-_:,\|/()\[\]]", True, "Protein name contains invalid characters", 7, True)

        ' Protein name warnings
        Me.SetRule(IValidateFastaFile.RuleTypes.ProteinName, "[:|][^:|;]*[:|;]", True, "Protein name contains two or more vertical bars", 4, True)
        Me.SetRule(IValidateFastaFile.RuleTypes.ProteinName, "[/()\[\],]", True, "Protein name contains undesirable characters", 3, True)

        ' Protein description warnings
        Me.SetRule(IValidateFastaFile.RuleTypes.ProteinDescription, """", True, "Protein description contains a quotation mark", 3)
        Me.SetRule(IValidateFastaFile.RuleTypes.ProteinDescription, "\t", True, "Protein description contains a tab character", 3)
        Me.SetRule(IValidateFastaFile.RuleTypes.ProteinDescription, "\\/", True, "Protein description contains an escaped slash: \/", 3)
        Me.SetRule(IValidateFastaFile.RuleTypes.ProteinDescription, "[\x00-\x08\x0E-\x1F]", True, "Protein description contains an escape code character", 7)
        Me.SetRule(IValidateFastaFile.RuleTypes.ProteinDescription, ".{900,}", True, MESSAGE_TEXT_PROTEIN_DESCRIPTION_TOO_LONG, 4, False)

        ' Protein sequence errors
        Me.SetRule(IValidateFastaFile.RuleTypes.ProteinSequence, "[ \t]", True, "A space or tab was found in the residues", 7)

        If Not mAllowAsteriskInResidues Then
            Me.SetRule(IValidateFastaFile.RuleTypes.ProteinSequence, "\*", True, MESSAGE_TEXT_ASTERISK_IN_RESIDUES, 7)
        End If

        If Not mAllowDashInResidues Then
            Me.SetRule(IValidateFastaFile.RuleTypes.ProteinSequence, "\-", True, MESSAGE_TEXT_DASH_IN_RESIDUES, 7)
        End If

        ' Note: we look for a space, tab, asterisk, and dash with separate rules (defined above), so we
        ' therefore include them in this RegEx
        Me.SetRule(IValidateFastaFile.RuleTypes.ProteinSequence, "[^A-IK-Z \t\*\-]", True, "Invalid residues found", 7, True)

        ' Protein sequence warnings
        Me.SetRule(IValidateFastaFile.RuleTypes.ProteinSequence, "U", True, "Residues line contains U (selenocysteine); this residue is unsupported by Sequest", 3)

    End Sub

    Protected Sub SetLocalErrorCode(ByVal eNewErrorCode As IValidateFastaFile.eValidateFastaFileErrorCodes)
        SetLocalErrorCode(eNewErrorCode, False)
    End Sub

    Protected Sub SetLocalErrorCode( _
        ByVal eNewErrorCode As IValidateFastaFile.eValidateFastaFileErrorCodes, _
        ByVal blnLeaveExistingErrorCodeUnchanged As Boolean)

        If blnLeaveExistingErrorCodeUnchanged AndAlso mLocalErrorCode <> IValidateFastaFile.eValidateFastaFileErrorCodes.NoError Then
            ' An error code is already defined; do not change it
        Else
            mLocalErrorCode = eNewErrorCode

            If eNewErrorCode = IValidateFastaFile.eValidateFastaFileErrorCodes.NoError Then
                If MyBase.ErrorCode = clsProcessFilesBaseClass.eProcessFilesErrorCodes.LocalizedError Then
                    MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.NoError)
                End If
            Else
                MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.LocalizedError)
            End If
        End If

    End Sub

    Protected Sub SetRule( _
        ByVal ruleType As IValidateFastaFile.RuleTypes, _
        ByVal regexToMatch As String, _
        ByVal doesMatchIndicateProblem As Boolean, _
        ByVal problemReturnMessage As String, _
        ByVal severityLevel As Short) Implements IValidateFastaFile.SetRule

        Me.SetRule( _
            ruleType, regexToMatch, _
            doesMatchIndicateProblem, _
            problemReturnMessage, _
            severityLevel, False)

    End Sub

    Protected Sub SetRule( _
        ByVal ruleType As IValidateFastaFile.RuleTypes, _
        ByVal regexToMatch As String, _
        ByVal doesMatchIndicateProblem As Boolean, _
        ByVal problemReturnMessage As String, _
        ByVal severityLevel As Short, _
        ByVal displayMatchAsExtraInfo As Boolean) Implements IValidateFastaFile.SetRule

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

    Private Sub SetRule(ByRef udtRules() As udtRuleDefinitionType, ByVal strMatchRegEx As String, ByVal blnMatchIndicatesProblem As Boolean, ByVal strMessageWhenProblem As String, ByVal bytSeverity As Short, ByVal blnDisplayMatchAsExtraInfo As Boolean)

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

    Protected Sub SplitFastaProteinHeaderLine( _
                    ByRef strHeaderLine As String, _
                    ByRef strProteinName As String, _
                    ByRef strProteinDescription As String, _
                    ByRef intDescriptionStartIndex As Integer)

        Dim intSpaceIndex As Integer
        Dim intTabIndex As Integer

        strProteinName = String.Empty
        strProteinDescription = String.Empty
        intDescriptionStartIndex = 0

        ' Make sure the protein name and description are valid
        ' Find the first space and/or tab
        intSpaceIndex = strHeaderLine.IndexOf(" "c)
        intTabIndex = strHeaderLine.IndexOf(ControlChars.Tab)

        If intSpaceIndex = 1 Then
            ' Space found directly after the > symbol
        ElseIf intTabIndex > 0 Then
            If intTabIndex = 1 Then
                ' Tab character found directly after the > symbol
                intSpaceIndex = intTabIndex
            Else
                ' Tab character found; does it separate the protein name and description?
                If intSpaceIndex <= 0 OrElse (intSpaceIndex > 0 AndAlso intTabIndex < intSpaceIndex) Then
                    intSpaceIndex = intTabIndex
                End If
            End If
        End If

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

    Protected Function VerifyLinefeedAtEOF(ByVal strInputFilePath As String, ByVal blnAddCrLfIfMissing As Boolean) As Boolean
        Dim fsInFile As System.IO.FileStream
        Dim intByte As Integer
        Dim bytOneByte As Byte

        Dim blnNeedToAddCrLf As Boolean
        Dim blnSuccess As Boolean

        Try
            ' Open the input file and validate that the final characters are CrLf, simply CR, or simply LF
            fsInFile = New System.IO.FileStream(strInputFilePath, System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite)

            If fsInFile.Length > 2 Then
                fsInFile.Seek(-1, IO.SeekOrigin.End)

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
                    Console.WriteLine("Appending CrLf return to: " & System.IO.Path.GetFileName(strInputFilePath))
                    bytOneByte = CType(13, Byte)
                    fsInFile.WriteByte(bytOneByte)

                    bytOneByte = CType(10, Byte)
                    fsInFile.WriteByte(bytOneByte)
                    blnSuccess = True
                Else
                    blnSuccess = False
                End If
            Else
                blnSuccess = True
            End If

            blnSuccess = True

        Catch ex As Exception
            SetLocalErrorCode(IValidateFastaFile.eValidateFastaFileErrorCodes.ErrorVerifyingLinefeedAtEOF)
            blnSuccess = False
        Finally
            If Not fsInFile Is Nothing Then
                fsInFile.Close()
            End If
        End Try

        Return blnSuccess

    End Function

    ' IComparer class to allow comparison of udtMsgInfoType items
    Private Class ErrorInfoComparerClass
        Implements IComparer

        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare

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
