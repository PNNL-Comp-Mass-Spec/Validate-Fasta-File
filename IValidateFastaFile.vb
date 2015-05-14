Public Interface IValidateFastaFile

    Event ProgressChanged(
        ByVal taskDescription As String,
        ByVal percentComplete As Single)
    Event ProgressCompleted()
    Event WroteLineEndNormalizedFASTA(ByVal newFilePath As String)

    Function ValidateFASTAFile(ByVal filePath As String) As Boolean

    Function ValidateFASTAFile(
        ByVal filePath As String,
        ByVal errorDumpOutputPath As String,
        ByVal parameterFilePath As String,
        ByVal resetErrorCodes As Boolean) As Boolean

    Function GetDefaultExtensionsToParse() As String()
    Function GetCurrentErrorMessage() As String

    Function SaveParameterSettingsToParameterFile(ByVal parameterFilePath As String) As Boolean
    Function LoadParameterFileSettings(ByVal parameterFilePath As String) As Boolean

    Function LookupMessageDescription(ByVal errorMessageCode As Integer) As String
    Function LookupMessageDescription(ByVal errorMessageCode As Integer, ByVal ExtraInfo As String) As String
    Function LookupMessageTypeString(ByVal entryType As eMsgTypeConstants) As String


    Sub AbortProcessingNow()

    Sub ClearAllRules()
    Sub ClearRules(ByVal ruleType As RuleTypes)

    Sub SetDefaultRules()

    Sub SetRule(
        ByVal ruleType As RuleTypes,
        ByVal regexToMatch As String,
        ByVal doesMatchIndicateProblem As Boolean,
        ByVal problemReturnMessage As String,
        ByVal severityLevel As Short)
    Sub SetRule(
        ByVal ruleType As RuleTypes,
        ByVal regexToMatch As String,
        ByVal doesMatchIndicateProblem As Boolean,
        ByVal problemReturnMessage As String,
        ByVal severityLevel As Short,
        ByVal displayMatchAsExtraInfo As Boolean)

    Property OptionSwitches(ByVal SwitchName As SwitchOptions) As Boolean

    ReadOnly Property FixedFASTAFileStats(
        ByVal ValueType As FixedFASTAFileValues) As Integer
    

    ReadOnly Property ErrorMessageTextByIndex(
        ByVal errorIndex As Integer,
        ByVal valueSeparator As String) As String
    ReadOnly Property FileErrorByIndex(ByVal errorIndex As Integer) As udtMsgInfoType
    ReadOnly Property FileErrorList As udtMsgInfoType()

    ReadOnly Property WarningMessageTextByIndex(
        ByVal warningIndex As Integer,
        ByVal valueSeparator As String) As String
    ReadOnly Property FileWarningByIndex(ByVal warningIndex As Integer) As udtMsgInfoType
    ReadOnly Property FileWarningList As udtMsgInfoType()

    ReadOnly Property LocalErrorCode As eValidateFastaFileErrorCodes

    ReadOnly Property ProteinCount As Integer
    ReadOnly Property FileLineCount As Integer
    ReadOnly Property ResidueCount As Long
    ReadOnly Property FASTAFilePath As String



    Property MaximumFileErrorsToTrack As Integer
    Property MinimumProteinNameLength As Integer
    Property MaximumProteinNameLength As Integer
    Property MaximumResiduesPerLine As Integer


    Property LongProteinNameSplitChars As String
    Property ProteinLineStartCharacter As Char
    Property ProteinNameInvalidCharactersToRemove As String

    Property ProteinNameFirstRefSepChars As String
    Property ProteinNameSubsequentRefSepChars As String

    Property ShowMessages As Boolean

    Structure udtMsgInfoType
        Public LineNumber As Integer
        Public ColNumber As Integer
        Public ProteinName As String
        Public MessageCode As Integer
        Public ExtraInfo As String
        Public Context As String
    End Structure

    Enum RuleTypes
        HeaderLine
        ProteinName
        ProteinDescription
        ProteinSequence
    End Enum

    Enum SwitchOptions
        AddMissingLinefeedatEOF
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
		FixedFastaKeepDuplicateNamedProteins		' Keep duplicate named proteins, unless the name and sequence match exactly, then they're removed
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

End Interface

