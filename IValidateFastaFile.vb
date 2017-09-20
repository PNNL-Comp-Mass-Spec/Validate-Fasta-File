Public Interface IValidateFastaFile

    Event ProgressChanged(
        taskDescription As String,
        percentComplete As Single)
    Event ProgressCompleted()
    Event WroteLineEndNormalizedFASTA(newFilePath As String)

    Function ValidateFASTAFile(filePath As String) As Boolean

    Function ValidateFASTAFile(
        filePath As String,
        errorDumpOutputPath As String,
        parameterFilePath As String,
        resetErrorCodes As Boolean) As Boolean

    Function GetDefaultExtensionsToParse() As String()
    Function GetCurrentErrorMessage() As String

    Function SaveParameterSettingsToParameterFile(parameterFilePath As String) As Boolean
    Function LoadParameterFileSettings(parameterFilePath As String) As Boolean

    Function LookupMessageDescription(errorMessageCode As Integer) As String
    Function LookupMessageDescription(errorMessageCode As Integer, ExtraInfo As String) As String
    Function LookupMessageTypeString(entryType As eMsgTypeConstants) As String


    Sub AbortProcessingNow()

    Sub ClearAllRules()
    Sub ClearRules(ruleType As RuleTypes)

    Sub SetDefaultRules()

    Sub SetRule(
        ruleType As RuleTypes,
        regexToMatch As String,
        doesMatchIndicateProblem As Boolean,
        problemReturnMessage As String,
        severityLevel As Short)
    Sub SetRule(
        ruleType As RuleTypes,
        regexToMatch As String,
        doesMatchIndicateProblem As Boolean,
        problemReturnMessage As String,
        severityLevel As Short,
        displayMatchAsExtraInfo As Boolean)

    Property OptionSwitches(SwitchName As SwitchOptions) As Boolean

    ReadOnly Property FixedFASTAFileStats(
        ValueType As FixedFASTAFileValues) As Integer


    ReadOnly Property ErrorMessageTextByIndex(
        errorIndex As Integer,
        valueSeparator As String) As String
    ReadOnly Property FileErrorByIndex(errorIndex As Integer) As udtMsgInfoType
    ReadOnly Property FileErrorList As List(Of udtMsgInfoType)

    ReadOnly Property WarningMessageTextByIndex(
        warningIndex As Integer,
        valueSeparator As String) As String
    ReadOnly Property FileWarningByIndex(warningIndex As Integer) As udtMsgInfoType
    ReadOnly Property FileWarningList As List(Of udtMsgInfoType)

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

End Interface

