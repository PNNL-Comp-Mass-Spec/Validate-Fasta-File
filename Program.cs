Option Strict On

Imports System.Reflection
Imports System.Threading
Imports PRISM

' This program will read in a Fasta file and write out stats on the number of proteins and number of residues
' It will also validate the protein name, descriptions, and sequences in the file

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

Module Program

    Public Const PROGRAM_DATE As String = "April 15, 2020"

    Private mInputFilePath As String
    Private mOutputDirectoryPath As String
    Private mParameterFilePath As String

    Private mUseStatsFile As Boolean
    Private mGenerateFixedFastaFile As Boolean

    Private mCheckForDuplicateProteinNames As Boolean
    Private mCheckForDuplicateProteinSequences As Boolean

    Private mFixedFastaRenameDuplicateNameProteins As Boolean
    Private mFixedFastaKeepDuplicateNamedProteins As Boolean

    Private mFixedFastaConsolidateDuplicateProteinSeqs As Boolean
    Private mFixedFastaConsolidateDupsIgnoreILDiff As Boolean
    Private mFixedFastaRemoveInvalidResidues As Boolean

    Private mAllowAsterisk As Boolean
    Private mAllowDash As Boolean

    Private mSaveBasicProteinHashInfoFile As Boolean
    Private mProteinHashFilePath As String

    Private mCreateModelXMLParameterFile As Boolean

    Private mRecurseDirectories As Boolean
    Private mMaxLevelsToRecurse As Integer

    Private WithEvents mValidateFastaFile As clsValidateFastaFile

    Private mLastProgressReportPctTime As DateTime
    Private mLastProgressReportTime As DateTime
    Private mLastProgressReportValue As Integer

    Public Function Main() As Integer
        ' Returns 0 if no error, error code if an error

        Dim returnCode As Integer
        Dim commandLineParser As New clsParseCommandLine
        Dim proceed As Boolean

        returnCode = 0
        mInputFilePath = String.Empty
        mOutputDirectoryPath = String.Empty
        mParameterFilePath = String.Empty

        mUseStatsFile = False
        mGenerateFixedFastaFile = False

        mCheckForDuplicateProteinNames = True
        mCheckForDuplicateProteinSequences = True

        mFixedFastaRenameDuplicateNameProteins = False
        mFixedFastaKeepDuplicateNamedProteins = False

        mFixedFastaConsolidateDuplicateProteinSeqs = False
        mFixedFastaConsolidateDupsIgnoreILDiff = False

        mAllowAsterisk = False
        mAllowDash = False

        mSaveBasicProteinHashInfoFile = False
        mProteinHashFilePath = String.Empty

        mRecurseDirectories = False
        mMaxLevelsToRecurse = 0

        mLastProgressReportPctTime = DateTime.UtcNow
        mLastProgressReportTime = DateTime.UtcNow

        Try
            proceed = False
            If commandLineParser.ParseCommandLine Then
                If SetOptionsUsingCommandLineParameters(commandLineParser) Then proceed = True
            End If

            If proceed And Not commandLineParser.NeedToShowHelp And mCreateModelXMLParameterFile Then
                If mParameterFilePath Is Nothing OrElse mParameterFilePath.Length = 0 Then
                    mParameterFilePath = Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().Location) & "_ModelSettings.xml"
                End If

                mValidateFastaFile = New clsValidateFastaFile
                mValidateFastaFile.SaveSettingsToParameterFile(mParameterFilePath)
                Console.WriteLine()
                Console.WriteLine("Created example XML parameter file: ")
                Console.WriteLine("  " & mParameterFilePath)
                Console.WriteLine()

            ElseIf Not proceed OrElse commandLineParser.NeedToShowHelp OrElse mInputFilePath.Length = 0 Then
                ShowProgramHelp()
                returnCode = -1
            Else

                mValidateFastaFile = New clsValidateFastaFile
                With mValidateFastaFile

                    .SetOptionSwitch(clsValidateFastaFile.SwitchOptions.OutputToStatsFile, mUseStatsFile)
                    .SetOptionSwitch(clsValidateFastaFile.SwitchOptions.GenerateFixedFASTAFile, mGenerateFixedFastaFile)

                    ' Also use mGenerateFixedFastaFile to set SaveProteinSequenceHashInfoFiles
                    .SetOptionSwitch(clsValidateFastaFile.SwitchOptions.SaveProteinSequenceHashInfoFiles, mGenerateFixedFastaFile)

                    .SetOptionSwitch(clsValidateFastaFile.SwitchOptions.FixedFastaRenameDuplicateNameProteins, mFixedFastaRenameDuplicateNameProteins)
                    .SetOptionSwitch(clsValidateFastaFile.SwitchOptions.FixedFastaKeepDuplicateNamedProteins, mFixedFastaKeepDuplicateNamedProteins)

                    .SetOptionSwitch(clsValidateFastaFile.SwitchOptions.FixedFastaConsolidateDuplicateProteinSeqs, mFixedFastaConsolidateDuplicateProteinSeqs)
                    .SetOptionSwitch(clsValidateFastaFile.SwitchOptions.FixedFastaConsolidateDupsIgnoreILDiff, mFixedFastaConsolidateDupsIgnoreILDiff)

                    .SetOptionSwitch(clsValidateFastaFile.SwitchOptions.FixedFastaRemoveInvalidResidues, mFixedFastaRemoveInvalidResidues)

                    .SetOptionSwitch(clsValidateFastaFile.SwitchOptions.AllowAsteriskInResidues, mAllowAsterisk)
                    .SetOptionSwitch(clsValidateFastaFile.SwitchOptions.AllowDashInResidues, mAllowDash)

                    .SetOptionSwitch(clsValidateFastaFile.SwitchOptions.SaveBasicProteinHashInfoFile, mSaveBasicProteinHashInfoFile)

                    .SetOptionSwitch(clsValidateFastaFile.SwitchOptions.CheckForDuplicateProteinNames, mCheckForDuplicateProteinNames)
                    .SetOptionSwitch(clsValidateFastaFile.SwitchOptions.CheckForDuplicateProteinSequences, mCheckForDuplicateProteinSequences)

                    ' Update the rules based on the options that were set above
                    .SetDefaultRules()

                    mValidateFastaFile.ExistingProteinHashFile = mProteinHashFilePath

                End With

                mValidateFastaFile.SkipConsoleWriteIfNoProgressListener = True

                ' Note: the following settings will be overridden if mParameterFilePath points to a valid parameter file that has these settings defined
                'With mValidateFastaFile
                '    .SetOptionSwitch(clsValidateFastaFile.SwitchOptions.AddMissingLineFeedAtEOF, )
                '    .SetOptionSwitch(clsValidateFastaFile.SwitchOptions.AllowAsteriskInResidues, )
                '    .MaximumFileErrorsToTrack()
                '    .MinimumProteinNameLength()
                '    .MaximumProteinNameLength()
                '    .SetOptionSwitch(clsValidateFastaFile.SwitchOptions.WarnBlankLinesBetweenProteins, )
                '    .SetOptionSwitch(clsValidateFastaFile.SwitchOptions.CheckForDuplicateProteinSequences, )
                '    .SetOptionSwitch(clsValidateFastaFile.SwitchOptions.SaveProteinSequenceHashInfoFiles, )
                'End With

                If mRecurseDirectories Then
                    If mValidateFastaFile.ProcessFilesAndRecurseDirectories(mInputFilePath, mOutputDirectoryPath, mOutputDirectoryPath, False, mParameterFilePath, mMaxLevelsToRecurse) Then
                        returnCode = 0
                    Else
                        returnCode = mValidateFastaFile.ErrorCode
                    End If
                Else
                    If mValidateFastaFile.ProcessFilesWildcard(mInputFilePath, mOutputDirectoryPath, mParameterFilePath) Then
                        returnCode = 0
                    Else
                        returnCode = mValidateFastaFile.ErrorCode
                        If returnCode <> 0 Then
                            ShowErrorMessage("Error while processing: " & mValidateFastaFile.GetErrorMessage())
                        End If
                    End If
                End If

                DisplayProgressPercent(mLastProgressReportValue, True)
            End If

        Catch ex As Exception
            ShowErrorMessage("Error occurred in Program->Main: " & ex.Message, ex)
            returnCode = -1
        End Try

        Return returnCode

    End Function

    Private Sub DisplayProgressPercent(percentComplete As Integer, addCarriageReturn As Boolean)
        If addCarriageReturn Then
            Console.WriteLine()
        End If
        If percentComplete > 100 Then percentComplete = 100
        Console.Write("Processing: " & percentComplete.ToString & "% ")
        If addCarriageReturn Then
            Console.WriteLine()
        End If
    End Sub

    Private Function GetAppVersion() As String
        Return FileProcessor.ProcessFilesOrDirectoriesBase.GetAppVersion(PROGRAM_DATE)
    End Function

    Private Function SetOptionsUsingCommandLineParameters(commandLineParser As clsParseCommandLine) As Boolean
        ' Returns True if no problems; otherwise, returns false

        Dim value As String = String.Empty
        Dim validParameters = New List(Of String) From {
            "I", "O", "P", "C",
            "SkipDupeNameCheck", "SkipDupeSeqCheck",
            "F", "R", "D", "L", "V",
            "KeepSameName", "AllowDash", "AllowAsterisk",
            "B", "HashFile",
            "X", "S"}
        Dim valueInteger As Integer

        Try
            ' Make sure no invalid parameters are present
            If commandLineParser.InvalidParametersPresent(validParameters) Then
                ShowErrorMessage("Invalid command line parameters",
                  (From item In commandLineParser.InvalidParameters(validParameters) Select "/" + item).ToList())
                Return False
            Else
                With commandLineParser
                    ' Query commandLineParser to see if various parameters are present
                    If .RetrieveValueForParameter("I", value) Then
                        mInputFilePath = value
                    ElseIf .NonSwitchParameterCount > 0 Then
                        mInputFilePath = .RetrieveNonSwitchParameter(0)
                    End If

                    If .RetrieveValueForParameter("O", value) Then mOutputDirectoryPath = value
                    If .RetrieveValueForParameter("P", value) Then mParameterFilePath = value
                    If .IsParameterPresent("C") Then mUseStatsFile = True

                    If .IsParameterPresent("SkipDupeNameCheck") Then mCheckForDuplicateProteinNames = False
                    If .IsParameterPresent("SkipDupeSeqCheck") Then mCheckForDuplicateProteinSequences = False

                    If .IsParameterPresent("F") Then mGenerateFixedFastaFile = True
                    If .IsParameterPresent("R") Then mFixedFastaRenameDuplicateNameProteins = True
                    If .IsParameterPresent("D") Then mFixedFastaConsolidateDuplicateProteinSeqs = True
                    If .IsParameterPresent("L") Then mFixedFastaConsolidateDupsIgnoreILDiff = True
                    If .IsParameterPresent("V") Then mFixedFastaRemoveInvalidResidues = True
                    If .IsParameterPresent("KeepSameName") Then mFixedFastaKeepDuplicateNamedProteins = True

                    If .IsParameterPresent("AllowAsterisk") Then mAllowAsterisk = True
                    If .IsParameterPresent("AllowDash") Then mAllowDash = True

                    If .IsParameterPresent("B") Then mSaveBasicProteinHashInfoFile = True
                    If .RetrieveValueForParameter("HashFile", value) Then mProteinHashFilePath = value

                    If .IsParameterPresent("X") Then mCreateModelXMLParameterFile = True

                    If .RetrieveValueForParameter("S", value) Then
                        mRecurseDirectories = True
                        If Integer.TryParse(value, valueInteger) Then
                            mMaxLevelsToRecurse = valueInteger
                        End If
                    End If

                End With

                Return True
            End If

        Catch ex As Exception
            ShowErrorMessage("Error parsing the command line parameters: " & ex.Message, ex)
        End Try
        Return False

    End Function

    Private Sub ShowErrorMessage(message As String, Optional ex As Exception = Nothing)
        ConsoleMsgUtils.ShowError(message, ex)
    End Sub

    Private Sub ShowErrorMessage(title As String, items As List(Of String))
        ConsoleMsgUtils.ShowErrors(title, items)
    End Sub

    Private Sub ShowProgramHelp()

        Try
            Dim exeName = Path.GetFileName(Assembly.GetExecutingAssembly().Location)

            Console.WriteLine("== Overview ==")
            Console.WriteLine()
            Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                "This program will read a Fasta File and display statistics on the number of proteins and number of residues. " &
                "It will also check that the protein names, descriptions, and sequences are in the correct format."))
            Console.WriteLine()
            Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                "The program can optionally create a new, fixed version of a fasta file where proteins with duplicate sequences " &
                "have been consolidated, and proteins with duplicate names have been renamed."))
            Console.WriteLine()
            Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                "To remove duplicates from huge fasta files (over 1 GB in size), " &
                "first create the ProteinHashes.txt file by calling this program with:"))
            Console.WriteLine("  {0} Proteins.fasta /B /SkipDupeSeqCheck /SkipDupeNameCheck", exeName)
            Console.WriteLine()
            Console.WriteLine("Next call the program again, providing the name of the ProteinHashes file:")
            Console.WriteLine("  {0} Proteins.fasta /HashFile:Proteins_ProteinHashes.txt", exeName)
            Console.WriteLine()
            Console.WriteLine("== Program syntax ==")
            Console.WriteLine()
            Console.WriteLine(exeName)
            Console.WriteLine(" /I:InputFilePath.fasta [/O:OutputDirectoryPath]")
            Console.WriteLine(" [/P:ParameterFilePath] [/C] ")
            Console.WriteLine(" [/F] [/R] [/D] [/L] [/V] [/KeepSameName]")
            Console.WriteLine(" [/AllowDash] [/AllowAsterisk]")
            Console.WriteLine(" [/SkipDupeNameCheck] [/SkipDupeSeqCheck]")
            Console.WriteLine(" [/B] [/HashFile]")
            Console.WriteLine(" [/X] [/S:[MaxLevel]]")
            Console.WriteLine()
            Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                "The input file path can contain the wildcard character * and should point to a fasta file."))
            Console.WriteLine()
            Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                "The output directory path is optional, and is only used if /C is used. If omitted, the output stats file " &
                "will be created in the directory containing the .Exe file."))
            Console.WriteLine()
            Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                "The parameter file path is optional. If included, it should point to a valid XML parameter file."))

            Console.WriteLine()
            Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                "Use /C to specify that an output file should be created, rather than displaying the results on the screen."))

            Console.WriteLine()
            Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                "Use /F to generate a new, fixed .Fasta file (long protein names will be auto-shortened). " &
                "At the same time, a file with protein names and hash values for each unique protein sequences " &
                "will be generated (_UniqueProteinSeqs.txt). This file will also list the other proteins " &
                "that have duplicate sequences as the first protein mapped to each sequence. If duplicate sequences " &
                "are found, then an easily parseable mapping file will also be created (_UniqueProteinSeqDuplicates.txt)."))
            Console.WriteLine()
            Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                "Use /R to rename proteins with duplicate names when using /F to generate a fixed fasta file."))
            Console.WriteLine()
            Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                "Use /D to consolidate proteins with duplicate protein sequences when using /F to generate a fixed fasta file."))
            Console.WriteLine()
            Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                "Use /L to ignore I/L (isoleucine vs. leucine) differences when consolidating proteins " &
                "with duplicate protein sequences while generating a fixed fasta file."))
            Console.WriteLine()
            Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                "Use /V to remove invalid residues (non-letter characters, including an asterisk) when using /F to generate a fixed fasta file."))
            Console.WriteLine()
            Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                "Use /KeepSameName to keep proteins with the same name but differing sequences when using /F to generate a fixed fasta file " &
                "(if they have the same name and same sequence, then will only retain one entry); ignored if /R or /D is used"))
            Console.WriteLine()
            Console.WriteLine("Use /AllowDash to allow a - in residues")
            Console.WriteLine("Use /AllowAsterisk to allow * in residues")
            Console.WriteLine()
            Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                "When parsing large fasta files, you can reduce the memory used by disabling the checking for duplicates"))

            Console.WriteLine(" /SkipDupeSeqCheck disables duplicate sequence checking (large memory footprint)")
            Console.WriteLine(" /SkipDupeNameCheck disables duplicate name checking (small memory footprint)")
            Console.WriteLine()
            Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                "Use /B to save a hash info file (even if not consolidating duplicates). " &
                "This is useful for parsing a large fasta file to obtain the sequence hash for each protein " &
                "(hash values are not cached in memory, thus small memory footprint)."))
            Console.WriteLine()
            Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                "Use /HashFile to specify a pre-computed hash file to use for determining which proteins to keep when generating a fixed fasta file"))
            Console.WriteLine(ConsoleMsgUtils.WrapParagraph("Use of /HashFile automatically enables /F and automatically disables /D, /R, and /B"))
            Console.WriteLine()
            Console.WriteLine("Use /X to specify that a model XML parameter file should be created.")
            Console.WriteLine(ConsoleMsgUtils.WrapParagraph(
                "Use /S to process all valid files in the input directory and subdirectories. " &
                "Include a number after /S (like /S:2) to limit the level of subdirectories to examine."))
            Console.WriteLine()

            Console.WriteLine("Program written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA) in 2012")
            Console.WriteLine("Version: " & GetAppVersion())
            Console.WriteLine()

            Console.WriteLine("E-mail: matthew.monroe@pnnl.gov or proteomics@pnnl.gov")
            Console.WriteLine("Website: https://omics.pnl.gov/ or https://panomics.pnnl.gov/")
            Console.WriteLine()

            ' Delay for 750 msec in case the user double clicked this file from within Windows Explorer (or started the program via a shortcut)
            Thread.Sleep(750)

        Catch ex As Exception
            ShowErrorMessage("Error displaying the program syntax: " & ex.Message, ex)
        End Try

    End Sub

    Private Sub mValidateFastaFile_ProgressChanged(taskDescription As String, percentComplete As Single) Handles mValidateFastaFile.ProgressUpdate
        Const PERCENT_REPORT_INTERVAL = 25
        Const PROGRESS_DOT_INTERVAL_MSEC = 500

        If percentComplete >= mLastProgressReportValue OrElse
           DateTime.UtcNow.Subtract(mLastProgressReportPctTime).TotalSeconds >= 30 Then

            mLastProgressReportPctTime = DateTime.UtcNow

            If mLastProgressReportValue > 0 Then
                Console.WriteLine()
            End If

            If percentComplete < mLastProgressReportValue Then
                DisplayProgressPercent(CInt(Math.Round(percentComplete, 0)), False)
            Else
                DisplayProgressPercent(mLastProgressReportValue, False)
            End If

            While percentComplete >= mLastProgressReportValue
                mLastProgressReportValue += PERCENT_REPORT_INTERVAL
            End While

            mLastProgressReportTime = DateTime.UtcNow
        Else
            If DateTime.UtcNow.Subtract(mLastProgressReportTime).TotalMilliseconds >= PROGRESS_DOT_INTERVAL_MSEC Then
                mLastProgressReportTime = DateTime.UtcNow
                Console.Write(".")
            End If
        End If
    End Sub

    Private Sub mValidateFastaFile_ProgressReset() Handles mValidateFastaFile.ProgressReset
        mLastProgressReportTime = DateTime.UtcNow
        mLastProgressReportPctTime = DateTime.UtcNow
        mLastProgressReportValue = 0
    End Sub
End Module
