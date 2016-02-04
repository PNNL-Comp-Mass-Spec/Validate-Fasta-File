Option Strict On

Imports System.Runtime.CompilerServices
' This program will read in a Fasta file and write out stats on the number of proteins and number of residues
' It will also validate the protein name, descriptions, and sequences in the file

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

Module modMain

    Public Const PROGRAM_DATE As String = "February 3, 2016"

    Private mInputFilePath As String
    Private mOutputFolderPath As String
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

    Private mCreateModelXMLParameterFile As Boolean

    Private mRecurseFolders As Boolean
    Private mRecurseFoldersMaxLevels As Integer

    Private mQuietMode As Boolean

    Private WithEvents mValidateFastaFile As clsValidateFastaFile

    Private mLastProgressReportPctTime As System.DateTime
    Private mLastProgressReportTime As System.DateTime
    Private mLastProgressReportValue As Integer

    Public Function Main() As Integer
        ' Returns 0 if no error, error code if an error

        Dim intReturnCode As Integer
        Dim objParseCommandLine As New clsParseCommandLine
        Dim blnProceed As Boolean

        intReturnCode = 0
        mInputFilePath = String.Empty
        mOutputFolderPath = String.Empty
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

        mRecurseFolders = False
        mRecurseFoldersMaxLevels = 0

        mLastProgressReportPctTime = DateTime.UtcNow
        mLastProgressReportTime = DateTime.UtcNow

        ' Uncomment to preview a stack trace output
        ' TestStackTraceFormatter()
        ' TestNestedStringDictionary()

        Try
            blnProceed = False
            If objParseCommandLine.ParseCommandLine Then
                If SetOptionsUsingCommandLineParameters(objParseCommandLine) Then blnProceed = True
            End If

            If blnProceed And Not objParseCommandLine.NeedToShowHelp And mCreateModelXMLParameterFile Then
                If mParameterFilePath Is Nothing OrElse mParameterFilePath.Length = 0 Then
                    mParameterFilePath = System.IO.Path.GetFileNameWithoutExtension(System.Reflection.Assembly.GetExecutingAssembly().Location) & "_ModelSettings.xml"
                End If

                mValidateFastaFile = New clsValidateFastaFile
                mValidateFastaFile.SaveSettingsToParameterFile(mParameterFilePath)
                Console.WriteLine()
                Console.WriteLine("Created example XML parameter file: ")
                Console.WriteLine("  " & mParameterFilePath)
                Console.WriteLine()

            ElseIf Not blnProceed OrElse objParseCommandLine.NeedToShowHelp OrElse mInputFilePath.Length = 0 Then
                ShowProgramHelp()
                intReturnCode = -1
            Else

                mValidateFastaFile = New clsValidateFastaFile
                With mValidateFastaFile
                    .ShowMessages = Not mQuietMode

                    .SetOptionSwitch(IValidateFastaFile.SwitchOptions.OutputToStatsFile, mUseStatsFile)
                    .SetOptionSwitch(IValidateFastaFile.SwitchOptions.GenerateFixedFASTAFile, mGenerateFixedFastaFile)

                    ' Also use mGenerateFixedFastaFile to set SaveProteinSequenceHashInfoFiles
                    .SetOptionSwitch(IValidateFastaFile.SwitchOptions.SaveProteinSequenceHashInfoFiles, mGenerateFixedFastaFile)

                    .SetOptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaRenameDuplicateNameProteins, mFixedFastaRenameDuplicateNameProteins)
                    .SetOptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaKeepDuplicateNamedProteins, mFixedFastaKeepDuplicateNamedProteins)

                    .SetOptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaConsolidateDuplicateProteinSeqs, mFixedFastaConsolidateDuplicateProteinSeqs)
                    .SetOptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaConsolidateDupsIgnoreILDiff, mFixedFastaConsolidateDupsIgnoreILDiff)

                    .SetOptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaRemoveInvalidResidues, mFixedFastaRemoveInvalidResidues)

                    .SetOptionSwitch(IValidateFastaFile.SwitchOptions.AllowAsteriskInResidues, mAllowAsterisk)
                    .SetOptionSwitch(IValidateFastaFile.SwitchOptions.AllowDashInResidues, mAllowDash)

                    .SetOptionSwitch(IValidateFastaFile.SwitchOptions.SaveBasicProteinHashInfoFile, mSaveBasicProteinHashInfoFile)

                    .SetOptionSwitch(IValidateFastaFile.SwitchOptions.CheckForDuplicateProteinNames, mCheckForDuplicateProteinNames)
                    .SetOptionSwitch(IValidateFastaFile.SwitchOptions.CheckForDuplicateProteinSequences, mCheckForDuplicateProteinSequences)

                    ' Update the rules based on the options that were set above
                    .SetDefaultRules()
                End With

                ' Note: the following settings will be overridden if mParameterFilePath points to a valid parameter file that has these settings defined
                'With objValidateFastaFile
                '    .SetOptionSwitch(IValidateFastaFile.SwitchOptions.AddMissingLinefeedatEOF, )
                '    .SetOptionSwitch(IValidateFastaFile.SwitchOptions.AllowAsteriskInResidues, )
                '    .MaximumFileErrorsToTrack()
                '    .MinimumProteinNameLength()
                '    .MaximumProteinNameLength()
                '    .SetOptionSwitch(IValidateFastaFile.SwitchOptions.WarnBlankLinesBetweenProteins, )
                '    .SetOptionSwitch(IValidateFastaFile.SwitchOptions.CheckForDuplicateProteinSequences, )
                '    .SetOptionSwitch(IValidateFastaFile.SwitchOptions.SaveProteinSequenceHashInfoFiles, )
                'End With

                If mRecurseFolders Then
                    If mValidateFastaFile.ProcessFilesAndRecurseFolders(mInputFilePath, mOutputFolderPath, mOutputFolderPath, False, mParameterFilePath, mRecurseFoldersMaxLevels) Then
                        intReturnCode = 0
                    Else
                        intReturnCode = mValidateFastaFile.ErrorCode
                    End If
                Else
                    If mValidateFastaFile.ProcessFilesWildcard(mInputFilePath, mOutputFolderPath, mParameterFilePath) Then
                        intReturnCode = 0
                    Else
                        intReturnCode = mValidateFastaFile.ErrorCode
                        If intReturnCode <> 0 AndAlso Not mQuietMode Then
                            ShowErrorMessage("Error while processing: " & mValidateFastaFile.GetErrorMessage())
                        End If
                    End If
                End If

                DisplayProgressPercent(mLastProgressReportValue, True)
            End If

        Catch ex As Exception
            ShowErrorMessage("Error occurred in modMain->Main: " & System.Environment.NewLine & ex.Message)
            intReturnCode = -1
        End Try

        Return intReturnCode

    End Function

    Private Sub DisplayProgressPercent(ByVal intPercentComplete As Integer, ByVal blnAddCarriageReturn As Boolean)
        If blnAddCarriageReturn Then
            Console.WriteLine()
        End If
        If intPercentComplete > 100 Then intPercentComplete = 100
        Console.Write("Processing: " & intPercentComplete.ToString & "% ")
        If blnAddCarriageReturn Then
            Console.WriteLine()
        End If
    End Sub

    Private Function GetAppVersion() As String
        Return clsProcessFilesBaseClass.GetAppVersion(PROGRAM_DATE)
    End Function

    Private Function SetOptionsUsingCommandLineParameters(ByVal objParseCommandLine As clsParseCommandLine) As Boolean
        ' Returns True if no problems; otherwise, returns false

        Dim strValue As String = String.Empty
        Dim lstValidParameters = New List(Of String) From {
            "I", "O", "P", "C",
            "SkipDupeNameCheck", "SkipDupeSeqCheck",
            "F", "R", "D", "L", "V",
            "KeepSameName", "AllowDash", "AllowAsterisk", "B", "X", "S", "Q"}
        Dim intValue As Integer

        Try
            ' Make sure no invalid parameters are present
            If objParseCommandLine.InvalidParametersPresent(lstValidParameters) Then
                ShowErrorMessage("Invalid commmand line parameters",
                  (From item In objParseCommandLine.InvalidParameters(lstValidParameters) Select "/" + item).ToList())
                Return False
            Else
                With objParseCommandLine
                    ' Query objParseCommandLine to see if various parameters are present
                    If .RetrieveValueForParameter("I", strValue) Then
                        mInputFilePath = strValue
                    ElseIf .NonSwitchParameterCount > 0 Then
                        mInputFilePath = .RetrieveNonSwitchParameter(0)
                    End If

                    If .RetrieveValueForParameter("O", strValue) Then mOutputFolderPath = strValue
                    If .RetrieveValueForParameter("P", strValue) Then mParameterFilePath = strValue
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
                    If .IsParameterPresent("X") Then mCreateModelXMLParameterFile = True

                    If .RetrieveValueForParameter("S", strValue) Then
                        mRecurseFolders = True
                        If Integer.TryParse(strValue, intValue) Then
                            mRecurseFoldersMaxLevels = intValue
                        End If
                    End If

                    If .RetrieveValueForParameter("Q", strValue) Then mQuietMode = True
                End With

                Return True
            End If

        Catch ex As Exception
            ShowErrorMessage("Error parsing the command line parameters: " & System.Environment.NewLine & ex.Message)
        End Try
        Return False

    End Function

    Private Sub ShowErrorMessage(ByVal strMessage As String)
        Dim strSeparator As String = "------------------------------------------------------------------------------"

        Console.WriteLine()
        Console.WriteLine(strSeparator)
        Console.WriteLine(strMessage)
        Console.WriteLine(strSeparator)
        Console.WriteLine()

        WriteToErrorStream(strMessage)
    End Sub

    Private Sub ShowErrorMessage(ByVal strTitle As String, ByVal items As List(Of String))
        Dim strSeparator As String = "------------------------------------------------------------------------------"
        Dim strMessage As String

        Console.WriteLine()
        Console.WriteLine(strSeparator)
        Console.WriteLine(strTitle)
        strMessage = strTitle & ":"

        For Each item As String In items
            Console.WriteLine("   " + item)
            strMessage &= " " & item
        Next
        Console.WriteLine(strSeparator)
        Console.WriteLine()

        WriteToErrorStream(strMessage)
    End Sub

    Private Sub ShowProgramHelp()

        Try

            Console.WriteLine("This program will read a Fasta File and display statistics on the number of proteins and number of residues.  It will also check that the protein names, descriptions, and sequences are in the correct format.")
            Console.WriteLine()
            Console.WriteLine("Program syntax:" & System.Environment.NewLine & IO.Path.GetFileName(System.Reflection.Assembly.GetExecutingAssembly().Location))
            Console.WriteLine("  /I:InputFilePath.fasta [/O:OutputFolderPath]")
            Console.WriteLine(" [/P:ParameterFilePath] [/C] ")
            Console.WriteLine(" [/F] [/R] [/D] [/L] [/V] [/KeepSameName]")
            Console.WriteLine(" [/AllowDash] [/AllowAsterisk]")
            Console.WriteLine(" [/SkipDupeNameCheck] [/SkipDupeSeqCheck]")
            Console.WriteLine(" [/B] [/X] [/S:[MaxLevel]] [/Q]")
            Console.WriteLine()

            Console.WriteLine("The input file path can contain the wildcard character * and should point to a fasta file.")
            Console.WriteLine("The output folder path is optional, and is only used if /C is used.  If omitted, the output stats file will be created in the folder containing the .Exe file.")
            Console.WriteLine("Use /C to specify that an output file should be created, rather than displaying the results on the screen.")
            Console.WriteLine()
            Console.WriteLine("Use /F to shorten long protein names and generate a new, fixed .Fasta file.  At the same time, a file with protein names and hash values for each unique protein sequences will be generated (_UniqueProteinSeqs.txt).  This file will also list the other proteins that have duplicate sequences as the first protein mapped to each sequence.  If duplicate sequences are found, then an easily parseable mapping file will also be created (_UniqueProteinSeqDuplicates.txt).")
            Console.WriteLine("Use /R to rename proteins with duplicate names when using /F to generate a fixed fasta file.")
            Console.WriteLine("Use /D to consolidate proteins with duplicate protein sequences when using /F to generate a fixed fasta file.")
            Console.WriteLine("Use /L to ignore I/L (isoleucine vs. leucine) differences when consolidating proteins with duplicate protein sequences while generating a fixed fasta file.")
            Console.WriteLine("Use /V to remove invalid residues (non-letter characters, including an asterisk) when using /F to generate a fixed fasta file.")
            Console.WriteLine("Use /KeepSameName to keep proteins with the same name but differing sequences when using /F to generate a fixed fasta file (if they have the same name and same sequence, then will only retain one entry); ignored if /R or /D is used")
            Console.WriteLine("Use /AllowDash to allow a - in residues")
            Console.WriteLine("use /AllowAsterisk to allow * in residues")
            Console.WriteLine()
            Console.WriteLine("When parsing large fasta files, you can reduce the memory used by disabling the checking for duplicates")
            Console.WriteLine(" /SkipDupeSeqCheck disables duplicate sequence checking (large memory footprint)")
            Console.WriteLine(" /SkipDupeNameCheck disables duplicate name checking (small memory footprint)")
            Console.WriteLine()
            Console.WriteLine("Use /B to save a hash info file (even if not consolidating duplicates).  This is useful for parsing a large fasta file to obtain the sequence hash for each protein (hash values are not cached in memory, thus small memory footprint)")
            Console.WriteLine()
            Console.WriteLine("The parameter file path is optional.  If included, it should point to a valid XML parameter file.")
            Console.WriteLine("Use /X to specify that a model XML parameter file should be created.")
            Console.WriteLine("Use /S to process all valid files in the input folder and subfolders. Include a number after /S (like /S:2) to limit the level of subfolders to examine.")
            Console.WriteLine("The optional /Q switch will suppress all error messages.")
            Console.WriteLine()

            Console.WriteLine("Program written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA) in 2012")
            Console.WriteLine("Version: " & GetAppVersion())
            Console.WriteLine()

            Console.WriteLine("E-mail: matthew.monroe@pnnl.gov or matt@alchemistmatt.com")
            Console.WriteLine("Website: http://omics.pnl.gov/ or http://panomics.pnnl.gov/")
            Console.WriteLine()

            ' Delay for 750 msec in case the user double clicked this file from within Windows Explorer (or started the program via a shortcut)
            System.Threading.Thread.Sleep(750)

        Catch ex As Exception
            ShowErrorMessage("Error displaying the program syntax: " & ex.Message)
        End Try

    End Sub

    Private Sub TestNestedStringDictionary()

        Dim testData = New clsNestedStringDictionary(Of Integer)(True, 2)

        testData.Add("apple", 5)
        testData.Add("orange", 4)
        testData.Add("pear", 2)
        testData.Add("penne", 7)
        testData.Add("watermelon", 3)
        testData.Add("pizza", 1)
        testData.Add("pie", 4)
        testData.Add("squash", 1)
        testData.Add("avodado", 2)
        testData.Add("a", 2)
        testData.Add("pine nuts", 8)
        testData.Add("", 4)

        Console.WriteLine(testData.GetSizeSummary())

        Try
            ' This should throw an exception if we instantiated clsNestedStringDictionary using ignoreCaseForKeys=True
            testData.Add("Apple", 8)
        Catch ex As Exception
            If testData.IgnoreCase Then
                Console.WriteLine(ex.Message & " Apple cannot be added (this is to be expected)")
            Else
                Console.WriteLine(ex.Message & " Apple cannot be added (this is unexpected; it should have worked)")
            End If

        End Try
        
        Console.WriteLine()

        For Each item In testData.GetSpanningKeys()
            Dim itemDictionary = testData.GetDictionaryForSpanningKey(item)
            Console.WriteLine(item & ": " & itemDictionary.Count & " item(s), starting with " & itemDictionary.First().Key & ": " & itemDictionary.First().Value)
        Next

        Console.WriteLine()
    End Sub

    Private Sub TestStackTraceFormatter()

        Dim test = ""
        test &= "   at System.Collections.Generic.Dictionary`2.Resize(Int32 newSize, Boolean forceNewHashCodes)" & Environment.NewLine
        test &= "   at System.Collections.Generic.Dictionary`2.Insert(TKey key, TValue value, Boolean add)" & Environment.NewLine
        test &= "   at ValidateFastaFile.clsValidateFastaFile.ProcessSequenceHashInfo(String strProteinName, StringBuilder& sbCurrentResidues, Dictionary`2& lstProteinSequenceHashes, Int32& intProteinSequenceHashCount, udtProteinHashInfoType[]& udtProteinSeqHashInfo, Boolean blnConsolidateDupsIgnoreILDiff, StreamWriter& swProteinSequenceHashBasic) in F:\My Documents\Projects\DataMining\Validate_Fasta_File\clsValidateFastaFile.vb:line 3172" & Environment.NewLine
        test &= "   at ValidateFastaFile.clsValidateFastaFile.ProcessResiduesForPreviousProtein(String strProteinName, StringBuilder& sbCurrentResidues, Dictionary`2& lstProteinSequenceHashes, Int32& intProteinSequenceHashCount, udtProteinHashInfoType[]& udtProteinSeqHashInfo, Boolean blnConsolidateDupsIgnoreILDiff, StreamWriter& swFixedFastaOut, Int32 intCurrentValidResidueLineLengthMax, StreamWriter& swProteinSequenceHashBasic) in F:\My Documents\Projects\DataMining\Validate_Fasta_File\clsValidateFastaFile.vb:line 3070" & Environment.NewLine
        test &= "   at ValidateFastaFile.clsValidateFastaFile.AnalyzeFastaFile(String strFastaFilePath) in F:\My Documents\Projects\DataMining\Validate_Fasta_File\clsValidateFastaFile.vb:line 986"

        Dim stackTraceData As IEnumerable(Of String) = clsStackTraceFormatter.GetExceptionStackTraceData(test)

        Dim sbStackTrace = New Text.StringBuilder()
        sbStackTrace.AppendLine(clsStackTraceFormatter.STACK_TRACE_TITLE)

        For index = 0 To stackTraceData.Count - 1
            sbStackTrace.AppendLine("  " & stackTraceData(index))
        Next

        Console.WriteLine(sbStackTrace.ToString())
        Console.WriteLine()

        TestThrowException()

    End Sub

    Public Sub TestThrowException()
        TestThrowException(0)
    End Sub

    Private Sub TestThrowException(recursionLevel As Integer)
        If recursionLevel > 3 Then
            Dim testData = New Dictionary(Of String, String)
            Try
                testData.Add("a", "apple")
                testData.Add("d", "donut")
                testData.Add("e", "egg")
                testData.Add("a", "avodcado")
            Catch ex As Exception
                Console.WriteLine(clsStackTraceFormatter.GetExceptionStackTrace(ex))
                Console.WriteLine()
                Console.WriteLine(clsStackTraceFormatter.GetExceptionStackTraceMultiLine(ex))
            End Try
        Else
            TestThrowException(recursionLevel + 1)
        End If

    End Sub

    Private Sub WriteToErrorStream(strErrorMessage As String)
        Try
            Using swErrorStream As System.IO.StreamWriter = New System.IO.StreamWriter(Console.OpenStandardError())
                swErrorStream.WriteLine(strErrorMessage)
            End Using
        Catch ex As Exception
            ' Ignore errors here
        End Try
    End Sub

    Private Sub mValidateFastaFile_ProgressChanged(ByVal taskDescription As String, ByVal percentComplete As Single) Handles mValidateFastaFile.ProgressChanged
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
