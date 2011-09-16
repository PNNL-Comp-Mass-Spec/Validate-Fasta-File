Option Strict On

' This program will read in a Fasta file and write out stats on the number of proteins and number of residues
' It will also validate the protein name, descriptions, and sequences in the file

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

Module modMain

	Public Const PROGRAM_DATE As String = "September 16, 2011"

    Private mInputFilePath As String
    Private mOutputFolderPath As String
    Private mParameterFilePath As String

    Private mUseStatsFile As Boolean
    Private mGenerateFixedFastaFile As Boolean
    Private mFixedFastaRenameDuplicateNameProteins As Boolean
    Private mFixedFastaConsolidateDuplicateProteinSeqs As Boolean
    Private mFixedFastaConsolidateDupsIgnoreILDiff As Boolean
    Private mSaveBasicProteinHashInfoFile As Boolean

    Private mCreateModelXMLParameterFile As Boolean

    Private mRecurseFolders As Boolean
    Private mRecurseFoldersMaxLevels As Integer

    Private mQuietMode As Boolean

    Private WithEvents mValidateFastaFile As clsValidateFastaFile
    Private mLastProgressReportTime As System.DateTime
    Private mLastProgressReportValue As Integer

    Public Function Main() As Integer
        ' Returns 0 if no error, error code if an error

        Dim intReturnCode As Integer
        Dim objParseCommandLine As New clsParseCommandLine
        Dim blnProceed As Boolean

        Dim ProteinCount As Integer = 0
        Dim ResidueCount As Long = 0

        intReturnCode = 0
        mInputFilePath = String.Empty
        mOutputFolderPath = String.Empty
        mParameterFilePath = String.Empty

        mUseStatsFile = False
        mGenerateFixedFastaFile = False

        mFixedFastaRenameDuplicateNameProteins = False
        mFixedFastaConsolidateDuplicateProteinSeqs = False
        mFixedFastaConsolidateDupsIgnoreILDiff = False
        mSaveBasicProteinHashInfoFile = False

        mRecurseFolders = False
        mRecurseFoldersMaxLevels = 0

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
                    .SetOptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaConsolidateDuplicateProteinSeqs, mFixedFastaConsolidateDuplicateProteinSeqs)
                    .SetOptionSwitch(IValidateFastaFile.SwitchOptions.FixedFastaConsolidateDupsIgnoreILDiff, mFixedFastaConsolidateDupsIgnoreILDiff)
                    .SetOptionSwitch(IValidateFastaFile.SwitchOptions.SaveBasicProteinHashInfoFile, mSaveBasicProteinHashInfoFile)
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

                If mOutputFolderPath Is Nothing OrElse mOutputFolderPath.Length = 0 Then
                    ' Define the output folder path as the path containing the .exe
                    mOutputFolderPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
                End If

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
                            MsgBox("Error while processing: " & mValidateFastaFile.GetErrorMessage(), MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
                        End If
                    End If
                End If


            End If

        Catch ex As Exception
            Console.WriteLine("Error occurred in modMain->Main: " & ControlChars.NewLine & ex.Message)
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
        'Return System.Windows.Forms.Application.ProductVersion & " (" & PROGRAM_DATE & ")"

        Return System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.ToString & " (" & PROGRAM_DATE & ")"
    End Function

    Private Function SetOptionsUsingCommandLineParameters(ByVal objParseCommandLine As clsParseCommandLine) As Boolean
        ' Returns True if no problems; otherwise, returns false

        Dim strValue As String = String.Empty
        Dim strValidParameters() As String = New String() {"I", "O", "P", "C", "F", "R", "D", "L", "B", "X", "S", "Q"}

        Try
            ' Make sure no invalid parameters are present
            If objParseCommandLine.InvalidParametersPresent(strValidParameters) Then
                Return False
            Else
                With objParseCommandLine
                    ' Query objParseCommandLine to see if various parameters are present
                    If .RetrieveValueForParameter("I", strValue) Then
                        mInputFilePath = strValue
                    Else
                        ' User didn't use /I:InputFile
                        ' See if they simply provided the file name
                        If .NonSwitchParameterCount > 0 Then
                            mInputFilePath = .RetrieveNonSwitchParameter(0)
                        End If
                    End If

                    If .RetrieveValueForParameter("O", strValue) Then mOutputFolderPath = strValue
                    If .RetrieveValueForParameter("P", strValue) Then mParameterFilePath = strValue
                    If .RetrieveValueForParameter("C", strValue) Then mUseStatsFile = True
                    If .RetrieveValueForParameter("F", strValue) Then mGenerateFixedFastaFile = True
                    If .RetrieveValueForParameter("R", strValue) Then mFixedFastaRenameDuplicateNameProteins = True
                    If .RetrieveValueForParameter("D", strValue) Then mFixedFastaConsolidateDuplicateProteinSeqs = True
                    If .RetrieveValueForParameter("L", strValue) Then mFixedFastaConsolidateDupsIgnoreILDiff = True
                    If .RetrieveValueForParameter("B", strValue) Then mSaveBasicProteinHashInfoFile = True
                    If .RetrieveValueForParameter("X", strValue) Then mCreateModelXMLParameterFile = True

                    If .RetrieveValueForParameter("S", strValue) Then
                        mRecurseFolders = True
                        If IsNumeric(strValue) Then
                            mRecurseFoldersMaxLevels = CInt(strValue)
                        End If
                    End If

                    If .RetrieveValueForParameter("Q", strValue) Then mQuietMode = True
                End With

                Return True
            End If

        Catch ex As Exception
             Console.WriteLine("Error parsing the command line parameters: " & ControlChars.NewLine & ex.Message)
        End Try

    End Function

    Private Sub ShowProgramHelp()

        Try

            Console.WriteLine("This program will read a Fasta File and display statistics on the number of proteins and number of residues.  It will also check that the protein names, descriptions, and sequences are in the correct format.")
            Console.WriteLine()
            Console.WriteLine("Program syntax:" & ControlChars.NewLine & IO.Path.GetFileName(System.Reflection.Assembly.GetExecutingAssembly().Location))
            Console.WriteLine("  /I:InputFilePath.fasta [/O:OutputFolderPath]")
            Console.WriteLine(" [/P:ParameterFilePath] [/C] ")
            Console.WriteLine(" [/F] [/R] [/D] [/L] [/B]")
            Console.WriteLine(" [/X] [/S:[MaxLevel]] [/Q]")
            Console.WriteLine()

            Console.WriteLine("The input file path can contain the wildcard character * and should point to a fasta file.")
            Console.WriteLine("The output folder path is optional, and is only used if /C is used.  If omitted, the output stats file will be created in the folder containing the .Exe file.")
            Console.WriteLine("Use /C to specify that an output file should be created, rather than displaying the results on the screen.")
            Console.WriteLine()
            Console.WriteLine("Use /F to shorten long protein names and remove invalid characters from the residues line, generating a new, fixed .Fasta file.  At the same time, a file with protein names and hash values for each unique protein sequences will be generated (_UniqueProteinSeqs.txt).  This file will also list the other proteins that have duplicate sequences as the first protein mapped to each sequence.  If duplicate sequences are found, then an easily parseable mapping file will also be created (_UniqueProteinSeqDuplicates.txt).")
            Console.WriteLine("Use /R to rename proteins with duplicate names when using /F to generate a fixed fasta file.")
            Console.WriteLine("Use /D to consolidate proteins with duplicate protein sequences when using /F to generate a fixed fasta file.")
            Console.WriteLine("Use /L to ignore I/L (isoleucine vs. leucine) differences when consolidating proteins with duplicate protein sequences while generating a fixed fasta file.")
            Console.WriteLine()

            Console.WriteLine("The parameter file path is optional.  If included, it should point to a valid XML parameter file.")
            Console.WriteLine("Use /X to specify that a model XML parameter file should be created.")
            Console.WriteLine("Use /S to process all valid files in the input folder and subfolders. Include a number after /S (like /S:2) to limit the level of subfolders to examine.")
            Console.WriteLine("The optional /Q switch will suppress all error messages.")
            Console.WriteLine()

            Console.WriteLine("Program written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA) in 2005")
            Console.WriteLine("Version: " & GetAppVersion())
            Console.WriteLine()

            Console.WriteLine("E-mail: matthew.monroe@pnl.gov or matt@alchemistmatt.com")
            Console.WriteLine("Website: http://ncrr.pnl.gov/ or http://www.sysbio.org/resources/staff/")
            Console.WriteLine()

            Console.WriteLine("Licensed under the Apache License, Version 2.0; you may not use this file except in compliance with the License.  " & _
                              "You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0")
            Console.WriteLine()

            Console.WriteLine("Notice: This computer software was prepared by Battelle Memorial Institute, " & _
                              "hereinafter the Contractor, under Contract No. DE-AC05-76RL0 1830 with the " & _
                              "Department of Energy (DOE).  All rights in the computer software are reserved " & _
                              "by DOE on behalf of the United States Government and the Contractor as " & _
                              "provided in the Contract.  NEITHER THE GOVERNMENT NOR THE CONTRACTOR MAKES ANY " & _
                              "WARRANTY, EXPRESS OR IMPLIED, OR ASSUMES ANY LIABILITY FOR THE USE OF THIS " & _
                              "SOFTWARE.  This notice including this sentence must appear on any copies of " & _
                              "this computer software.")

            ' Delay for 750 msec in case the user double clicked this file from within Windows Explorer (or started the program via a shortcut)
            System.Threading.Thread.Sleep(750)

        Catch ex As Exception
            Console.WriteLine("Error displaying the program syntax: " & ex.Message)
        End Try

    End Sub

    Private Sub mValidateFastaFile_ProgressChanged(ByVal taskDescription As String, ByVal percentComplete As Single) Handles mValidateFastaFile.ProgressChanged
        Const PERCENT_REPORT_INTERVAL As Integer = 25
        Const PROGRESS_DOT_INTERVAL_MSEC As Integer = 250

        If percentComplete >= mLastProgressReportValue Then
            If mLastProgressReportValue > 0 Then
                Console.WriteLine()
            End If
            DisplayProgressPercent(mLastProgressReportValue, False)
            mLastProgressReportValue += PERCENT_REPORT_INTERVAL
            mLastProgressReportTime = DateTime.UtcNow
        Else
            If DateTime.UtcNow.Subtract(mLastProgressReportTime).TotalMilliseconds > PROGRESS_DOT_INTERVAL_MSEC Then
                mLastProgressReportTime = DateTime.UtcNow
                Console.Write(".")
            End If
        End If
    End Sub

    Private Sub mValidateFastaFile_ProgressReset() Handles mValidateFastaFile.ProgressReset
        mLastProgressReportTime = DateTime.UtcNow
        mLastProgressReportValue = 0
    End Sub
End Module
