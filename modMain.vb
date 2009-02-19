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

    Public Const PROGRAM_DATE As String = "July 28, 2008"

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

    Public Function Main() As Integer
        ' Returns 0 if no error, error code if an error

        Dim objValidateFastaFile As clsValidateFastaFile

        Dim intReturnCode As Integer
        Dim objParseCommandLine As New clsParseCommandLine
        Dim blnProceed As Boolean

        Dim strCmdLine As String
        Dim strParameters() As String
        Dim intIndex As Integer

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

            blnProceed = False
            If objParseCommandLine.ParseCommandLine Then
                If SetOptionsUsingCommandLineParameters(objParseCommandLine) Then blnProceed = True
            End If

            If blnProceed And Not objParseCommandLine.NeedToShowHelp And mCreateModelXMLParameterFile Then
                If mParameterFilePath Is Nothing OrElse mParameterFilePath.Length = 0 Then
                    mParameterFilePath = System.IO.Path.GetFileNameWithoutExtension(System.Reflection.Assembly.GetExecutingAssembly().Location) & "_ModelSettings.xml"
                End If

                objValidateFastaFile = New clsValidateFastaFile
                objValidateFastaFile.SaveSettingsToParameterFile(mParameterFilePath)

            ElseIf Not blnProceed OrElse objParseCommandLine.NeedToShowHelp OrElse mInputFilePath.Length = 0 Then
                ShowProgramHelp()
                intReturnCode = -1
            Else

                objValidateFastaFile = New clsValidateFastaFile
                With objValidateFastaFile
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

                ''' Note: the following settings will be overridden if mParameterFilePath points to a valid parameter file that has these settings defined
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
                    If objValidateFastaFile.ProcessFilesAndRecurseFolders(mInputFilePath, mOutputFolderPath, mOutputFolderPath, False, mParameterFilePath, mRecurseFoldersMaxLevels) Then
                        intReturnCode = 0
                    Else
                        intReturnCode = objValidateFastaFile.ErrorCode
                    End If
                Else
                    If objValidateFastaFile.ProcessFilesWildcard(mInputFilePath, mOutputFolderPath, mParameterFilePath) Then
                        intReturnCode = 0
                    Else
                        intReturnCode = objValidateFastaFile.ErrorCode
                        If intReturnCode <> 0 AndAlso Not mQuietMode Then
                            MsgBox("Error while processing: " & objValidateFastaFile.GetErrorMessage(), MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
                        End If
                    End If
                End If


            End If

        Catch ex As Exception
            If Not mQuietMode Then
                MsgBox("Error occurred: " & ControlChars.NewLine & ex.Message, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
            End If
            intReturnCode = -1
        End Try

        Return intReturnCode

    End Function


    Private Function SetOptionsUsingCommandLineParameters(ByVal objParseCommandLine As clsParseCommandLine) As Boolean
        ' Returns True if no problems; otherwise, returns false

        Dim strValue As String
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
            If mQuietMode Then
                Throw New System.Exception("Error parsing the command line parameters", ex)
            Else
                MsgBox("Error parsing the command line parameters: " & ControlChars.NewLine & ex.Message, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
            End If
        End Try

    End Function

    Private Sub ShowProgramHelp()

        Dim strSyntax As String
        Dim ioPath As System.IO.Path

        Try

            strSyntax = "This program will read a Fasta File and display statistics on the number of proteins and number of residues.  It will also check that the protein names, descriptions, and sequences are in the correct format." & ControlChars.NewLine
            strSyntax &= "Program syntax:" & ControlChars.NewLine & ioPath.GetFileName(System.Reflection.Assembly.GetExecutingAssembly().Location)
            strSyntax &= " /I:InputFilePath.fasta [/O:OutputFolderPath] [/P:ParameterFilePath] [/C] [/F] [/R] [/D] [/L] [/B] [/X] [/S:[MaxLevel]] [/Q]" & ControlChars.NewLine & ControlChars.NewLine

            strSyntax &= "The input file path can contain the wildcard character * and should point to a fasta file." & ControlChars.NewLine
            strSyntax &= "The output folder path is optional, and is only used if /C is used.  If omitted, the output stats file will be created in the folder containing the .Exe file." & ControlChars.NewLine
            strSyntax &= "Use /C to specify that an output file should be created, rather than displaying the results on the screen." & ControlChars.NewLine & ControlChars.NewLine

            strSyntax &= "Use /F to shorten long protein names and remove invalid characters from the residues line, generating a new, fixed .Fasta file.  At the same time, a file with protein names and hash values for each unique protein sequences will be generated (_UniqueProteinSeqs.txt).  This file will also list the other proteins that have duplicate sequences as the first protein mapped to each sequence.  If duplicate sequences are found, then an easily parseable mapping file will also be created (_UniqueProteinSeqDuplicates.txt)." & ControlChars.NewLine
            strSyntax &= "Use /R to rename proteins with duplicate names when using /F to generate a fixed fasta file." & ControlChars.NewLine
            strSyntax &= "Use /D to consolidate proteins with duplicate protein sequences when using /F to generate a fixed fasta file." & ControlChars.NewLine
            strSyntax &= "Use /L to ignore I/L (isoleucine vs. leucine) differences when consolidating proteins with duplicate protein sequences while generating a fixed fasta file." & ControlChars.NewLine & ControlChars.NewLine

            strSyntax &= "The parameter file path is optional.  If included, it should point to a valid XML parameter file." & ControlChars.NewLine
            strSyntax &= "Use /X to specify that a model XML parameter file should be created." & ControlChars.NewLine
            strSyntax &= "Use /S to process all valid files in the input folder and subfolders. Include a number after /S (like /S:2) to limit the level of subfolders to examine." & ControlChars.NewLine
            strSyntax &= "The optional /Q switch will suppress all error messages." & ControlChars.NewLine & ControlChars.NewLine

            strSyntax &= "Program written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA) in 2005" & ControlChars.NewLine
            strSyntax &= "Program date: " & PROGRAM_DATE & ControlChars.NewLine & ControlChars.NewLine

            strSyntax &= "E-mail: matthew.monroe@pnl.gov or matt@alchemistmatt.com" & ControlChars.NewLine
            strSyntax &= "Website: http://ncrr.pnl.gov/ or http://www.sysbio.org/resources/staff/" & ControlChars.NewLine & ControlChars.NewLine

            strSyntax &= "Licensed under the Apache License, Version 2.0; you may not use this file except in compliance with the License.  "
            strSyntax &= "You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0" & ControlChars.NewLine & ControlChars.NewLine

            strSyntax &= "Notice: This computer software was prepared by Battelle Memorial Institute, "
            strSyntax &= "hereinafter the Contractor, under Contract No. DE-AC05-76RL0 1830 with the "
            strSyntax &= "Department of Energy (DOE).  All rights in the computer software are reserved "
            strSyntax &= "by DOE on behalf of the United States Government and the Contractor as "
            strSyntax &= "provided in the Contract.  NEITHER THE GOVERNMENT NOR THE CONTRACTOR MAKES ANY "
            strSyntax &= "WARRANTY, EXPRESS OR IMPLIED, OR ASSUMES ANY LIABILITY FOR THE USE OF THIS "
            strSyntax &= "SOFTWARE.  This notice including this sentence must appear on any copies of "
            strSyntax &= "this computer software." & ControlChars.NewLine

            If Not mQuietMode Then
                MsgBox(strSyntax, MsgBoxStyle.Information Or MsgBoxStyle.OKOnly, "Syntax")
            End If

        Catch ex As Exception
            If mQuietMode Then
                Throw New System.Exception("Error displaying the program syntax", ex)
            Else
                MsgBox("Error displaying the program syntax: " & ControlChars.NewLine & ex.Message, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
            End If
        End Try

    End Sub

End Module
