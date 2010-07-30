Option Strict On

' Program written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
' Copyright 2005, Battelle Memorial Institute.  All Rights Reserved.

' This program can be used to test the use of the clsValidateFastaFiles in the ValidateFastaFiles.Dll file
' Last modified May 8, 2007

Module modMain

    Public Function Main() As Integer
        ' Returns 0 if no error, error code if an error

        Dim strTestFilePath As String = "JunkTest.fasta"

        Dim objValidateFastaFile As ValidateFastaFile.clsValidateFastaFile

        Dim strParameters() As String

        Dim blnSuccess As Boolean

        Dim intCount As Integer

        Dim intReturnCode As Integer

        Try
            ' See if the user provided a custom filepath at the command line
            Try
                ' This command will fail if the program is called from a network share
                strParameters = System.Environment.GetCommandLineArgs()

                If Not strParameters Is Nothing AndAlso strParameters.Length > 1 Then
                    ' Note that strParameters(0) is the path to the Executable for the calling program
                    strTestFilePath = strParameters(1)
                End If

            Catch ex As Exception
                ' Ignore errors here
            End Try

            Console.WriteLine("Examining file: " & strTestFilePath)

            objValidateFastaFile = New ValidateFastaFile.clsValidateFastaFile
            With objValidateFastaFile
                .SetOptionSwitch(ValidateFastaFile.IValidateFastaFile.SwitchOptions.OutputToStatsFile, True)

                ' Note: the following settings will be overridden if parameter file with these settings defined is provided to .ProcessFile()
                .SetOptionSwitch(ValidateFastaFile.IValidateFastaFile.SwitchOptions.AddMissingLinefeedatEOF, False)
                .SetOptionSwitch(ValidateFastaFile.IValidateFastaFile.SwitchOptions.AllowAsteriskInResidues, True)

                .MaximumFileErrorsToTrack = 5               ' The maximum number of errors for each type of error; the total error count is always available, but detailed information is only saved for this many errors or warnings of each type
                .MinimumProteinNameLength = 3
                .MaximumProteinNameLength = 34

                .SetOptionSwitch(ValidateFastaFile.IValidateFastaFile.SwitchOptions.WarnBlankLinesBetweenProteins, False)
            End With

            ' Analyze the fasta file; returns true if the analysis was successful (even if the file contains errors or warnings)
            blnSuccess = objValidateFastaFile.ProcessFile(strTestFilePath, String.Empty)

            If blnSuccess Then
                With objValidateFastaFile

                    intCount = .ErrorWarningCounts(ValidateFastaFile.IValidateFastaFile.eMsgTypeConstants.ErrorMsg, ValidateFastaFile.IValidateFastaFile.ErrorWarningCountTypes.Total)
                    If intCount = 0 Then
                        Console.WriteLine(" No errors were found")
                    Else
                        Console.WriteLine(" " & intCount.ToString & " errors were found")
                    End If

                    intCount = .ErrorWarningCounts(ValidateFastaFile.IValidateFastaFile.eMsgTypeConstants.WarningMsg, ValidateFastaFile.IValidateFastaFile.ErrorWarningCountTypes.Total)
                    If intCount = 0 Then
                        Console.WriteLine(" No warnings were found")
                    Else
                        Console.WriteLine(" " & intCount.ToString & " warnings were found")
                    End If

                    '' Could enumerate the errors using the following
                    'For intIndex = 0 To .FileErrorCountSpecified - 1
                    '    Console.WriteLine(.GetFileErrorTextByIndex(intIndex, ControlChars.Tab))
                    'Next intIndex

                End With
            Else
                Console.WriteLine("Error calling objValidateFastaFile.ProcessFile: " & objValidateFastaFile.GetErrorMessage())
            End If

        Catch ex As Exception
            Console.WriteLine("Error occurred: " & ex.Message)
            intReturnCode = -1
        End Try

        Return intReturnCode

    End Function



End Module
