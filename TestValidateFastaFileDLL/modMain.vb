Option Strict On

Imports PRISM
Imports ValidateFastaFile

' Program written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
' Copyright 2005, Battelle Memorial Institute.  All Rights Reserved.

' This program can be used to test the use of the clsValidateFastaFiles in the ValidateFastaFiles.Dll file
' Last modified May 8, 2007

Module modMain

    Public Function Main() As Integer
        ' Returns 0 if no error, error code if an error

        Dim testFilePath = "JunkTest.fasta"

        Dim fastaFileValidator As clsValidateFastaFile

        Dim parameters() As String

        Dim success As Boolean

        Dim count As Integer

        Dim returnCode As Integer

        Try
            ' See if the user provided a custom filepath at the command line
            Try
                ' This command will fail if the program is called from a network share
                parameters = Environment.GetCommandLineArgs()

                If Not parameters Is Nothing AndAlso parameters.Length > 1 Then
                    ' Note that parameters(0) is the path to the Executable for the calling program
                    testFilePath = parameters(1)
                End If

            Catch ex As Exception
                ' Ignore errors here
            End Try

            Console.WriteLine("Examining file: " & testFilePath)

            fastaFileValidator = New clsValidateFastaFile()
            With fastaFileValidator
                .SetOptionSwitch(clsValidateFastaFile.SwitchOptions.OutputToStatsFile, True)

                ' Note: the following settings will be overridden if parameter file with these settings defined is provided to .ProcessFile()
                .SetOptionSwitch(clsValidateFastaFile.SwitchOptions.AddMissingLinefeedatEOF, False)
                .SetOptionSwitch(clsValidateFastaFile.SwitchOptions.AllowAsteriskInResidues, True)

                .MaximumFileErrorsToTrack = 5               ' The maximum number of errors for each type of error; the total error count is always available, but detailed information is only saved for this many errors or warnings of each type
                .MinimumProteinNameLength = 3
                .MaximumProteinNameLength = 34

                .SetOptionSwitch(clsValidateFastaFile.SwitchOptions.WarnBlankLinesBetweenProteins, False)
            End With

            ' Analyze the fasta file; returns true if the analysis was successful (even if the file contains errors or warnings)
            success = fastaFileValidator.ProcessFile(testFilePath, String.Empty)

            If success Then
                With fastaFileValidator

                    count = .ErrorWarningCounts(clsValidateFastaFile.eMsgTypeConstants.ErrorMsg, clsValidateFastaFile.ErrorWarningCountTypes.Total)
                    If count = 0 Then
                        Console.WriteLine(" No errors were found")
                    Else
                        Console.WriteLine(" " & count.ToString & " errors were found")
                    End If

                    count = .ErrorWarningCounts(clsValidateFastaFile.eMsgTypeConstants.WarningMsg, clsValidateFastaFile.ErrorWarningCountTypes.Total)
                    If count = 0 Then
                        Console.WriteLine(" No warnings were found")
                    Else
                        Console.WriteLine(" " & count.ToString & " warnings were found")
                    End If

                    '' Could enumerate the errors using the following
                    For index = 0 To count - 1
                        Console.WriteLine(.ErrorMessageTextByIndex(index, ControlChars.Tab))
                    Next index

                End With
            Else
                ConsoleMsgUtils.ShowError("Error calling validateFastaFile.ProcessFile: " & fastaFileValidator.GetErrorMessage())
            End If

        Catch ex As Exception
            ConsoleMsgUtils.ShowError("Error occurred: " & ex.Message)
            returnCode = -1
        End Try

        Return returnCode

    End Function



End Module
