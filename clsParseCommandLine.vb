Option Strict On

' This class can be used to parse the text following the program name when a 
'  program is started from the command line
'
' Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
' Copyright 2005, Battelle Memorial Institute.  All Rights Reserved.
' Started November 8, 2003
'
' Last modified February 8, 2005

Friend Class clsParseCommandLine

    Private htSwitches As Hashtable
    Private mShowHelp As Boolean

    Public ReadOnly Property NeedToShowHelp() As Boolean
        Get
            Return mShowHelp
        End Get
    End Property

    Public ReadOnly Property ParameterCount() As Integer
        Get
            If Not htSwitches Is Nothing Then
                Return htSwitches.Count
            Else
                Return 0
            End If
        End Get
    End Property

    Public Function InvalidParametersPresent(ByVal strParameterList() As String, Optional ByVal blnCaseSensitive As Boolean = False) As Boolean
        ' Returns true if any of the parameters are not present in strParameterList()

        Dim intIndex As Integer
        Dim blnMatchFound As Boolean

        Try
            Dim iEnum As System.Collections.IDictionaryEnumerator = htSwitches.GetEnumerator()

            Do While iEnum.MoveNext()
                blnMatchFound = False
                For intIndex = 0 To strParameterList.Length - 1
                    If blnCaseSensitive Then
                        If CStr(iEnum.Key) = strParameterList(intIndex) Then
                            blnMatchFound = True
                            Exit For
                        End If
                    Else
                        If CStr(iEnum.Key).ToUpper = strParameterList(intIndex).ToUpper Then
                            blnMatchFound = True
                            Exit For
                        End If
                    End If
                Next intIndex

                If Not blnMatchFound Then Return True
            Loop

        Catch ex As Exception
            Throw New System.Exception("Error in InvalidParametersPresent", ex)
        End Try

    End Function

    Public Function ParseCommandLine(Optional ByVal strSwitchStartChar As Char = "/"c, Optional ByVal strSwitchParameterChar As Char = ":"c) As Boolean
        ' Returns True if any command line parameters were found
        ' Otherwise, returns false
        '
        ' If /? or /help is found, then returns False and sets mShowHelp to True

        Dim strCmdLine As String
        Dim strKey As String, strValue As String
        Dim intCharLoc As Integer

        Dim intIndex As Integer
        Dim strParameters() As String

        htSwitches = New Hashtable

        Try
            Try
                ' This command will fail if the program is called from a network share
                strCmdLine = System.Environment.CommandLine()
                strParameters = System.Environment.GetCommandLineArgs()
            Catch ex As Exception
                MsgBox("This program cannot be run from a network share.  Please map a drive to the network share you are currently accessing or copy the program files and required DLL's to your local computer.")
                mShowHelp = True
                Return False
            End Try

            If strCmdLine Is Nothing OrElse strCmdLine.Length = 0 Then
                Return False
            ElseIf strCmdLine.IndexOf(strSwitchStartChar & "?") > 0 Or strCmdLine.ToLower.IndexOf(strSwitchStartChar & "help") > 0 Then
                mShowHelp = True
                Return False
            End If

            ' Parse the command line
            htSwitches.Clear()

            ' Note that strParameters(0) is the path to the Executable for the calling program
            For intIndex = 1 To strParameters.Length - 1

                ' Look for strSwitchParameterChar in strParameters(intIndex)
                If strParameters(intIndex).Length > 0 Then
                    intCharLoc = strParameters(intIndex).IndexOf(strSwitchParameterChar)

                    strKey = strParameters(intIndex)
                    strValue = ""

                    If intCharLoc >= 0 Then
                        strValue = strKey.Substring(intCharLoc + 1).Trim

                        ' Remove any starting and ending quotation marks
                        strValue = strValue.TrimStart(ControlChars.Quote)
                        strValue = strValue.TrimEnd(ControlChars.Quote)

                        strKey = strKey.Substring(0, intCharLoc)
                    End If

                    If strKey.Substring(0, 1) = "-" Or strKey.Substring(0, 1) = "/" Then
                        strKey = strKey.Substring(1)
                    End If
                    strKey = strKey.Trim

                    ' Note: .Item() will add strKey if it doesn't exist (which is normally the case)
                    htSwitches.Item(strKey) = strValue
                End If
            Next intIndex

            If htSwitches.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Throw New System.Exception("Error in ParseCommandLine", ex)
        End Try

    End Function

    Public Function RetrieveParameter(ByVal intParameterIndex As Integer, ByRef strKey As String, ByRef strValue As String) As Boolean
        ' Returns True if the parameter exists; returns false otherwise

        Dim intIndex As Integer

        Try
            strKey = ""
            strValue = ""
            If intParameterIndex < htSwitches.Count Then
                Dim iEnum As System.Collections.IDictionaryEnumerator = htSwitches.GetEnumerator()

                intIndex = 0
                Do While iEnum.MoveNext()
                    If intIndex = intParameterIndex Then
                        strKey = CStr(iEnum.Key)
                        strValue = CStr(iEnum.Value)
                        Return True
                    End If
                    intIndex += 1
                Loop
            Else
                Return False
            End If
        Catch ex As Exception
            Throw New System.Exception("Error in RetrieveParameter", ex)
        End Try

    End Function

    Public Function RetrieveValueForParameter(ByVal strKey As String, ByRef strValue As String, Optional ByVal blnCaseSensitive As Boolean = False) As Boolean
        ' Returns True if the parameter exists; returns false otherwise

        Try
            strValue = ""
            If blnCaseSensitive Then
                If htSwitches.ContainsKey(strKey) Then
                    strValue = CStr(htSwitches(strKey))
                    Return True
                Else
                    Return False
                End If
            Else
                Dim iEnum As System.Collections.IDictionaryEnumerator = htSwitches.GetEnumerator()

                Do While iEnum.MoveNext()
                    If CStr(iEnum.Key).ToUpper = strKey.ToUpper Then
                        strValue = CStr(htSwitches(iEnum.Key))
                        Return True
                    End If
                Loop
                Return False
            End If
        Catch ex As Exception
            Throw New System.Exception("Error in RetrieveValueForParameter", ex)
        End Try

    End Function

End Class
