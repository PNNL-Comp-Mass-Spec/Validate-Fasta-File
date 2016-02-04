Option Strict On

Imports System.IO
Imports System.Collections.Generic
Imports System.Text
Imports System.Text.RegularExpressions

''' <summary>
''' This class produces an easier-to read stack trace for an exception
''' See the descriptions for functions GetExceptionStackTrace and 
''' GetExceptionStackTraceMultiLine for example text
''' </summary>
''' <remarks></remarks>
Public Class clsStackTraceFormatter

    Public Const STACK_TRACE_TITLE As String = "Stack trace: "
    Public Const STACK_CHAIN_SEPARATOR As String = "-:-"
    Public Const FINAL_FILE_PREFIX As String = " in "

    ''' <summary>
    ''' Parses the StackTrace text of the given exception to return a compact description of the current stack
    ''' </summary>
    ''' <param name="objException"></param>
    ''' <returns>
    ''' String of the form:
    ''' "Stack trace: clsCodeTest.Test-:-clsCodeTest.TestException-:-clsCodeTest.InnerTestException in clsCodeTest.vb:line 86"
    ''' </returns>
    ''' <remarks>Useful for removing the full file paths included in the default stack trace</remarks>
    Public Shared Function GetExceptionStackTrace(objException As Exception) As String

        Dim stackTraceData As IEnumerable(Of String) = GetExceptionStackTraceData(objException)

        Dim sbStackTrace = New StringBuilder()
        For index = 0 To stackTraceData.Count - 1
            If index = stackTraceData.Count - 1 AndAlso stackTraceData(index).StartsWith(FINAL_FILE_PREFIX) Then
                sbStackTrace.Append(stackTraceData(index))
                Exit For
            End If

            If index = 0 Then
                sbStackTrace.Append(STACK_TRACE_TITLE & stackTraceData(index))
            Else
                sbStackTrace.Append(STACK_CHAIN_SEPARATOR & stackTraceData(index))
            End If
        Next

        Return sbStackTrace.ToString()
    End Function

    ''' <summary>
    ''' Parses the StackTrace text of the given exception to return a cleaned up description of the current stack,
    ''' with one line for each function in the call tree
    ''' </summary>
    ''' <param name="ex">Exception</param>
    ''' <returns>
    ''' Stack trace: 
    '''   clsCodeTest.Test
    '''   clsCodeTest.TestException
    '''   clsCodeTest.InnerTestException 
    '''    in clsCodeTest.vb:line 86
    ''' </returns>
    ''' <remarks>Useful for removing the full file paths included in the default stack trace</remarks>
    Public Shared Function GetExceptionStackTraceMultiLine(ex As Exception) As String

        Dim stackTraceData As IEnumerable(Of String) = GetExceptionStackTraceData(ex)

        Dim sbStackTrace = New StringBuilder()
        sbStackTrace.AppendLine(STACK_TRACE_TITLE)

        For index = 0 To stackTraceData.Count - 1
            sbStackTrace.AppendLine("  " & stackTraceData(index))
        Next

        Return sbStackTrace.ToString()

    End Function

    ''' <summary>
    ''' Parses the StackTrace text of the given exception to return a cleaned up description of the current stack
    ''' </summary>
    ''' <param name="ex">Exception</param>
    ''' <returns>
    ''' List of function names; for example:
    '''   clsCodeTest.Test
    '''   clsCodeTest.TestException
    '''   clsCodeTest.InnerTestException 
    '''    in clsCodeTest.vb:line 86
    ''' </returns>
    ''' <remarks></remarks>
    Public Shared Function GetExceptionStackTraceData(ex As Exception) As IEnumerable(Of String)
        Return GetExceptionStackTraceData(ex.StackTrace)
    End Function

    ''' <summary>
    ''' Parses the given StackTrace text to return a cleaned up description of the current stack
    ''' </summary>
    ''' <param name="stackTraceText">Exception.StackTrace data</param>
    ''' <returns>
    ''' List of function names; for example:
    '''   clsCodeTest.Test
    '''   clsCodeTest.TestException
    '''   clsCodeTest.InnerTestException 
    '''    in clsCodeTest.vb:line 86
    ''' </returns>
    ''' <remarks></remarks>
    Public Shared Function GetExceptionStackTraceData(stackTraceText As String) As IEnumerable(Of String)
        Const REGEX_FUNCTION_NAME = "at ([^(]+)\("
        Const REGEX_FILE_NAME = "in .+\\(.+)"

        Const CODE_LINE_PREFIX = ":line "
        Const REGEX_LINE_IN_CODE = CODE_LINE_PREFIX & "\d+"

        Dim lstFunctions = New List(Of String)
        Dim strFinalFile As String = String.Empty

        Dim reFunctionName As New Regex(REGEX_FUNCTION_NAME, RegexOptions.Compiled Or RegexOptions.IgnoreCase)
        Dim reFileName As New Regex(REGEX_FILE_NAME, RegexOptions.Compiled Or RegexOptions.IgnoreCase)
        Dim reLineInCode As New Regex(REGEX_LINE_IN_CODE, RegexOptions.Compiled Or RegexOptions.IgnoreCase)

        If String.IsNullOrWhiteSpace(stackTraceText) Then
            Dim emptyStackTrace = New List(Of String)
            emptyStackTrace.Add("Empty stack trace")
            Return emptyStackTrace
        End If

        ' Process each line in objException.StackTrace
        ' Populate strFunctions() with the function name of each line
        Using trTextReader = New StringReader(stackTraceText)

            Do While trTextReader.Peek > -1
                Dim strLine = trTextReader.ReadLine()

                If Not String.IsNullOrEmpty(strLine) Then
                    Dim strCurrentFunction = String.Empty

                    Dim functionMatch = reFunctionName.Match(strLine)
                    Dim lineMatch = reLineInCode.Match(strLine)

                    ' Also extract the file name where the Exception occurred
                    Dim fileMatch = reFileName.Match(strLine)
                    Dim currentFunctionFile As String

                    If fileMatch.Success Then
                        currentFunctionFile = fileMatch.Groups(1).Value
                        If strFinalFile.Length = 0 Then
                            Dim lineMatchFinalFile = reLineInCode.Match(currentFunctionFile)
                            If lineMatchFinalFile.Success Then
                                strFinalFile = currentFunctionFile.Substring(0, lineMatchFinalFile.Index)
                            Else
                                strFinalFile = String.Copy(currentFunctionFile)
                            End If
                        End If
                    Else
                        currentFunctionFile = String.Empty
                    End If

                    If functionMatch.Success Then
                        strCurrentFunction = functionMatch.Groups(1).Value
                    Else
                        ' Look for the word " in "
                        Dim intIndex = strLine.ToLower().IndexOf(" in ", StringComparison.Ordinal)
                        If intIndex = 0 Then
                            ' " in" not found; look for the first space after startIndex 4
                            intIndex = strLine.IndexOf(" ", 4, StringComparison.Ordinal)
                        End If

                        If intIndex = 0 Then
                            ' Space not found; use the entire string
                            intIndex = strLine.Length - 1
                        End If

                        If intIndex > 0 Then
                            strCurrentFunction = strLine.Substring(0, intIndex)
                        End If

                    End If

                    Dim functionDescription As String = String.Copy(strCurrentFunction)

                    If Not String.IsNullOrEmpty(currentFunctionFile) Then
                        If String.IsNullOrEmpty(strFinalFile) OrElse
                           Not TrimLinePrefix(strFinalFile, CODE_LINE_PREFIX).Equals(
                               TrimLinePrefix(currentFunctionFile, CODE_LINE_PREFIX),
                               StringComparison.InvariantCulture) Then
                            functionDescription &= FINAL_FILE_PREFIX & currentFunctionFile
                        End If
                    End If

                    If lineMatch.Success AndAlso Not functionDescription.Contains(CODE_LINE_PREFIX) Then
                        functionDescription &= lineMatch.Value
                    End If

                    lstFunctions.Add(functionDescription)

                End If
            Loop

        End Using

        Dim stackTraceData = New List(Of String)
        stackTraceData.AddRange(lstFunctions)
        stackTraceData.Reverse()

        If Not String.IsNullOrWhiteSpace(strFinalFile) Then
            stackTraceData.Add(FINAL_FILE_PREFIX & strFinalFile)
        End If

        Return stackTraceData

    End Function

    Private Shared Function TrimLinePrefix(strFileDescription As String, codeLinePrefix As String) As String

        Dim matchIndex = strFileDescription.IndexOf(codeLinePrefix, StringComparison.Ordinal)
        If matchIndex > 0 Then
            Return strFileDescription.Substring(0, matchIndex)
        End If

        Return strFileDescription
    End Function

End Class
