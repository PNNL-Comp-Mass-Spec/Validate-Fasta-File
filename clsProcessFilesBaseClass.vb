Option Strict On

' This class can be used as a base class for classes that process a file or files, and create
' new output files in an output folder.  Note that this class contains simple error codes that
' can be set from any derived classes.  The derived classes can also set their own local error codes
'
' Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
' Copyright 2005, Battelle Memorial Institute.  All Rights Reserved.
' Started November 9, 2003

Public MustInherit Class clsProcessFilesBaseClass

    Public Sub New()
        mFileDate = "January 1, 2006"
        mErrorCode = eProcessFilesErrorCodes.NoError
        mProgressStepDescription = String.Empty
    End Sub

#Region "Enums and Classwide Variables"
    Public Enum eProcessFilesErrorCodes
        NoError = 0
        InvalidInputFilePath = 1
        InvalidOutputFolderPath = 2
        ParameterFileNotFound = 4
        InvalidParameterFile = 8
        FilePathError = 16
        LocalizedError = 32
        UnspecifiedError = -1
    End Enum

    ''' Copy the following to any derived classes
    ''Public Enum eDerivedClassErrorCodes
    ''    NoError = 0
    ''    UnspecifiedError = -1
    ''End Enum

    ''Private mLocalErrorCode As eDerivedClassErrorCodes

    ''Public ReadOnly Property LocalErrorCode() As eDerivedClassErrorCodes
    ''    Get
    ''        Return mLocalErrorCode
    ''    End Get
    ''End Property

    Private mShowMessages As Boolean
    Private mErrorCode As eProcessFilesErrorCodes

    Protected mFileDate As String
    Protected mAbortProcessing As Boolean

    Public Event ProgressReset()
    Public Event ProgressChanged(ByVal taskDescription As String, ByVal percentComplete As Single)     ' PercentComplete ranges from 0 to 100, but can contain decimal percentage values
    Public Event ProgressComplete()

    Protected mProgressStepDescription As String
    Protected mProgressPercentComplete As Single        ' Ranges from 0 to 100, but can contain decimal percentage values

#End Region

#Region "Interface Functions"
    Public Property AbortProcessing() As Boolean
        Get
            Return mAbortProcessing
        End Get
        Set(ByVal Value As Boolean)
            mAbortProcessing = Value
        End Set
    End Property

    Public ReadOnly Property ErrorCode() As eProcessFilesErrorCodes
        Get
            Return mErrorCode
        End Get
    End Property

    Public ReadOnly Property FileVersion() As String
        Get
            FileVersion = GetVersionForExecutingAssembly()
        End Get
    End Property

    Public ReadOnly Property FileDate() As String
        Get
            FileDate = mFileDate
        End Get
    End Property

    Public Overridable ReadOnly Property ProgressStepDescription() As String
        Get
            Return mProgressStepDescription
        End Get
    End Property

    ' ProgressPercentComplete ranges from 0 to 100, but can contain decimal percentage values
    Public ReadOnly Property ProgressPercentComplete() As Single
        Get
            Return CType(Math.Round(mProgressPercentComplete, 2), Single)
        End Get
    End Property

    Public Property ShowMessages() As Boolean
        Get
            Return mShowMessages
        End Get
        Set(ByVal Value As Boolean)
            mShowMessages = Value
        End Set
    End Property
#End Region

    Public Sub AbortProcessingNow()
        mAbortProcessing = True
    End Sub

    Protected Function CleanupFilePaths(ByRef strInputFilePath As String, ByRef strOutputFolderPath As String) As Boolean
        ' Returns True if success, False if failure

        Dim ioFileInfo As System.IO.FileInfo
        Dim ioFolder As System.IO.DirectoryInfo

        Try
            ' Make sure strInputFilePath points to a valid file
            ioFileInfo = New System.IO.FileInfo(strInputFilePath)

            If Not ioFileInfo.Exists() Then
                If Me.ShowMessages Then
                    System.Windows.Forms.MessageBox.Show("Input file not found: " & ControlChars.NewLine & strInputFilePath, "File not found", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation)
                End If
                mErrorCode = eProcessFilesErrorCodes.InvalidInputFilePath
                CleanupFilePaths = False
            Else
                If strOutputFolderPath Is Nothing OrElse strOutputFolderPath.Length = 0 Then
                    ' Define strOutputFolderPath based on strInputFilePath
                    strOutputFolderPath = ioFileInfo.DirectoryName
                End If

                ' Make sure strOutputFolderPath points to a folder
                ioFolder = New System.IO.DirectoryInfo(strOutputFolderPath)

                If Not ioFolder.Exists() Then
                    ' strOutputFolderPath points to a non-existent folder; attempt to create it
                    ioFolder.Create()
                End If

                CleanupFilePaths = True
            End If

        Catch ex As Exception
            If Me.ShowMessages Then
                System.Windows.Forms.MessageBox.Show("Error cleaning up the file paths: " & ControlChars.NewLine & ex.Message, "Error", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation)
            Else
                Throw New System.Exception("Error cleaning up the file paths", ex)
            End If
        End Try

    End Function

    Protected Function GetBaseClassErrorMessage() As String
        ' Returns String.Empty if no error

        Dim strErrorMessage As String

        Select Case Me.ErrorCode
            Case eProcessFilesErrorCodes.NoError
                strErrorMessage = String.Empty
            Case eProcessFilesErrorCodes.InvalidInputFilePath
                strErrorMessage = "Invalid input file path"
            Case eProcessFilesErrorCodes.InvalidOutputFolderPath
                strErrorMessage = "Invalid output folder path"
            Case eProcessFilesErrorCodes.ParameterFileNotFound
                strErrorMessage = "Parameter file not found"
            Case eProcessFilesErrorCodes.InvalidParameterFile
                strErrorMessage = "Invalid parameter file"
            Case eProcessFilesErrorCodes.FilePathError
                strErrorMessage = "General file path error"
            Case eProcessFilesErrorCodes.LocalizedError
                strErrorMessage = "Localized error"
            Case eProcessFilesErrorCodes.UnspecifiedError
                strErrorMessage = "Unspecified error"
            Case Else
                ' This shouldn't happen
                strErrorMessage = "Unknown error state"
        End Select

        Return strErrorMessage

    End Function

    Private Function GetVersionForExecutingAssembly() As String

        Dim strVersion As String

        Try
            strVersion = System.Reflection.Assembly.GetExecutingAssembly.GetName.Version.ToString()
        Catch ex As Exception
            strVersion = "??.??.??.??"
        End Try

        Return strVersion

    End Function

    Public Overridable Function GetDefaultExtensionsToParse() As String()
        Dim strExtensionsToParse(0) As String

        strExtensionsToParse(0) = ".*"

        Return strExtensionsToParse

    End Function

    Public MustOverride Function GetErrorMessage() As String

    Public Function ProcessFilesWildcard(ByVal strInputFilePath As String, ByVal strOutputFolderPath As String) As Boolean
        Return ProcessFilesWildcard(strInputFilePath, strOutputFolderPath, String.Empty)
    End Function

    Public Function ProcessFilesWildcard(ByVal strInputFilePath As String, ByVal strOutputFolderPath As String, ByVal strParameterFilePath As String) As Boolean
        Return ProcessFilesWildcard(strInputFilePath, strOutputFolderPath, strParameterFilePath, True)
    End Function

    Public Function ProcessFilesWildcard(ByVal strInputFilePath As String, ByVal strOutputFolderPath As String, ByVal strParameterFilePath As String, ByVal blnResetErrorCode As Boolean) As Boolean
        ' Returns True if success, False if failure

        Dim blnSuccess As Boolean
        Dim intMatchCount As Integer

        Dim strCleanPath As String
        Dim strInputFolderPath As String

        Dim ioPath As System.IO.Path

        Dim ioFileMatch As System.IO.FileInfo
        Dim ioFileInfo As System.IO.FileInfo
        Dim ioFolderInfo As System.IO.DirectoryInfo

        mAbortProcessing = False
        blnSuccess = True
        Try
            ' Possibly reset the error code
            If blnResetErrorCode Then mErrorCode = eProcessFilesErrorCodes.NoError

            ' See if strInputFilePath contains a wildcard
            If Not strInputFilePath Is Nothing AndAlso (strInputFilePath.IndexOf("*") >= 0 Or strInputFilePath.IndexOf("?") >= 0) Then
                ' Obtain a list of the matching  files

                ' Copy the path into strCleanPath and replace any * or ? characters with _
                strCleanPath = strInputFilePath.Replace("*", "_")
                strCleanPath = strCleanPath.Replace("?", "_")

                ioFileInfo = New System.IO.FileInfo(strCleanPath)
                If ioFileInfo.Directory.Exists Then
                    strInputFolderPath = ioFileInfo.DirectoryName
                Else
                    ' Use the current working directory
                    strInputFolderPath = ioPath.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
                End If

                ioFolderInfo = New System.io.DirectoryInfo(strInputFolderPath)

                ' Remove any directory information from strInputFilePath
                strInputFilePath = ioPath.GetFileName(strInputFilePath)

                intMatchCount = 0
                For Each ioFileMatch In ioFolderInfo.GetFiles(strInputFilePath)

                    blnSuccess = ProcessFile(ioFileMatch.FullName, strOutputFolderPath, strParameterFilePath, blnResetErrorCode)

                    If Not blnSuccess Or mAbortProcessing Then Exit For
                    intMatchCount += 1

                    If intMatchCount Mod 100 = 0 Then Console.Write(".")

                Next ioFileMatch

                If intMatchCount = 0 And Me.ShowMessages Then
                    If mErrorCode = eProcessFilesErrorCodes.NoError Then
                        System.Windows.Forms.MessageBox.Show("No match was found for the input file path:" & ControlChars.NewLine & strInputFilePath, "File not found", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation)
                    End If
                Else
                    Console.WriteLine()
                End If
            Else
                blnSuccess = ProcessFile(strInputFilePath, strOutputFolderPath, strParameterFilePath, blnResetErrorCode)
            End If

        Catch ex As Exception
            If Me.ShowMessages Then
                System.Windows.Forms.MessageBox.Show("Error in ProcessFilesWildcard: " & ControlChars.NewLine & ex.Message, "Error", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation)
            Else
                Throw New System.Exception("Error in ProcessFilesWildcard", ex)
            End If
        End Try

        Return blnSuccess

    End Function

    Public Function ProcessFile(ByVal strInputFilePath As String, ByVal strOutputFolderPath As String) As Boolean
        Return ProcessFile(strInputFilePath, strOutputFolderPath, String.Empty, True)
    End Function

    Public Function ProcessFile(ByVal strInputFilePath As String, ByVal strOutputFolderPath As String, ByVal strParameterFilePath As String) As Boolean
        Return ProcessFile(strInputFilePath, strOutputFolderPath, strParameterFilePath, True)
    End Function

    Public MustOverride Function ProcessFile(ByVal strInputFilePath As String, ByVal strOutputFolderPath As String, ByVal strParameterFilePath As String, ByVal blnResetErrorCode As Boolean) As Boolean

    Public Function ProcessFilesAndRecurseFolders(ByVal strInputFilePathOrFolder As String, ByVal strOutputFolderName As String) As Boolean
        Return ProcessFilesAndRecurseFolders(strInputFilePathOrFolder, strOutputFolderName, String.Empty)
    End Function

    Public Function ProcessFilesAndRecurseFolders(ByVal strInputFilePathOrFolder As String, ByVal strOutputFolderName As String, ByVal strParameterFilePath As String) As Boolean
        Return ProcessFilesAndRecurseFolders(strInputFilePathOrFolder, strOutputFolderName, String.Empty, False, strParameterFilePath)
    End Function

    Public Function ProcessFilesAndRecurseFolders(ByVal strInputFilePathOrFolder As String, ByVal strOutputFolderName As String, ByVal strParameterFilePath As String, ByVal strExtensionsToParse() As String) As Boolean
        Return ProcessFilesAndRecurseFolders(strInputFilePathOrFolder, strOutputFolderName, String.Empty, False, strParameterFilePath, 0, strExtensionsToParse)
    End Function

    Public Function ProcessFilesAndRecurseFolders(ByVal strInputFilePathOrFolder As String, ByVal strOutputFolderName As String, ByVal strOutputFolderAlternatePath As String, ByVal blnRecreateFolderHierarchyInAlternatePath As Boolean) As Boolean
        Return ProcessFilesAndRecurseFolders(strInputFilePathOrFolder, strOutputFolderName, strOutputFolderAlternatePath, blnRecreateFolderHierarchyInAlternatePath, String.Empty)
    End Function

    Public Function ProcessFilesAndRecurseFolders(ByVal strInputFilePathOrFolder As String, ByVal strOutputFolderName As String, ByVal strOutputFolderAlternatePath As String, ByVal blnRecreateFolderHierarchyInAlternatePath As Boolean, ByVal strParameterFilePath As String) As Boolean
        Return ProcessFilesAndRecurseFolders(strInputFilePathOrFolder, strOutputFolderName, strOutputFolderAlternatePath, blnRecreateFolderHierarchyInAlternatePath, strParameterFilePath, 0)
    End Function

    Public Function ProcessFilesAndRecurseFolders(ByVal strInputFilePathOrFolder As String, ByVal strOutputFolderName As String, ByVal strOutputFolderAlternatePath As String, ByVal blnRecreateFolderHierarchyInAlternatePath As Boolean, ByVal strParameterFilePath As String, ByVal intRecurseFoldersMaxLevels As Integer) As Boolean
        Return ProcessFilesAndRecurseFolders(strInputFilePathOrFolder, strOutputFolderName, strOutputFolderAlternatePath, blnRecreateFolderHierarchyInAlternatePath, strParameterFilePath, intRecurseFoldersMaxLevels, GetDefaultExtensionsToParse())
    End Function

    Public Function ProcessFilesAndRecurseFolders(ByVal strInputFilePathOrFolder As String, ByVal strOutputFolderName As String, ByVal strOutputFolderAlternatePath As String, ByVal blnRecreateFolderHierarchyInAlternatePath As Boolean, ByVal strParameterFilePath As String, ByVal intRecurseFoldersMaxLevels As Integer, ByVal strExtensionsToParse() As String) As Boolean
        ' Calls ProcessFiles for all files in strInputFilePathOrFolder and below having an extension listed in strExtensionsToParse()
        ' The extensions should be of the form ".TXT" or ".RAW" (i.e. a period then the extension)
        ' If any of the extensions is "*" or ".*" then all files will be processed
        ' If strInputFilePathOrFolder contains a filename with a wildcard (* or ?), then that information will be 
        '  used to filter the files that are processed
        ' If intRecurseFoldersMaxLevels is <=0 then we recurse infinitely

        Dim strCleanPath As String
        Dim strInputFolderPath As String

        Dim ioFileInfo As System.IO.FileInfo
        Dim ioPath As System.IO.Path
        Dim ioFolderInfo As System.IO.DirectoryInfo

        Dim blnSuccess As Boolean
        Dim intFileProcessCount, intFileProcessFailCount As Integer

        ' Examine strInputFilePathOrFolder to see if it contains a filename; if not, assume it points to a folder
        ' First, see if it contains a * or ?
        Try
            If Not strInputFilePathOrFolder Is Nothing AndAlso (strInputFilePathOrFolder.IndexOf("*") >= 0 Or strInputFilePathOrFolder.IndexOf("?") >= 0) Then
                ' Copy the path into strCleanPath and replace any * or ? characters with _
                strCleanPath = strInputFilePathOrFolder.Replace("*", "_")
                strCleanPath = strCleanPath.Replace("?", "_")

                ioFileInfo = New System.IO.FileInfo(strCleanPath)
                If ioFileInfo.Directory.Exists Then
                    strInputFolderPath = ioFileInfo.DirectoryName
                Else
                    ' Use the current working directory
                    strInputFolderPath = ioPath.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
                End If

                ' Remove any directory information from strInputFilePath
                strInputFilePathOrFolder = ioPath.GetFileName(strInputFilePathOrFolder)

            Else
                ioFolderInfo = New System.IO.DirectoryInfo(strInputFilePathOrFolder)
                If ioFolderInfo.Exists Then
                    strInputFolderPath = ioFolderInfo.FullName
                    strInputFilePathOrFolder = "*"
                Else
                    If ioFolderInfo.Parent.Exists Then
                        strInputFolderPath = ioFolderInfo.Parent.FullName
                        strInputFilePathOrFolder = System.IO.Path.GetFileName(strInputFilePathOrFolder)
                    Else
                        ' Unable to determine the input folder path
                        strInputFolderPath = String.Empty
                    End If
                End If
            End If

            If Not strInputFolderPath Is Nothing AndAlso strInputFolderPath.Length > 0 Then

                ' Validate the output folder path
                If Not strOutputFolderAlternatePath Is Nothing AndAlso strOutputFolderAlternatePath.Length > 0 Then
                    Try
                        ioFolderInfo = New System.IO.DirectoryInfo(strOutputFolderAlternatePath)
                        If Not ioFolderInfo.Exists Then ioFolderInfo.Create()
                    Catch ex As Exception
                        Debug.Assert(False, ex.Message)
                        mErrorCode = clsProcessFilesBaseClass.eProcessFilesErrorCodes.InvalidOutputFolderPath
                        Return False
                    End Try
                End If

                ' Initialize some parameters
                mAbortProcessing = False
                intFileProcessCount = 0
                intFileProcessFailCount = 0

                ' Call RecurseFoldersWork
                blnSuccess = RecurseFoldersWork(strInputFolderPath, strInputFilePathOrFolder, strOutputFolderName, strParameterFilePath, strOutputFolderAlternatePath, blnRecreateFolderHierarchyInAlternatePath, strExtensionsToParse, intFileProcessCount, intFileProcessFailCount, 1, intRecurseFoldersMaxLevels)

            Else
                mErrorCode = clsProcessFilesBaseClass.eProcessFilesErrorCodes.InvalidInputFilePath
                Return False
            End If

        Catch ex As Exception
            Debug.Assert(False, ex.Message)
            blnSuccess = False
        End Try

        Return blnSuccess

    End Function

    Private Function RecurseFoldersWork(ByVal strInputFolderPath As String, ByVal strFileNameMatch As String, ByVal strOutputFolderName As String, ByVal strParameterFilePath As String, ByVal strOutputFolderAlternatePath As String, ByVal blnRecreateFolderHierarchyInAlternatePath As Boolean, ByVal strExtensionsToParse() As String, ByRef intFileProcessCount As Integer, ByRef intFileProcessFailCount As Integer, ByVal intRecursionLevel As Integer, ByVal intRecurseFoldersMaxLevels As Integer) As Boolean
        ' If intRecurseFoldersMaxLevels is <=0 then we recurse infinitely

        Dim ioInputFolderInfo As System.IO.DirectoryInfo
        Dim ioSubFolderInfo As System.IO.DirectoryInfo
        Dim ioFileMatch As System.io.FileInfo

        Dim intExtensionIndex As Integer
        Dim blnProcessAllExtensions As Boolean

        Dim strOutputFolderPathToUse As String
        Dim blnSuccess As Boolean

        Try
            ioInputFolderInfo = New System.IO.DirectoryInfo(strInputFolderPath)
        Catch ex As Exception
            ' Input folder path error
            Debug.Assert(False, ex.Message)
            mErrorCode = eProcessFilesErrorCodes.InvalidInputFilePath
            Return False
        End Try

        Try
            If Not strOutputFolderAlternatePath Is Nothing AndAlso strOutputFolderAlternatePath.Length > 0 Then
                If blnRecreateFolderHierarchyInAlternatePath Then
                    strOutputFolderAlternatePath = System.IO.Path.Combine(strOutputFolderAlternatePath, ioInputFolderInfo.Name)
                End If
                strOutputFolderPathToUse = System.IO.Path.Combine(strOutputFolderAlternatePath, strOutputFolderName)
            Else
                strOutputFolderPathToUse = strOutputFolderName
            End If
        Catch ex As Exception
            ' Output file path error
            Debug.Assert(False, ex.Message)
            mErrorCode = eProcessFilesErrorCodes.InvalidOutputFolderPath
            Return False
        End Try

        Try
            ' Validate strExtensionsToParse()
            For intExtensionIndex = 0 To strExtensionsToParse.Length - 1
                If strExtensionsToParse(intExtensionIndex) Is Nothing Then
                    strExtensionsToParse(intExtensionIndex) = String.Empty
                Else
                    If Not strExtensionsToParse(intExtensionIndex).StartsWith(".") Then
                        strExtensionsToParse(intExtensionIndex) = "." & strExtensionsToParse(intExtensionIndex)
                    End If

                    If strExtensionsToParse(intExtensionIndex) = ".*" Then
                        blnProcessAllExtensions = True
                        Exit For
                    Else
                        strExtensionsToParse(intExtensionIndex) = strExtensionsToParse(intExtensionIndex).ToUpper
                    End If
                End If
            Next intExtensionIndex
        Catch ex As Exception
            Debug.Assert(False, ex.Message)
            mErrorCode = eProcessFilesErrorCodes.UnspecifiedError
            Return False
        End Try

        Try
            Console.WriteLine("Examining " & strInputFolderPath)

            ' Process any matching files in this folder
            blnSuccess = True
            For Each ioFileMatch In ioInputFolderInfo.GetFiles(strFileNameMatch)

                For intExtensionIndex = 0 To strExtensionsToParse.Length - 1
                    If blnProcessAllExtensions OrElse ioFileMatch.Extension.ToUpper = strExtensionsToParse(intExtensionIndex) Then
                        blnSuccess = ProcessFile(ioFileMatch.FullName, strOutputFolderPathToUse, strParameterFilePath, True)
                        If Not blnSuccess Then
                            intFileProcessFailCount += 1
                            blnSuccess = True
                        Else
                            intFileProcessCount += 1
                        End If
                        Exit For
                    End If

                    If mAbortProcessing Then Exit For

                Next intExtensionIndex
            Next ioFileMatch
        Catch ex As Exception
            Debug.Assert(False, ex.Message)
            mErrorCode = eProcessFilesErrorCodes.InvalidInputFilePath
            Return False
        End Try

        If Not mAbortProcessing Then
            ' If intRecurseFoldersMaxLevels is <=0 then we recurse infinitely
            '  otherwise, compare intRecursionLevel to intRecurseFoldersMaxLevels
            If intRecurseFoldersMaxLevels <= 0 OrElse intRecursionLevel <= intRecurseFoldersMaxLevels Then
                ' Call this function for each of the subfolders of ioInputFolderInfo
                For Each ioSubFolderInfo In ioInputFolderInfo.GetDirectories()
                    blnSuccess = RecurseFoldersWork(ioSubFolderInfo.FullName, strFileNameMatch, strOutputFolderName, strParameterFilePath, strOutputFolderAlternatePath, blnRecreateFolderHierarchyInAlternatePath, strExtensionsToParse, intFileProcessCount, intFileProcessFailCount, intRecursionLevel + 1, intRecurseFoldersMaxLevels)
                    If Not blnSuccess Then Exit For
                Next ioSubFolderInfo
            End If
        End If

        Return blnSuccess

    End Function

    Protected Sub ResetProgress()
        RaiseEvent ProgressReset()
    End Sub

    Protected Sub ResetProgress(ByVal strProgressStepDescription As String)
        UpdateProgress(strProgressStepDescription, 0)
        RaiseEvent ProgressReset()
    End Sub

    Protected Sub SetBaseClassErrorCode(ByVal eNewErrorCode As eProcessFilesErrorCodes)
        mErrorCode = eNewErrorCode
    End Sub

    Protected Sub UpdateProgress(ByVal strProgressStepDescription As String)
        UpdateProgress(strProgressStepDescription, mProgressPercentComplete)
    End Sub

    Protected Sub UpdateProgress(ByVal sngPercentComplete As Single)
        UpdateProgress(Me.ProgressStepDescription, sngPercentComplete)
    End Sub

    Protected Sub UpdateProgress(ByVal strProgressStepDescription As String, ByVal sngPercentComplete As Single)
        mProgressStepDescription = String.Copy(strProgressStepDescription)
        If sngPercentComplete < 0 Then
            sngPercentComplete = 0
        ElseIf sngPercentComplete > 100 Then
            sngPercentComplete = 100
        End If
        mProgressPercentComplete = sngPercentComplete

        RaiseEvent ProgressChanged(Me.ProgressStepDescription, Me.ProgressPercentComplete)
    End Sub

    Protected Sub OperationComplete()
        RaiseEvent ProgressComplete()
    End Sub

    '' The following functions should be placed in any derived class
    '' Cannot define as MustOverride since it contains a customized enumerated type (eDerivedClassErrorCodes) in the function declaration

    ''Private Sub SetLocalErrorCode(ByVal eNewErrorCode As eDerivedClassErrorCodes)
    ''    SetLocalErrorCode(eNewErrorCode, False)
    ''End Sub

    ''Private Sub SetLocalErrorCode(ByVal eNewErrorCode As eDerivedClassErrorCodes, ByVal blnLeaveExistingErrorCodeUnchanged As Boolean)
    ''    If blnLeaveExistingErrorCodeUnchanged AndAlso mLocalErrorCode <> eDerivedClassErrorCodes.NoError Then
    ''        ' An error code is already defined; do not change it
    ''    Else
    ''        mLocalErrorCode = eNewErrorCode

    ''        If eNewErrorCode = eDerivedClassErrorCodes.NoError Then
    ''            If MyBase.ErrorCode = clsProcessFilesBaseClass.eProcessFilesErrorCodes.LocalizedError Then
    ''                MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.NoError)
    ''            End If
    ''        Else
    ''            MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.LocalizedError)
    ''        End If
    ''    End If

    ''End Sub

End Class
