Public Class clsCustomValidateFastaFiles
    Inherits clsValidateFastaFile

#Region "Structures and enums"
    Structure udtErrorInfoExtended
        Sub New(
                lineNumber As Integer,
                proteinName As String,
                messageText As String,
                extraInfo As String,
                type As String)

            Me.LineNumber = lineNumber
            Me.ProteinName = proteinName
            Me.MessageText = messageText
            Me.ExtraInfo = extraInfo
            Me.Type = type

        End Sub

        Public LineNumber As Integer
        Public ProteinName As String
        Public MessageText As String
        Public ExtraInfo As String
        Public Type As String
    End Structure

    Enum eValidationOptionConstants As Integer
        AllowAsterisksInResidues = 0
        AllowDashInResidues = 1
        AllowAllSymbolsInProteinNames = 2
    End Enum

    Enum eValidationMessageTypes As Integer
        ErrorMsg = 0
        WarningMsg = 1
    End Enum

#End Region

    Private ReadOnly m_CurrentFileErrors As List(Of udtErrorInfoExtended)
    Private ReadOnly m_CurrentFileWarnings As List(Of udtErrorInfoExtended)

    Private m_CachedFastaFilePath As String

    ' Note: this array gets initialized with space for 10 items
    ' If eValidationOptionConstants gets more than 10 entries, then this array will need to be expanded
    Private ReadOnly mValidationOptions() As Boolean

    ''' <summary>
    ''' Constructor
    ''' </summary>
    Public Sub New()
        MyBase.New()

        FullErrorCollection = New Dictionary(Of String, List(Of udtErrorInfoExtended))
        FullWarningCollection = New Dictionary(Of String, List(Of udtErrorInfoExtended))

        m_CurrentFileErrors = New List(Of udtErrorInfoExtended)
        m_CurrentFileWarnings = New List(Of udtErrorInfoExtended)

        ' Reserve space for tracking up to 10 validation updates (expand later if needed)
        ReDim mValidationOptions(10)
    End Sub

    Public Sub ClearErrorList()
        If Not FullErrorCollection Is Nothing Then
            FullErrorCollection.Clear()
        End If

        If Not FullWarningCollection Is Nothing Then
            FullWarningCollection.Clear()
        End If
    End Sub

    ''' <summary>
    ''' Keys are fasta filename
    ''' Values are the list of fasta file errors
    ''' </summary>
    Public ReadOnly Property FullErrorCollection As Dictionary(Of String, List(Of udtErrorInfoExtended))

    ''' <summary>
    ''' Keys are fasta filename
    ''' Values are the list of fasta file warnings
    ''' </summary>
    Public ReadOnly Property FullWarningCollection As Dictionary(Of String, List(Of udtErrorInfoExtended))

    Public ReadOnly Property FASTAFileValid(FASTAFileName As String) As Boolean
        Get
            If FullErrorCollection Is Nothing Then
                Return True
            Else
                Return Not FullErrorCollection.ContainsKey(FASTAFileName)
            End If
        End Get
    End Property

    Public ReadOnly Property FASTAFileHasWarnings(fastaFileName As String) As Boolean
        Get
            If FullWarningCollection Is Nothing Then
                Return False
            Else
                Return FullWarningCollection.ContainsKey(fastaFileName)
            End If
        End Get
    End Property


    Public ReadOnly Property RecordedFASTAFileErrors(fastaFileName As String) As List(Of udtErrorInfoExtended)
        Get
            Dim errorList As List(Of udtErrorInfoExtended) = Nothing
            If FullErrorCollection.TryGetValue(fastaFileName, errorList) Then
                Return errorList
            End If
            Return New List(Of udtErrorInfoExtended)
        End Get
    End Property

    Public ReadOnly Property RecordedFASTAFileWarnings(fastaFileName As String) As List(Of udtErrorInfoExtended)
        Get
            Dim warningList As List(Of udtErrorInfoExtended) = Nothing
            If FullWarningCollection.TryGetValue(fastaFileName, warningList) Then
                Return warningList
            End If
            Return New List(Of udtErrorInfoExtended)
        End Get
    End Property

    Public ReadOnly Property NumFilesWithErrors As Integer
        Get
            If FullWarningCollection Is Nothing Then
                Return 0
            Else
                Return FullErrorCollection.Count
            End If
        End Get
    End Property

    Private Sub RecordFastaFileProblem(
        lineNumber As Integer,
        proteinName As String,
        errorMessageCode As Integer,
        extraInfo As String,
        messageType As eValidationMessageTypes)

        Dim msgString As String = LookupMessageDescription(errorMessageCode, extraInfo)

        RecordFastaFileProblemToHash(lineNumber, proteinName, msgString, extraInfo, messageType)

    End Sub

    Private Sub RecordFastaFileProblem(
      lineNumber As Integer,
      proteinName As String,
      errorMessage As String,
      extraInfo As String,
      messageType As eValidationMessageTypes)

        RecordFastaFileProblemToHash(lineNumber, proteinName, errorMessage, extraInfo, messageType)
    End Sub

    Private Sub RecordFastaFileProblemToHash(
        lineNumber As Integer,
        proteinName As String,
        messageString As String,
        extraInfo As String,
        messageType As eValidationMessageTypes)

        If Not mFastaFilePath.Equals(m_CachedFastaFilePath) Then
            ' New File being analyzed
            m_CurrentFileErrors.Clear()
            m_CurrentFileWarnings.Clear()

            m_CachedFastaFilePath = String.Copy(mFastaFilePath)
        End If

        If messageType = eValidationMessageTypes.WarningMsg Then
            ' Treat as warning
            m_CurrentFileWarnings.Add(New udtErrorInfoExtended(
                lineNumber, proteinName, messageString, extraInfo, "Warning"))

            FullWarningCollection.Item(Path.GetFileName(m_CachedFastaFilePath)) = m_CurrentFileWarnings
        Else

            ' Treat as error
            m_CurrentFileErrors.Add(New udtErrorInfoExtended(
                lineNumber, proteinName, messageString, extraInfo, "Error"))

            FullErrorCollection.Item(Path.GetFileName(m_CachedFastaFilePath)) = m_CurrentFileErrors
        End If

    End Sub

    Public Sub SetValidationOptions(eValidationOptionName As eValidationOptionConstants, enabled As Boolean)
        mValidationOptions(eValidationOptionName) = enabled
    End Sub

    ''' <summary>
    ''' Calls SimpleProcessFile(), which calls clsValidateFastaFile.ProcessFile to validate filePath
    ''' </summary>
    ''' <param name="filePath"></param>
    ''' <returns>True if the file was successfully processed (even if it contains errors)</returns>
    Public Function StartValidateFASTAFile(filePath As String) As Boolean

        Dim success = SimpleProcessFile(filePath)

        If success Then
            If MyBase.ErrorWarningCounts(eMsgTypeConstants.WarningMsg, ErrorWarningCountTypes.Total) > 0 Then
                ' The file has warnings; we need to record them using RecordFastaFileProblem

                Dim warnings = MyBase.GetFileWarnings()

                For Each item In warnings
                    With item
                        RecordFastaFileProblem(.LineNumber, .ProteinName, .MessageCode, String.Empty, eValidationMessageTypes.WarningMsg)
                    End With
                Next


            End If

            If MyBase.ErrorWarningCounts(eMsgTypeConstants.ErrorMsg, ErrorWarningCountTypes.Total) > 0 Then
                ' The file has errors; we need to record them using RecordFastaFileProblem
                ' However, we might ignore some of the errors

                Dim errors = MyBase.GetFileErrors()

                For Each item In errors
                    With item
                        Dim errorMessage = LookupMessageDescription(.MessageCode, .ExtraInfo)

                        Dim ignoreError = False
                        Select Case errorMessage
                            Case MESSAGE_TEXT_ASTERISK_IN_RESIDUES
                                If mValidationOptions(eValidationOptionConstants.AllowAsterisksInResidues) Then
                                    ignoreError = True
                                End If

                            Case MESSAGE_TEXT_DASH_IN_RESIDUES
                                If mValidationOptions(eValidationOptionConstants.AllowDashInResidues) Then
                                    ignoreError = True
                                End If

                        End Select

                        If Not ignoreError Then
                            RecordFastaFileProblem(.LineNumber, .ProteinName, errorMessage, .ExtraInfo, eValidationMessageTypes.ErrorMsg)
                        End If
                    End With
                Next
            End If

            Return True
        Else
            ' SimpleProcessFile returned False
            Return False
        End If

    End Function

End Class
