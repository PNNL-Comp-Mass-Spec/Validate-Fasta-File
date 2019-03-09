Public Class clsCustomValidateFastaFiles
    Inherits clsValidateFastaFile

#Region "Structures and enums"
    Structure udtErrorInfoExtended
        Sub New(
                intLineNumber As Integer,
                strProteinName As String,
                strMessageText As String,
                strExtraInfo As String,
                strType As String)

            LineNumber = intLineNumber
            ProteinName = strProteinName
            MessageText = strMessageText
            ExtraInfo = strExtraInfo
            Type = strType

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

    Public ReadOnly Property FASTAFileHasWarnings(FASTAFileName As String) As Boolean
        Get
            If FullWarningCollection Is Nothing Then
                Return False
            Else
                Return FullWarningCollection.ContainsKey(FASTAFileName)
            End If
        End Get
    End Property


    Public ReadOnly Property RecordedFASTAFileErrors(FASTAFileName As String) As List(Of udtErrorInfoExtended)
        Get
            Dim errorList As List(Of udtErrorInfoExtended) = Nothing
            If FullErrorCollection.TryGetValue(FASTAFileName, errorList) Then
                Return errorList
            End If
            Return New List(Of udtErrorInfoExtended)
        End Get
    End Property

    Public ReadOnly Property RecordedFASTAFileWarnings(FASTAFileName As String) As List(Of udtErrorInfoExtended)
        Get
            Dim warningList As List(Of udtErrorInfoExtended) = Nothing
            If FullWarningCollection.TryGetValue(FASTAFileName, warningList) Then
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
        intLineNumber As Integer,
        strProteinName As String,
        intErrorMessageCode As Integer,
        strExtraInfo As String,
        eType As eValidationMessageTypes)

        Dim msgString As String = LookupMessageDescription(intErrorMessageCode, strExtraInfo)

        RecordFastaFileProblemToHash(intLineNumber, strProteinName, msgString, strExtraInfo, eType)

    End Sub

    Private Sub RecordFastaFileProblem(
     intLineNumber As Integer,
     strProteinName As String,
     strErrorMessage As String,
     strExtraInfo As String,
     eType As eValidationMessageTypes)

        RecordFastaFileProblemToHash(intLineNumber, strProteinName, strErrorMessage, strExtraInfo, eType)
    End Sub

    Private Sub RecordFastaFileProblemToHash(
        intLineNumber As Integer,
        strProteinName As String,
        intMessageString As String,
        strExtraInfo As String,
        eType As eValidationMessageTypes)

        If Not mFastaFilePath.Equals(m_CachedFastaFilePath) Then
            ' New File being analyzed
            m_CurrentFileErrors.Clear()
            m_CurrentFileWarnings.Clear()

            m_CachedFastaFilePath = String.Copy(mFastaFilePath)
        End If

        If eType = eValidationMessageTypes.WarningMsg Then
            ' Treat as warning
            m_CurrentFileWarnings.Add(New udtErrorInfoExtended(
                intLineNumber, strProteinName, intMessageString, strExtraInfo, "Warning"))

            FullWarningCollection.Item(Path.GetFileName(m_CachedFastaFilePath)) = m_CurrentFileWarnings
        Else

            ' Treat as error
            m_CurrentFileErrors.Add(New udtErrorInfoExtended(
                intLineNumber, strProteinName, intMessageString, strExtraInfo, "Error"))

            FullErrorCollection.Item(Path.GetFileName(m_CachedFastaFilePath)) = m_CurrentFileErrors
        End If

    End Sub

    Public Sub SetValidationOptions(eValidationOptionName As eValidationOptionConstants, blnEnabled As Boolean)
        mValidationOptions(eValidationOptionName) = blnEnabled
    End Sub

    Public Function StartValidateFASTAFile(filePath As String) As Boolean
        ' This function calls SimpleProcessFile(), which calls clsValidateFastaFile.ProcessFile to validate filePath
        ' The function returns true if the file was successfully processed (even if it contains errors)

        Dim blnSuccess = SimpleProcessFile(filePath)

        If blnSuccess Then
            If MyBase.ErrorWarningCounts(eMsgTypeConstants.WarningMsg, ErrorWarningCountTypes.Total) > 0 Then
                ' The file has warnings; we need to record them using RecordFastaFileProblem

                Dim lstWarnings = MyBase.GetFileWarnings()

                For Each item In lstWarnings
                    With item
                        RecordFastaFileProblem(.LineNumber, .ProteinName, .MessageCode, String.Empty, eValidationMessageTypes.WarningMsg)
                    End With
                Next


            End If

            If MyBase.ErrorWarningCounts(eMsgTypeConstants.ErrorMsg, ErrorWarningCountTypes.Total) > 0 Then
                ' The file has errors; we need to record them using RecordFastaFileProblem
                ' However, we might ignore some of the errors

                Dim lstErrors = MyBase.GetFileErrors()

                For Each item In lstErrors
                    With item
                        Dim strErrorMessage = LookupMessageDescription(.MessageCode, .ExtraInfo)

                        Dim blnIgnoreError = False
                        Select Case strErrorMessage
                            Case MESSAGE_TEXT_ASTERISK_IN_RESIDUES
                                If mValidationOptions(eValidationOptionConstants.AllowAsterisksInResidues) Then
                                    blnIgnoreError = True
                                End If

                            Case MESSAGE_TEXT_DASH_IN_RESIDUES
                                If mValidationOptions(eValidationOptionConstants.AllowDashInResidues) Then
                                    blnIgnoreError = True
                                End If

                        End Select

                        If Not blnIgnoreError Then
                            RecordFastaFileProblem(.LineNumber, .ProteinName, strErrorMessage, .ExtraInfo, eValidationMessageTypes.ErrorMsg)
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
