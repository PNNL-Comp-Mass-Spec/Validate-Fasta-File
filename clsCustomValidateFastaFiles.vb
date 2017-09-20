Public Interface ICustomValidation
    Inherits IValidateFastaFile

    Enum eValidationOptionConstants As Integer
        AllowAsterisksInResidues = 0
        AllowDashInResidues = 1
        AllowAllSymbolsInProteinNames = 2
    End Enum

    Enum eValidationMessageTypes As Integer
        ErrorMsg = 0
        WarningMsg = 1
    End Enum

    ReadOnly Property FullErrorCollection As Dictionary(Of String, List(Of udtErrorInfoExtended))
    ReadOnly Property FullWarningCollection As Dictionary(Of String, List(Of udtErrorInfoExtended))

    ReadOnly Property RecordedFASTAFileErrors(FASTAFileName As String) As List(Of udtErrorInfoExtended)
    ReadOnly Property RecordedFASTAFileWarnings(FASTAFileName As String) As List(Of udtErrorInfoExtended)

    ReadOnly Property FASTAFileValid(FASTAFileName As String) As Boolean
    ReadOnly Property NumberOfFilesWithErrors As Integer

    ReadOnly Property FASTAFileHasWarnings(FASTAFileName As String) As Boolean

    Sub ClearErrorList()

    Sub SetValidationOptions(eValidationOptionName As eValidationOptionConstants, blnEnabled As Boolean)

    Function StartValidateFASTAFile(filePath As String) As Boolean

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

End Interface

Public Class clsCustomValidateFastaFiles
    Inherits clsValidateFastaFile
    Implements ICustomValidation

    ''' <summary>
    ''' Keys are fasta filename
    ''' Values are the list of fasta file errors
    ''' </summary>
    Private ReadOnly m_FileErrorList As Dictionary(Of String, List(Of ICustomValidation.udtErrorInfoExtended))

    ''' <summary>
    ''' Keys are fasta filename
    ''' Values are the list of fasta file warnings
    ''' </summary>
    Private ReadOnly m_FileWarningList As Dictionary(Of String, List(Of ICustomValidation.udtErrorInfoExtended))

    Private ReadOnly m_CurrentFileErrors As List(Of ICustomValidation.udtErrorInfoExtended)
    Private ReadOnly m_CurrentFileWarnings As List(Of ICustomValidation.udtErrorInfoExtended)

    Private m_CachedFastaFilePath As String

    ' Note: this array gets initialized with space for 10 items
    ' If eValidationOptionConstants gets more than 10 entries, then this array will need to be expanded
    Private ReadOnly mValidationOptions() As Boolean


    Public Sub New()
        MyBase.New()

        m_FileErrorList = New Dictionary(Of String, List(Of ICustomValidation.udtErrorInfoExtended))
        m_FileWarningList = New Dictionary(Of String, List(Of ICustomValidation.udtErrorInfoExtended))

        m_CurrentFileErrors = New List(Of ICustomValidation.udtErrorInfoExtended)
        m_CurrentFileWarnings = New List(Of ICustomValidation.udtErrorInfoExtended)

        ' Reserve space for tracking up to 10 validation updates (expand later if needed)
        ReDim mValidationOptions(10)
    End Sub

    Public Sub ClearErrorList() Implements ICustomValidation.ClearErrorList
        If Not m_FileErrorList Is Nothing Then
            m_FileErrorList.Clear()
        End If

        If Not m_FileWarningList Is Nothing Then
            m_FileWarningList.Clear()
        End If
    End Sub

    Public ReadOnly Property FullErrorCollection As Dictionary(Of String, List(Of ICustomValidation.udtErrorInfoExtended)) _
        Implements ICustomValidation.FullErrorCollection
        Get
            Return m_FileErrorList
        End Get
    End Property

    Public ReadOnly Property FullWarningCollection As Dictionary(Of String, List(Of ICustomValidation.udtErrorInfoExtended)) _
        Implements ICustomValidation.FullWarningCollection
        Get
            Return m_FileWarningList
        End Get
    End Property

    Public ReadOnly Property FASTAFileValid(FASTAFileName As String) As Boolean _
        Implements ICustomValidation.FASTAFileValid
        Get
            If m_FileErrorList Is Nothing Then
                Return True
            Else
                Return Not m_FileErrorList.ContainsKey(FASTAFileName)
            End If
        End Get
    End Property

    Public ReadOnly Property FASTAFileHasWarnings(FASTAFileName As String) As Boolean _
    Implements ICustomValidation.FASTAFileHasWarnings
        Get
            If m_FileWarningList Is Nothing Then
                Return False
            Else
                Return m_FileWarningList.ContainsKey(FASTAFileName)
            End If
        End Get
    End Property


    Public ReadOnly Property RecordedFASTAFileErrors(FASTAFileName As String) As List(Of ICustomValidation.udtErrorInfoExtended) _
        Implements ICustomValidation.RecordedFASTAFileErrors
        Get
            Dim errorList As List(Of ICustomValidation.udtErrorInfoExtended) = Nothing
            If m_FileErrorList.TryGetValue(FASTAFileName, errorList) Then
                Return errorList
            End If
            Return New List(Of ICustomValidation.udtErrorInfoExtended)
        End Get
    End Property

    Public ReadOnly Property RecordedFASTAFileWarnings(FASTAFileName As String) As List(Of ICustomValidation.udtErrorInfoExtended) _
        Implements ICustomValidation.RecordedFASTAFileWarnings
        Get
            Dim warningList As List(Of ICustomValidation.udtErrorInfoExtended) = Nothing
            If m_FileWarningList.TryGetValue(FASTAFileName, warningList) Then
                Return warningList
            End If
            Return New List(Of ICustomValidation.udtErrorInfoExtended)
        End Get
    End Property

    Public ReadOnly Property NumFilesWithErrors As Integer _
        Implements ICustomValidation.NumberOfFilesWithErrors
        Get
            If m_FileWarningList Is Nothing Then
                Return 0
            Else
                Return m_FileErrorList.Count
            End If
        End Get
    End Property

    Private Sub RecordFastaFileProblem(
        intLineNumber As Integer,
        strProteinName As String,
        intErrorMessageCode As Integer,
        strExtraInfo As String,
        eType As ICustomValidation.eValidationMessageTypes)

        Dim msgString As String = LookupMessageDescription(intErrorMessageCode, strExtraInfo)

        RecordFastaFileProblemToHash(intLineNumber, strProteinName, msgString, strExtraInfo, eType)

    End Sub

    Private Sub RecordFastaFileProblem(
     intLineNumber As Integer,
     strProteinName As String,
     strErrorMessage As String,
     strExtraInfo As String,
     eType As ICustomValidation.eValidationMessageTypes)

        RecordFastaFileProblemToHash(intLineNumber, strProteinName, strErrorMessage, strExtraInfo, eType)
    End Sub

    Private Sub RecordFastaFileProblemToHash(
        intLineNumber As Integer,
        strProteinName As String,
        intMessageString As String,
        strExtraInfo As String,
        eType As ICustomValidation.eValidationMessageTypes)

        If Not mFastaFilePath.Equals(m_CachedFastaFilePath) Then
            ' New File being analyzed
            m_CurrentFileErrors.Clear()
            m_CurrentFileWarnings.Clear()

            m_CachedFastaFilePath = String.Copy(mFastaFilePath)
        End If

        If eType = ICustomValidation.eValidationMessageTypes.WarningMsg Then
            ' Treat as warning
            m_CurrentFileWarnings.Add(New ICustomValidation.udtErrorInfoExtended(
                intLineNumber, strProteinName, intMessageString, strExtraInfo, "Warning"))

            m_FileWarningList.Item(Path.GetFileName(m_CachedFastaFilePath)) = m_CurrentFileWarnings
        Else

            ' Treat as error
            m_CurrentFileErrors.Add(New ICustomValidation.udtErrorInfoExtended(
                intLineNumber, strProteinName, intMessageString, strExtraInfo, "Error"))

            m_FileErrorList.Item(Path.GetFileName(m_CachedFastaFilePath)) = m_CurrentFileErrors
        End If

    End Sub

    Public Sub SetValidationOptions(eValidationOptionName As ICustomValidation.eValidationOptionConstants, blnEnabled As Boolean) Implements ICustomValidation.SetValidationOptions
        mValidationOptions(eValidationOptionName) = blnEnabled
    End Sub

    Public Function StartValidateFASTAFile(filePath As String) As Boolean Implements ICustomValidation.StartValidateFASTAFile
        ' This function calls SimpleProcessFile(), which calls clsValidateFastaFile.ProcessFile to validate filePath
        ' The function returns true if the file was successfully processed (even if it contains errors)

        Dim blnSuccess = SimpleProcessFile(filePath)

        If blnSuccess Then
            If MyBase.ErrorWarningCounts(IValidateFastaFile.eMsgTypeConstants.WarningMsg, IValidateFastaFile.ErrorWarningCountTypes.Total) > 0 Then
                ' The file has warnings; we need to record them using RecordFastaFileProblem

                Dim lstWarnings = MyBase.GetFileWarnings()

                For Each item In lstWarnings
                    With item
                        RecordFastaFileProblem(.LineNumber, .ProteinName, .MessageCode, String.Empty, ICustomValidation.eValidationMessageTypes.WarningMsg)
                    End With
                Next


            End If

            If MyBase.ErrorWarningCounts(IValidateFastaFile.eMsgTypeConstants.ErrorMsg, IValidateFastaFile.ErrorWarningCountTypes.Total) > 0 Then
                ' The file has errors; we need to record them using RecordFastaFileProblem
                ' However, we might ignore some of the errors

                Dim lstErrors = MyBase.GetFileErrors()

                For Each item In lstErrors
                    With item
                        Dim strErrorMessage = LookupMessageDescription(.MessageCode, .ExtraInfo)

                        Dim blnIgnoreError = False
                        Select Case strErrorMessage
                            Case MESSAGE_TEXT_ASTERISK_IN_RESIDUES
                                If mValidationOptions(ICustomValidation.eValidationOptionConstants.AllowAsterisksInResidues) Then
                                    blnIgnoreError = True
                                End If

                            Case MESSAGE_TEXT_DASH_IN_RESIDUES
                                If mValidationOptions(ICustomValidation.eValidationOptionConstants.AllowDashInResidues) Then
                                    blnIgnoreError = True
                                End If

                        End Select

                        If Not blnIgnoreError Then
                            RecordFastaFileProblem(.LineNumber, .ProteinName, strErrorMessage, .ExtraInfo, ICustomValidation.eValidationMessageTypes.ErrorMsg)
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
