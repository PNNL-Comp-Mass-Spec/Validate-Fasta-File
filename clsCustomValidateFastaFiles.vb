Public Interface ICustomValidation
    Inherits IValidateFastaFile

    Enum eValidationOptionConstants As Integer
        AllowAsterisksInResidues = 0
        AllowDashInResidues = 1
    End Enum

    Enum eValidationMessageTypes As Integer
        ErrorMsg = 0
        WarningMsg = 1
    End Enum

    ReadOnly Property FullErrorCollection() As Hashtable
    ReadOnly Property FullWarningCollection() As Hashtable

    ReadOnly Property RecordedFASTAFileErrors(ByVal FASTAFileName As String) As ArrayList
    ReadOnly Property RecordedFASTAFileWarnings(ByVal FASTAFileName As String) As ArrayList

    ReadOnly Property FASTAFileValid(ByVal FASTAFileName As String) As Boolean
    ReadOnly Property NumberOfFilesWithErrors() As Integer

    ReadOnly Property FASTAFileHasWarnings(ByVal FASTAFileName As String) As Boolean

    Sub ClearErrorList()

    Sub SetValidationOptions(ByVal eValidationOptionName As eValidationOptionConstants, ByVal blnEnabled As Boolean)

    Function StartValidateFASTAFile(ByVal FASTAFilePath As String) As Boolean

    Structure udtErrorInfoExtended
        Sub New( _
            ByVal LineNumber As Integer, _
            ByVal ProteinName As String, _
            ByVal MessageText As String, _
            ByVal ExtraInfo As String, _
            ByVal Type As String)

            Me.LineNumber = LineNumber
            Me.ProteinName = ProteinName
            Me.MessageText = MessageText
            Me.ExtraInfo = ExtraInfo
            Me.Type = Type

        End Sub
        Public LineNumber As Integer
        Public ProteinName As String
        Public MessageText As String
        Public ExtraInfo As String
        Public Type As String
    End Structure

End Interface


Public Class clsCustomValidateFastaFiles
    Inherits ValidateFastaFile.clsValidateFastaFile
    Implements ICustomValidation

    Protected m_FileErrorList As Hashtable
    Protected m_FileWarningList As Hashtable

    Protected m_CurrentFileErrors As ArrayList
    Protected m_CurrentFileWarnings As ArrayList

    Protected m_CachedFastaFilePath As String

    ' Note: this array gets initialized with space for 10 items
    ' If eValidationOptionConstants gets more than 10 entries, then this array will need to be expanded
    Protected mValidationOptions() As Boolean


    Public Sub New()
        MyBase.New()
        Me.m_FileErrorList = New Hashtable  'key is fasta filename, value is CurrentFileErrors HashTable for that file
        Me.m_FileWarningList = New Hashtable  'key is fasta filename, value is CurrentFileWarnings HashTable for that file

        ' Reserve space for tracking up to 10 validation updates (expand later if needed)
        ReDim mValidationOptions(10)
    End Sub

    Public Sub ClearErrorList() Implements ICustomValidation.ClearErrorList
        If Not Me.m_FileErrorList Is Nothing Then
            Me.m_FileErrorList.Clear()
        End If

        If Not Me.m_FileWarningList Is Nothing Then
            Me.m_FileWarningList.Clear()
        End If
    End Sub

    Public ReadOnly Property FullErrorCollection() As Hashtable _
        Implements ICustomValidation.FullErrorCollection
        Get
            Return Me.m_FileErrorList
        End Get
    End Property

    Public ReadOnly Property FullWarningCollection() As Hashtable _
     Implements ICustomValidation.FullWarningCollection
        Get
            Return Me.m_FileWarningList
        End Get
    End Property

    Public ReadOnly Property FASTAFileValid(ByVal FASTAFileName As String) As Boolean _
        Implements ICustomValidation.FASTAFileValid
        Get
            If Me.m_FileErrorList Is Nothing Then
                Return True
            Else
                Return Not Me.m_FileErrorList.ContainsKey(FASTAFileName)
            End If
        End Get
    End Property

    Public ReadOnly Property FASTAFileHasWarnings(ByVal FASTAFileName As String) As Boolean _
    Implements ICustomValidation.FASTAFileHasWarnings
        Get
            If Me.m_FileWarningList Is Nothing Then
                Return False
            Else
                Return Me.m_FileWarningList.ContainsKey(FASTAFileName)
            End If
        End Get
    End Property


    Public ReadOnly Property RecordedFASTAFileErrors(ByVal FASTAFileName As String) As ArrayList _
        Implements ICustomValidation.RecordedFASTAFileErrors
        Get
            Return DirectCast(Me.m_FileErrorList(FASTAFileName), ArrayList)
        End Get
    End Property

    Public ReadOnly Property RecordedFASTAFileWarnings(ByVal FASTAFileName As String) As ArrayList _
    Implements ICustomValidation.RecordedFASTAFileWarnings
        Get
            Return DirectCast(Me.m_FileWarningList(FASTAFileName), ArrayList)
        End Get
    End Property

    Public ReadOnly Property NumFilesWithErrors() As Integer _
        Implements ICustomValidation.NumberOfFilesWithErrors
        Get
            If m_FileWarningList Is Nothing Then
                Return 0
            Else
                Return Me.m_FileErrorList.Count
            End If
        End Get
    End Property

    Protected Sub RecordFastaFileProblem( _
        ByVal intLineNumber As Integer, _
        ByVal strProteinName As String, _
        ByVal intErrorMessageCode As Integer, _
        ByVal strExtraInfo As String, _
        ByVal eType As ICustomValidation.eValidationMessageTypes)

        Dim msgString As String = Me.LookupMessageDescription(intErrorMessageCode, strExtraInfo)

        Me.RecordFastaFileProblemToHash(intLineNumber, strProteinName, msgString, strExtraInfo, eType)

    End Sub

    Protected Sub RecordFastaFileProblem( _
     ByVal intLineNumber As Integer, _
     ByVal strProteinName As String, _
     ByVal strErrorMessage As String, _
     ByVal strExtraInfo As String, _
     ByVal eType As ICustomValidation.eValidationMessageTypes)

        Me.RecordFastaFileProblemToHash(intLineNumber, strProteinName, strErrorMessage, strExtraInfo, eType)
    End Sub

    Private Sub RecordFastaFileProblemToHash( _
        ByVal intLineNumber As Integer, _
        ByVal strProteinName As String, _
        ByVal intMessageString As String, _
        ByVal strExtraInfo As String, _
        ByVal eType As ICustomValidation.eValidationMessageTypes)

        If Not Me.mFastaFilePath.Equals(Me.m_CachedFastaFilePath) Then
            ' New File being analyzed
            Me.m_CurrentFileErrors = New ArrayList
            Me.m_CurrentFileWarnings = New ArrayList

            Me.m_CachedFastaFilePath = String.Copy(mFastaFilePath)
        End If

        If eType = ICustomValidation.eValidationMessageTypes.WarningMsg Then
            ' Treat as warning
            Me.m_CurrentFileWarnings.Add(New ICustomValidation.udtErrorInfoExtended( _
                intLineNumber, strProteinName, intMessageString, strExtraInfo, "Warning"))

            Me.m_FileWarningList.Item(System.IO.Path.GetFileName(Me.m_CachedFastaFilePath)) = Me.m_CurrentFileWarnings
        Else

            ' Treat as error
            Me.m_CurrentFileErrors.Add(New ICustomValidation.udtErrorInfoExtended( _
                intLineNumber, strProteinName, intMessageString, strExtraInfo, "Error"))

            Me.m_FileErrorList.Item(System.IO.Path.GetFileName(Me.m_CachedFastaFilePath)) = Me.m_CurrentFileErrors
        End If

    End Sub

    Public Sub SetValidationOptions(ByVal eValidationOptionName As ICustomValidation.eValidationOptionConstants, ByVal blnEnabled As Boolean) Implements ICustomValidation.SetValidationOptions
        mValidationOptions(eValidationOptionName) = blnEnabled
    End Sub

    Public Function StartValidateFASTAFile(ByVal FASTAFilePath As String) As Boolean Implements ICustomValidation.StartValidateFASTAFile
        ' This function calls SimpleProcessFile(), which calls clsValidateFastaFile.ProcessFile to validate FASTAFilePath
        ' The function returns true if the file was successfully processed (even if it contains errors)

        Dim udtErrors() As IValidateFastaFile.udtMsgInfoType
        Dim udtWarnings() As IValidateFastaFile.udtMsgInfoType

        Dim intErrorCount As Integer
        Dim intIndex As Integer

        Dim strErrorMessage As String

        Dim blnIgnoreError As Boolean
        Dim blnSuccess As Boolean

        intErrorCount = 0
        blnSuccess = SimpleProcessFile(FASTAFilePath)

        If blnSuccess Then
            If MyBase.ErrorWarningCounts(IValidateFastaFile.eMsgTypeConstants.WarningMsg, IValidateFastaFile.ErrorWarningCountTypes.Total) > 0 Then
                ' The file has warnings; we need to record them using RecordFastaFileProblem

                udtWarnings = MyBase.GetFileWarnings()

                For intIndex = 0 To udtWarnings.Length - 1
                    With udtWarnings(intIndex)
                        RecordFastaFileProblem(.LineNumber, .ProteinName, .MessageCode, String.Empty, ICustomValidation.eValidationMessageTypes.WarningMsg)
                    End With
                Next


            End If

            If MyBase.ErrorWarningCounts(IValidateFastaFile.eMsgTypeConstants.ErrorMsg, IValidateFastaFile.ErrorWarningCountTypes.Total) > 0 Then
                ' The file has errors; we need to record them using RecordFastaFileProblem
                ' However, we might ignore some of the errors

                udtErrors = MyBase.GetFileErrors()

                For intIndex = 0 To udtErrors.Length - 1
                    With udtErrors(intIndex)
                        strErrorMessage = Me.LookupMessageDescription(.MessageCode, .ExtraInfo)

                        blnIgnoreError = False
                        Select Case strErrorMessage
                            Case ValidateFastaFile.clsValidateFastaFile.MESSAGE_TEXT_ASTERISK_IN_RESIDUES
                                If mValidationOptions(ICustomValidation.eValidationOptionConstants.AllowAsterisksInResidues) Then
                                    blnIgnoreError = True
                                End If

                            Case ValidateFastaFile.clsValidateFastaFile.MESSAGE_TEXT_DASH_IN_RESIDUES
                                If mValidationOptions(ICustomValidation.eValidationOptionConstants.AllowDashInResidues) Then
                                    blnIgnoreError = True
                                End If

                        End Select

                        If Not blnIgnoreError Then
                            intErrorCount += 1
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
