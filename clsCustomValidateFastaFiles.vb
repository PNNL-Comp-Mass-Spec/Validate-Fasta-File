Public Interface ICustomValidation
    Inherits IValidateFastaFile

    ReadOnly Property FullErrorCollection() As Hashtable
    ReadOnly Property RecordedFASTAFileErrors(ByVal FASTAFilePath As String) As ArrayList
    ReadOnly Property FASTAFileValid(ByVal FASTAFilePath As String) As Boolean
    ReadOnly Property NumberOfFilesWithErrors() As Integer

    Sub ClearErrorList()

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
    Protected m_CurrentFileErrors As ArrayList
    Protected m_CachedFastaFileName As String

      Public Sub New()
        MyBase.New()
        Me.m_FileErrorList = New Hashtable  'key is fasta filename, value is currentfileerrors hashtable for that file

    End Sub

    Public Sub ClearErrorList() Implements ICustomValidation.ClearErrorList
        If Not Me.m_FileErrorList Is Nothing Then
            Me.m_FileErrorList.Clear()
        End If
    End Sub

    Public ReadOnly Property FullErrorCollection() As Hashtable _
        Implements ICustomValidation.FullErrorCollection
        Get
            Return Me.m_FileErrorList
        End Get
    End Property

    Public ReadOnly Property FASTAFileValid(ByVal FASTAFileName As String) As Boolean _
        Implements ICustomValidation.FASTAFileValid
        Get
            Return Not Me.m_FileErrorList.ContainsKey(FASTAFileName)
        End Get
    End Property

    Public ReadOnly Property RecordedFASTAFileErrors(ByVal FASTAFileName As String) As ArrayList _
        Implements ICustomValidation.RecordedFASTAFileErrors
        Get
            Return DirectCast(Me.m_FileErrorList(FASTAFileName), ArrayList)
        End Get
    End Property

    Public ReadOnly Property NumFilesWithErrors() As Integer _
        Implements ICustomValidation.NumberOfFilesWithErrors
        Get
            Return Me.m_FileErrorList.Count
        End Get
    End Property

    Protected Sub RecordFastaFileWarning( _
        ByVal intLineNumber As Integer, _
        ByVal strProteinName As String, _
        ByVal eWarningMessageCode As eMessageCodeConstants)

        RecordFastaFileProblem(intLineNumber, strProteinName, eWarningMessageCode, String.Empty)
    End Sub


    Protected Sub RecordFastaFileProblem( _
        ByVal intLineNumber As Integer, _
        ByVal strProteinName As String, _
        ByVal eErrorMessageCode As eMessageCodeConstants, _
        ByVal strExtraInfo As String)

        Dim msgString As String = Me.LookupMessageDescription(eErrorMessageCode, strExtraInfo)

        Me.RecordFastaFileProblemToHash(intLineNumber, strProteinName, msgString, strExtraInfo, "Error")

    End Sub

    Private Sub RecordFastaFileProblemToHash( _
        ByVal intLineNumber As Integer, _
        ByVal strProteinName As String, _
        ByVal intMessageString As String, _
        ByVal strExtraInfo As String, _
        ByVal type As String)



        If Me.mFastaFilePath.Equals(Me.m_CachedFastaFileName) = False Then
            If Me.m_CachedFastaFileName Is Nothing Then
                Me.m_CachedFastaFileName = Me.mFastaFilePath
            End If
            'New File being analyzed
            Me.m_CurrentFileErrors = New ArrayList
            Me.m_CachedFastaFileName = String.Copy(mFastaFilePath)

        End If


        Me.m_CurrentFileErrors.Add(New ICustomValidation.udtErrorInfoExtended( _
            intLineNumber, strProteinName, intMessageString, strExtraInfo, type))
        Me.m_FileErrorList.Item(System.IO.Path.GetFileName(Me.m_CachedFastaFileName)) = Me.m_CurrentFileErrors

        Me.m_CachedFastaFileName = String.Copy(mFastaFilePath)

    End Sub

End Class
