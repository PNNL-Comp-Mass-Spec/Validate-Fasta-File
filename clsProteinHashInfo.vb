Public Class clsProteinHashInfo
    Private ReadOnly mSequenceHash As String
    Private ReadOnly mSequenceLength As Integer
    Private ReadOnly mSequenceStart As String                  ' The first 20 residues of the protein sequence
    Private ReadOnly mProteinNameFirst As String

    ''' <summary>
    ''' Additional protein names
    ''' </summary>
    ''' <remarks>mProteinNameFirst is not stored here; only additional proteins</remarks>
    Private mAdditionalProteins As SortedSet(Of String)

    ''' <summary>
    ''' SHA-1 has of the protein sequence
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property SequenceHash As String
        Get
            Return mSequenceHash
        End Get
    End Property

    ''' <summary>
    ''' Number of residues in the protein sequence
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property SequenceLength As Integer
        Get
            Return mSequenceLength
        End Get
    End Property

    ''' <summary>
    ''' The first 20 residues of the protein sequence
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property SequenceStart As String
        Get
            Return mSequenceStart
        End Get
    End Property

    ''' <summary>
    ''' First protein associated with this hash value
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ProteinNameFirst As String
        Get
            Return mProteinNameFirst
        End Get
    End Property

    Public ReadOnly Property AdditionalProteins As IEnumerable(Of String)
        Get
            Return mAdditionalProteins
        End Get
    End Property

    ''' <summary>
    ''' Greater than 0 if multiple entries have the same name and same sequence
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DuplicateProteinNameCount As Integer

    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <param name="seqHash"></param>
    ''' <param name="sbResidues"></param>
    ''' <param name="proteinName"></param>
    ''' <remarks></remarks>
    Public Sub New(seqHash As String, sbResidues As Text.StringBuilder, proteinName As String)
        mSequenceHash = seqHash

        mSequenceLength = sbResidues.Length
        mSequenceStart = sbResidues.ToString.Substring(0, Math.Min(sbResidues.Length, 20))

        mProteinNameFirst = proteinName
        mAdditionalProteins = New SortedSet(Of String)
        DuplicateProteinNameCount = 0
    End Sub

    Public Sub AddAdditionalProtein(proteinName As String)
        If Not mAdditionalProteins.Contains(proteinName) Then
            mAdditionalProteins.Add(proteinName)
        End If
    End Sub

    Public Overrides Function ToString() As String
        Return mProteinNameFirst & ": " & mSequenceHash
    End Function
End Class
