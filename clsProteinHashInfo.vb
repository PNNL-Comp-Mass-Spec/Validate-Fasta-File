Public Class clsProteinHashInfo

    ''' <summary>
    ''' Additional protein names
    ''' </summary>
    ''' <remarks>mProteinNameFirst is not stored here; only additional proteins</remarks>
    Private ReadOnly mAdditionalProteins As SortedSet(Of String)

    ''' <summary>
    ''' SHA-1 has of the protein sequence
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property SequenceHash As String

    ''' <summary>
    ''' Number of residues in the protein sequence
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property SequenceLength As Integer

    ''' <summary>
    ''' The first 20 residues of the protein sequence
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property SequenceStart As String

    ''' <summary>
    ''' First protein associated with this hash value
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ProteinNameFirst As String

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
        SequenceHash = seqHash

        SequenceLength = sbResidues.Length
        SequenceStart = sbResidues.ToString.Substring(0, Math.Min(sbResidues.Length, 20))

        ProteinNameFirst = proteinName
        mAdditionalProteins = New SortedSet(Of String)
        DuplicateProteinNameCount = 0
    End Sub

    Public Sub AddAdditionalProtein(proteinName As String)
        If Not mAdditionalProteins.Contains(proteinName) Then
            mAdditionalProteins.Add(proteinName)
        End If
    End Sub

    Public Overrides Function ToString() As String
        Return ProteinNameFirst & ": " & SequenceHash
    End Function
End Class
