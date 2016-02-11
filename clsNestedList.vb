''' <summary>
''' This class keeps track of a list of strings that each has an associated integer value
''' Internally it uses a dictionary to track several lists, storing each added string/integer pair to one of the lists
''' based on the first few letters of the newly added string
''' </summary>
''' <remarks></remarks>

Public Class clsNestedStringIntList

    Private ReadOnly mData As Dictionary(Of String, List(Of KeyValuePair(Of String, Integer)))
    Private ReadOnly mSpannerCharLength As Byte

    Private mDataIsSorted As Boolean

    Private ReadOnly mSearchComparer As clsKeySearchComparer

    ''' <summary>
    ''' Number of items stored with Add()
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Count As Integer
        Get
            Dim totalItems = 0
            For Each subList In mData.Values
                totalItems += subList.Count
            Next
            Return totalItems
        End Get
    End Property

    ''' <summary>
    ''' The number of characters at the start of stored items to use when adding items to clsNestedStringDictionary instances
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>
    ''' If this value is too short, all of the items added to the clsNestedStringDictionary instance 
    ''' will be tracked by the same dictionary, which could result in a dictionary surpassing the 2 GB boundary
    ''' </remarks>
    Public ReadOnly Property SpannerCharLength As Byte
        Get
            Return mSpannerCharLength
        End Get
    End Property

    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <param name="spannerCharLength"></param>
    ''' <remarks>
    ''' If spannerCharLength is too small, all of the items added to this class instance using Add() will be
    ''' tracked by the same list, which could result in a list surpassing the 2 GB boundary
    ''' </remarks>
    Public Sub New(Optional spannerCharLength As Byte = 1)

        mData = New Dictionary(Of String, List(Of KeyValuePair(Of String, Integer)))(StringComparer.InvariantCulture)

        If spannerCharLength < 1 Then
            mSpannerCharLength = 1
        Else
            mSpannerCharLength = spannerCharLength
        End If

        mDataIsSorted = True

        mSearchComparer = New clsKeySearchComparer()
    End Sub

    ''' <summary>
    ''' Appends an item to the list
    ''' </summary>
    ''' <param name="item">String to add</param>
    ''' <remarks></remarks>
    Public Sub Add(item As String, value As Integer)

        Dim subList As List(Of KeyValuePair(Of String, Integer)) = Nothing
        Dim spannerKey = GetSpannerKey(item)

        If Not (mData.TryGetValue(spannerKey, subList)) Then
            subList = New List(Of KeyValuePair(Of String, Integer))
            mData.Add(spannerKey, subList)
        End If

        Dim lastIndexBeforeUpdate = subList.Count - 1
        subList.Add(New KeyValuePair(Of String, Integer)(item, value))

        If mDataIsSorted AndAlso subList.Count > 1 Then
            ' Check whether the list is still sorted
            If String.Compare(subList(lastIndexBeforeUpdate).Key, subList(lastIndexBeforeUpdate + 1).Key) > 0 Then
                mDataIsSorted = False
            End If
        End If

    End Sub

    ''' <summary>
    ''' Remove the stored items
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Clear()
        For Each item In mData
            item.Value.Clear()
        Next

        mData.Clear()
        mDataIsSorted = True
    End Sub

    ''' <summary>
    ''' Check for the existence of a string item (ignoring the associated integer)
    ''' </summary>
    ''' <param name="item">String to find</param>
    ''' <returns>True if the item exists, otherwise false</returns>
    ''' <remarks>For large lists call Sort() prior to calling this function</remarks>
    Public Function Contains(item As String) As Boolean

        Dim subList As List(Of KeyValuePair(Of String, Integer)) = Nothing
        Dim spannerKey = GetSpannerKey(item)

        If (mData.TryGetValue(spannerKey, subList)) Then
            Dim searchItem = New KeyValuePair(Of String, Integer)(item, 0)

            If mDataIsSorted Then
                If subList.BinarySearch(searchItem, mSearchComparer) >= 0 Then Return True
                Return False
            Else
                Return subList.Contains(searchItem, mSearchComparer)
            End If

        End If

        Return False

    End Function

    ''' <summary>
    ''' Return a string summarizing the number of items in the List associated with each spanning key
    ''' </summary>
    ''' <returns>String description of the stored data</returns>
    ''' <remarks>
    ''' Example return strings:
    '''   1 spanning key:  'a' with 1 item
    '''   2 spanning keys: 'a' with 1 item and 'o' with 1 item
    '''   3 spanning keys: including 'a' with 1 item, 'o' with 1 item, and 'p' with 1 item
    '''   5 spanning keys: including 'a' with 2 items, 'p' with 2 items, and 'w' with 1 item
    ''' </remarks>
    Public Function GetSizeSummary() As String

        Dim summary = mData.Keys.Count & " spanning keys"

        Dim keyNames = mData.Keys.ToList()

        keyNames.Sort()

        If keyNames.Count = 1 Then
            summary = "1 spanning key:  " &
                GetSpanningKeyDescription(keyNames(0))

        ElseIf keyNames.Count = 2 Then
            summary &= ": " &
                GetSpanningKeyDescription(keyNames(0)) & " and " &
                GetSpanningKeyDescription(keyNames(1))

        ElseIf keyNames.Count > 2 Then
            Dim midPoint = CInt(Math.Floor(keyNames.Count / 2))

            summary &= ": including " &
                GetSpanningKeyDescription(keyNames(0)) & ", " &
                GetSpanningKeyDescription(keyNames(midPoint)) & ", and " &
                GetSpanningKeyDescription(keyNames(keyNames.Count - 1))

        End If

        Return summary

    End Function

    Private Function GetSpanningKeyDescription(keyName As String) As String

        Dim keyDescription = "'" & keyName & "' with " & mData(keyName).Count & " item"
        If mData(keyName).Count = 1 Then
            Return keyDescription
        Else
            Return keyDescription & "s"
        End If

    End Function

    ''' <summary>
    ''' Retrieve the List associated with the given spanner key
    ''' </summary>
    ''' <param name="keyName"></param>
    ''' <returns>The list, or nothing if the key is not found</returns>
    ''' <remarks></remarks>
    Public Function GetListForSpanningKey(keyName As String) As List(Of KeyValuePair(Of String, Integer))
        Dim subList As List(Of KeyValuePair(Of String, Integer)) = Nothing
        If mData.TryGetValue(keyName, subList) Then
            Return subList
        End If

        Return Nothing
    End Function

    ''' <summary>
    ''' Retrieve the list of spanning keys in use
    ''' </summary>
    ''' <returns>List of keys</returns>
    ''' <remarks></remarks>
    Public Function GetSpanningKeys() As List(Of String)
        Return mData.Keys.ToList()
    End Function

    ''' <summary>
    ''' Return the integer associated with the given string item
    ''' </summary>
    ''' <param name="item">String to find</param>
    ''' <returns>Integer value if found, otherwise nothing</returns>
    ''' <remarks>For large lists call Sort() prior to calling this function</remarks>
    Public Function GetValueForItem(item As String, Optional valueIfNotFound As Integer = -1) As Integer

        Dim subList As List(Of KeyValuePair(Of String, Integer)) = Nothing
        Dim spannerKey = GetSpannerKey(item)

        If (mData.TryGetValue(spannerKey, subList)) Then
            Dim searchItem = New KeyValuePair(Of String, Integer)(item, 0)

            If mDataIsSorted Then
                Dim matchIndex = subList.BinarySearch(searchItem, mSearchComparer)
                If matchIndex >= 0 Then
                    Return subList(matchIndex).Value
                End If

            Else
                For intIndex = 0 To subList.Count - 1
                    If String.Equals(subList(intIndex).Key, item) Then
                        Return subList(intIndex).Value
                    End If
                Next
            End If

        End If

        Return valueIfNotFound

    End Function

    ''' <summary>
    ''' Checks whether the items are sorted
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsSorted() As Boolean
        For Each subList In mData.Values
            For index = 1 To subList.Count - 1
                If String.Compare(subList(index - 1).Key, subList(index).Key) > 0 Then
                    Return False
                End If
            Next
        Next
        mDataIsSorted = True
    End Function

    ''' <summary>
    ''' Removes all occurrences of the item from the list
    ''' </summary>
    ''' <param name="item">String to remove</param>
    ''' <remarks></remarks>
    Public Sub Remove(item As String)

        Dim subList As List(Of KeyValuePair(Of String, Integer)) = Nothing
        Dim spannerKey = GetSpannerKey(item)

        If (mData.TryGetValue(spannerKey, subList)) Then
            subList.RemoveAll(Function(i) String.Equals(i.Key, item))
        End If

    End Sub
    
    ''' <summary>
    ''' Update the integer associated with the given string item
    ''' </summary>
    ''' <param name="item">String to find</param>
    ''' <remarks>For large lists call Sort() prior to calling this function</remarks>
    ''' <returns>True item was found and updated, false if the item does not exist</returns>
    Public Function SetValueForItem(item As String, value As Integer) As Boolean

        Dim subList As List(Of KeyValuePair(Of String, Integer)) = Nothing
        Dim spannerKey = GetSpannerKey(item)

        If (mData.TryGetValue(spannerKey, subList)) Then
            Dim searchItem = New KeyValuePair(Of String, Integer)(item, 0)

            If mDataIsSorted Then
                Dim matchIndex = subList.BinarySearch(searchItem, mSearchComparer)
                If matchIndex >= 0 Then
                    subList(matchIndex) = New KeyValuePair(Of String, Integer)(item, value)
                    Return True
                End If
            Else
                Dim matchCount = 0
                For intIndex = 0 To subList.Count - 1
                    If String.Equals(subList(intIndex).Key, item) Then
                        subList(intIndex) = New KeyValuePair(Of String, Integer)(item, value)
                        matchCount += 1
                    End If
                Next

                If matchCount > 0 Then
                    Return True
                End If
            End If

        End If

        Return False
    End Function

    ''' <summary>
    ''' Sorts all of the stored items
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Sort()
        If mDataIsSorted Then Return

        For Each subList In mData.Values
            subList.Sort(mSearchComparer)
        Next
        mDataIsSorted = True
    End Sub

    Private Function GetSpannerKey(key As String) As String
        If key Is Nothing Then
            Throw New ArgumentNullException(key, "Key cannot be null")
        End If

        If key.Length <= mSpannerCharLength Then
            Return key
        End If

        Return key.Substring(0, mSpannerCharLength)
    End Function

    Private Class clsKeySearchComparer
        Implements IComparer(Of KeyValuePair(Of String, Integer)), IEqualityComparer(Of KeyValuePair(Of String, Integer))

        Public Function Compare(x As KeyValuePair(Of String, Integer), y As KeyValuePair(Of String, Integer)) As Integer Implements IComparer(Of KeyValuePair(Of String, Integer)).Compare
            If x.Key < y.Key Then
                Return -1
            ElseIf x.Key > y.Key Then
                Return 1
            Else
                Return 0
            End If

        End Function

        Public Shadows Function Equals(x As KeyValuePair(Of String, Integer), y As KeyValuePair(Of String, Integer)) As Boolean Implements IEqualityComparer(Of KeyValuePair(Of String, Integer)).Equals
            Return String.Equals(x.Key, y.Key)
        End Function

        Public Shadows Function GetHashCode(obj As KeyValuePair(Of String, Integer)) As Integer Implements IEqualityComparer(Of KeyValuePair(Of String, Integer)).GetHashCode
            Return obj.Key.GetHashCode()
        End Function
    End Class
    
End Class
