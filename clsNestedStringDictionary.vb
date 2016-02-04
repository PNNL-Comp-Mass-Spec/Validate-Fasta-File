Imports System.Runtime.InteropServices

''' <summary>
''' This class implements of dictionary where keys are strings and values are type T (for example string or integer)
''' Internally it uses a set of dictionaries to track the data, binning the data into separate dictionaries
''' based on the first few letters of the keys of an added key/value pair
''' </summary>
''' <typeparam name="T">Type for values</typeparam>
''' <remarks></remarks>
Public Class clsNestedStringDictionary(Of T)

    Private ReadOnly mData As Dictionary(Of String, Dictionary(Of String, T))
    Private ReadOnly mSpannerCharLength As Byte
    Private ReadOnly mComparer As StringComparer
    Private ReadOnly mIgnoreCase As Boolean

    ''' <summary>
    ''' Number of items stored with Add()
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Count As Integer
        Get
            Dim totalItems = 0
            For Each subDictionary In mData.Values
                totalItems += subDictionary.Count
            Next
            Return totalItems
        End Get
    End Property

    ''' <summary>
    ''' True when we are ignoring case for stored keys
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property IgnoreCase As Boolean
        Get
            Return mIgnoreCase
        End Get
    End Property

    ''' <summary>
    ''' The number of characters at the start of keystrings to use when adding items to clsNestedStringDictionary instances
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
    ''' <param name="ignoreCaseForKeys">True to create case-insensitve dictionaries (and thus ignore differences between uppercase and lowercase letters)</param>
    ''' <param name="spannerCharLength"></param>
    ''' <remarks>
    ''' If spannerCharLength is too small, all of the items added to this class instance using Add() will be
    ''' tracked by the same dictionary, which could result in a dictionary surpassing the 2 GB boundary</remarks>
    ''' </remarks>
    Public Sub New(Optional ignoreCaseForKeys As Boolean = False, Optional spannerCharLength As Byte = 1)

        mIgnoreCase = ignoreCaseForKeys
        If mIgnoreCase Then
            mComparer = StringComparer.InvariantCultureIgnoreCase
        Else
            mComparer = StringComparer.InvariantCulture
        End If

        mData = New Dictionary(Of String, Dictionary(Of String, T))(mComparer)

        If spannerCharLength < 1 Then
            mSpannerCharLength = 1
        Else
            mSpannerCharLength = spannerCharLength
        End If

    End Sub

    ''' <summary>
    ''' Store a key and its associated value
    ''' </summary>
    ''' <param name="key">String to store</param>
    ''' <param name="value">Value of type T</param>
    ''' <remarks></remarks>
    ''' <exception cref="System.ArgumentException">Thrown if the key has already been stored</exception>
    Public Sub Add(key As String, value As T)

        Dim subDictionary As Dictionary(Of String, T) = Nothing
        Dim spannerKey = GetSpannerKey(key)

        If Not (mData.TryGetValue(spannerKey, subDictionary)) Then
            subDictionary = New Dictionary(Of String, T)(mComparer)
            mData.Add(spannerKey, subDictionary)
        End If

        subDictionary.Add(key, value)

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
    End Sub

    ''' <summary>
    ''' Check for the existence of a key
    ''' </summary>
    ''' <param name="key"></param>
    ''' <returns>True if the key exists, otherwise false</returns>
    ''' <remarks></remarks>
    Public Function ContainsKey(key As String) As Boolean

        Dim subDictionary As Dictionary(Of String, T) = Nothing
        Dim spannerKey = GetSpannerKey(key)

        If (mData.TryGetValue(spannerKey, subDictionary)) Then
            Return subDictionary.ContainsKey(key)
        End If

        Return False

    End Function

    ''' <summary>
    ''' Return a string summarizing the number of items in the dictionary associated with each spanning key
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

        Dim keyDescription = "'" & keyName & "' with " & mData(keyName).Values.Count & " item"
        If mData(keyName).Values.Count = 1 Then
            Return keyDescription
        Else
            Return keyDescription & "s"
        End If

    End Function

    ''' <summary>
    ''' Retrieve the dictionary associated with the given spanner key
    ''' </summary>
    ''' <param name="keyName"></param>
    ''' <returns>The dictionary, or nothing if the key is not found</returns>
    ''' <remarks></remarks>
    Public Function GetDictionaryForSpanningKey(keyName As String) As Dictionary(Of String, T)
        Dim subDictionary As Dictionary(Of String, T) = Nothing
        If mData.TryGetValue(keyName, subDictionary) Then
            Return subDictionary
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
    ''' Try to get the value associated with the key
    ''' </summary>
    ''' <param name="key">Key to find</param>
    ''' <param name="value">Value found, or nothing if no match</param>
    ''' <returns>True if a match was found, otherwise nothing</returns>
    ''' <remarks></remarks>
    Public Function TryGetValue(key As String, <Out> ByRef value As T) As Boolean

        Dim subDictionary As Dictionary(Of String, T) = Nothing
        Dim spannerKey = GetSpannerKey(key)

        If (mData.TryGetValue(spannerKey, subDictionary)) Then
            Return subDictionary.TryGetValue(key, value)
        End If

        value = Nothing
        Return False

    End Function

    Private Function GetSpannerKey(key As String) As String
        If key Is Nothing Then
            Throw New ArgumentNullException(key, "Key cannot be null")
        End If

        If key.Length <= mSpannerCharLength Then
            Return key
        End If

        Return key.Substring(0, mSpannerCharLength)
    End Function

End Class
