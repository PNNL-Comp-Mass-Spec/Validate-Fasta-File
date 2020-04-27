
/// <summary>
/// This class keeps track of a list of strings that each has an associated integer value
/// Internally it uses a dictionary to track several lists, storing each added string/integer pair to one of the lists
/// based on the first few letters of the newly added string
/// </summary>
/// <remarks></remarks>

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ValidateFastaFile
{
    public class clsNestedStringIntList
    {
        private readonly Dictionary<string, List<KeyValuePair<string, int>>> mData;
        private readonly bool mRaiseExceptionIfAddedDataNotSorted;

        private bool mDataIsSorted;

        // mSearchComparer uses StringComparison.Ordinal
        private readonly clsKeySearchComparer mSearchComparer;

        /// <summary>
        /// Number of items stored with Add()
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public int Count
        {
            get
            {
                int totalItems = 0;
                foreach (var subList in mData.Values)
                    totalItems += subList.Count;

                return totalItems;
            }
        }

        /// <summary>
        /// The number of characters at the start of stored items to use when adding items to clsNestedStringDictionary instances
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks>
        /// If this value is too short, all of the items added to the clsNestedStringDictionary instance
        /// will be tracked by the same dictionary, which could result in a dictionary surpassing the 2 GB boundary
        /// </remarks>
        public byte SpannerCharLength { get; }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="spannerCharLength"></param>
        /// <remarks>
        /// If spannerCharLength is too small, all of the items added to this class instance using Add() will be
        /// tracked by the same list, which could result in a list surpassing the 2 GB boundary
        /// </remarks>
        public clsNestedStringIntList(byte spannerCharLength = 1, bool raiseExceptionIfAddedDataNotSorted = false)
        {
            mData = new Dictionary<string, List<KeyValuePair<string, int>>>(StringComparer.InvariantCulture);

            if (spannerCharLength < 1)
            {
                SpannerCharLength = 1;
            }
            else
            {
                SpannerCharLength = spannerCharLength;
            }

            mRaiseExceptionIfAddedDataNotSorted = raiseExceptionIfAddedDataNotSorted;

            mDataIsSorted = true;

            mSearchComparer = new clsKeySearchComparer();
        }

        /// <summary>
        /// Appends an item to the list
        /// </summary>
        /// <param name="item">String to add</param>
        /// <remarks></remarks>
        public void Add(string item, int value)
        {
            string spannerKey = GetSpannerKey(item);

            if (!mData.TryGetValue(spannerKey, out var subList))
            {
                subList = new List<KeyValuePair<string, int>>();
                mData.Add(spannerKey, subList);
            }

            int lastIndexBeforeUpdate = subList.Count - 1;
            subList.Add(new KeyValuePair<string, int>(item, value));

            if (mDataIsSorted && subList.Count > 1)
            {
                // Check whether the list is still sorted
                if (string.Compare(subList[lastIndexBeforeUpdate].Key, subList[lastIndexBeforeUpdate + 1].Key, StringComparison.Ordinal) > 0)
                {
                    if (mRaiseExceptionIfAddedDataNotSorted)
                    {
                        throw new Exception("Sort issue adding " + item + " using spannerKey " + spannerKey);
                    }

                    mDataIsSorted = false;
                }
            }
        }

        /// <summary>
        /// Read a tab-delimited file, comparing the value of the text in a given column on adjacent lines
        /// to determine the appropriate spanner length when instantiating a new instance of this class
        /// </summary>
        /// <param name="fiDataFile"></param>
        /// <param name="keyColumnIndex"></param>
        /// <param name="hasHeaderLine"></param>
        /// <returns>The appropriate spanner length</returns>
        /// <remarks></remarks>
        public static byte AutoDetermineSpannerCharLength(
            FileInfo fiDataFile,
            int keyColumnIndex,
            bool hasHeaderLine)
        {
            // ReSharper disable once NotAccessedVariable
            long linesRead = 0L;

            try
            {
                var keyStartLetters = new Dictionary<string, int>();

                int previousKeyLength = 0;
                string previousKey = string.Empty;

                using (var dataReader = new StreamReader(new FileStream(fiDataFile.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)))
                {
                    if (!dataReader.EndOfStream && hasHeaderLine)
                    {
                        dataReader.ReadLine();
                    }

                    while (!dataReader.EndOfStream)
                    {
                        string dataLine = dataReader.ReadLine();
                        linesRead += 1;

                        if (string.IsNullOrEmpty(dataLine))
                        {
                            continue;
                        }

                        var dataValues = dataLine.Split('\t');
                        if (keyColumnIndex >= dataValues.Length)
                        {
                            continue;
                        }

                        string currentKey = dataValues[keyColumnIndex];

                        if (previousKeyLength == 0)
                        {
                            previousKey = string.Copy(currentKey);
                            previousKeyLength = previousKey.Length;
                            continue;
                        }

                        int currentKeyLength = currentKey.Length;
                        int charIndex = 0;

                        while (charIndex < previousKeyLength)
                        {
                            if (charIndex >= currentKeyLength)
                            {
                                break;
                            }

                            if (previousKey[charIndex] != currentKey[charIndex])
                            {
                                // Difference found; add/update the dictionary
                                break;
                            }

                            charIndex += 1;
                        }

                        int charsInCommon = charIndex;
                        if (charsInCommon > 0)
                        {
                            string baseName = previousKey.Substring(0, charsInCommon);

                            if (keyStartLetters.TryGetValue(baseName, out var matchCount))
                            {
                                keyStartLetters[baseName] = matchCount + 1;
                            }
                            else
                            {
                                keyStartLetters.Add(baseName, 1);
                            }
                        }
                    }
                }

                // Determine the appropriate spanner length given the observation counts of the base names
                byte idealSpannerCharLength = DetermineSpannerLengthUsingStartLetterStats(keyStartLetters);
                return idealSpannerCharLength;
            }
            catch (Exception ex)
            {
                throw new Exception("Error in AutoDetermineSpannerCharLength", ex);
            }
        }

        /// <summary>
        /// Determine the appropriate spanner length given the observation counts of the base names
        /// </summary>
        /// <param name="keyStartLetters">
        /// Dictionary where keys are base names (characters in common between adjacent items)
        /// and values are the observation count of each base name</param>
        /// <returns>Spanner key length that fits the majority of the entries in keyStartLetters</returns>
        /// <remarks></remarks>
        public static byte DetermineSpannerLengthUsingStartLetterStats(Dictionary<string, int> keyStartLetters)
        {
            // Compute the average observation count in keyStartLetters
            double averageCount = (from item in keyStartLetters select item.Value).Average();

            // Determine the shortest base name in proteinStartLetters for items with a count of the average or higher
            int minimumLength = (from item in keyStartLetters where item.Value >= averageCount select item.Key.Length).Min();

            byte optimalSpannerCharLength = 1;

            if (minimumLength > 0)
            {
                if (minimumLength > 100)
                {
                    optimalSpannerCharLength = 100;
                }
                else
                {
                    optimalSpannerCharLength = (byte)(minimumLength + 1);
                }

                var query = (from item in keyStartLetters where item.Key.Length == minimumLength && item.Value >= averageCount select item).ToList();

                if (query.Count > 0)
                {
                    Console.WriteLine("Shortest common prefix: " + query[0].Key + ", length " + minimumLength + ", seen " + query[0].Value.ToString("#,##0") + " times");
                }
                else
                {
                    Console.WriteLine("Shortest common prefix: ????");
                }
            }
            else
            {
                Console.WriteLine("Minimum length in keyStartLetters is 0; this is unexpected (DetermineSpannerLengthUsingStartLetterStats)");
            }

            return optimalSpannerCharLength;
        }

        /// <summary>
        /// Remove the stored items
        /// </summary>
        /// <remarks></remarks>
        public void Clear()
        {
            foreach (var item in mData)
                item.Value.Clear();

            mData.Clear();
            mDataIsSorted = true;
        }

        /// <summary>
        /// Check for the existence of a string item (ignoring the associated integer)
        /// </summary>
        /// <param name="item">String to find</param>
        /// <returns>True if the item exists, otherwise false</returns>
        /// <remarks>For large lists call Sort() prior to calling this function</remarks>
        public bool Contains(string item)
        {
            string spannerKey = GetSpannerKey(item);

            if (mData.TryGetValue(spannerKey, out var subList))
            {
                var searchItem = new KeyValuePair<string, int>(item, 0);

                if (mDataIsSorted)
                {
                    if (subList.BinarySearch(searchItem, mSearchComparer) >= 0)
                        return true;
                    return false;
                }
                else
                {
                    return subList.Contains(searchItem, mSearchComparer);
                }
            }

            return false;
        }

        /// <summary>
        /// Return a string summarizing the number of items in the List associated with each spanning key
        /// </summary>
        /// <returns>String description of the stored data</returns>
        /// <remarks>
        /// Example return strings:
        /// 1 spanning key:  'a' with 1 item
        /// 2 spanning keys: 'a' with 1 item and 'o' with 1 item
        /// 3 spanning keys: including 'a' with 1 item, 'o' with 1 item, and 'p' with 1 item
        /// 5 spanning keys: including 'a' with 2 items, 'p' with 2 items, and 'w' with 1 item
        /// </remarks>
        public string GetSizeSummary()
        {
            string summary = mData.Keys.Count + " spanning keys";

            var keyNames = mData.Keys.ToList();

            keyNames.Sort(StringComparer.Ordinal);

            if (keyNames.Count == 1)
            {
                summary = "1 spanning key:  " +
                    GetSpanningKeyDescription(keyNames[0]);
            }
            else if (keyNames.Count == 2)
            {
                summary += ": " +
                    GetSpanningKeyDescription(keyNames[0]) + " and " +
                    GetSpanningKeyDescription(keyNames[1]);
            }
            else if (keyNames.Count > 2)
            {
                int midPoint = keyNames.Count / 2;

                summary += ": including " +
                    GetSpanningKeyDescription(keyNames[0]) + ", " +
                    GetSpanningKeyDescription(keyNames[midPoint]) + ", and " +
                    GetSpanningKeyDescription(keyNames[keyNames.Count - 1]);
            }

            return summary;
        }

        private string GetSpanningKeyDescription(string keyName)
        {
            string keyDescription = "'" + keyName + "' with " + mData[keyName].Count + " item";
            if (mData[keyName].Count == 1)
            {
                return keyDescription;
            }
            else
            {
                return keyDescription + "s";
            }
        }

        /// <summary>
        /// Retrieve the List associated with the given spanner key
        /// </summary>
        /// <param name="keyName"></param>
        /// <returns>The list, or nothing if the key is not found</returns>
        /// <remarks></remarks>
        public List<KeyValuePair<string, int>> GetListForSpanningKey(string keyName)
        {
            if (mData.TryGetValue(keyName, out var subList))
            {
                return subList;
            }

            return null;
        }

        /// <summary>
        /// Retrieve the list of spanning keys in use
        /// </summary>
        /// <returns>List of keys</returns>
        /// <remarks></remarks>
        public List<string> GetSpanningKeys()
        {
            return mData.Keys.ToList();
        }

        /// <summary>
        /// Return the integer associated with the given string item
        /// </summary>
        /// <param name="item">String to find</param>
        /// <returns>Integer value if found, otherwise nothing</returns>
        /// <remarks>For large lists call Sort() prior to calling this function</remarks>
        public int GetValueForItem(string item, int valueIfNotFound = -1)
        {
            string spannerKey = GetSpannerKey(item);

            if (mData.TryGetValue(spannerKey, out var subList))
            {
                var searchItem = new KeyValuePair<string, int>(item, 0);

                if (mDataIsSorted)
                {
                    // mSearchComparer uses StringComparison.Ordinal
                    int matchIndex = subList.BinarySearch(searchItem, mSearchComparer);
                    if (matchIndex >= 0)
                    {
                        return subList[matchIndex].Value;
                    }
                }
                else
                {
                    // Use a brute-force search
                    for (int intIndex = 0; intIndex <= subList.Count - 1; intIndex++)
                    {
                        if (string.Equals(subList[intIndex].Key, item))
                        {
                            return subList[intIndex].Value;
                        }
                    }
                }
            }

            return valueIfNotFound;
        }

        /// <summary>
        /// Checks whether the items are sorted
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public bool IsSorted()
        {
            foreach (var subList in mData.Values)
            {
                for (int index = 1; index <= subList.Count - 1; index++)
                {
                    if (string.Compare(subList[index - 1].Key, subList[index].Key, StringComparison.Ordinal) > 0)
                    {
                        return false;
                    }
                }
            }

            mDataIsSorted = true;
            return default;
        }

        /// <summary>
        /// Removes all occurrences of the item from the list
        /// </summary>
        /// <param name="item">String to remove</param>
        /// <remarks></remarks>
        public void Remove(string item)
        {
            string spannerKey = GetSpannerKey(item);

            if (mData.TryGetValue(spannerKey, out var subList))
            {
                subList.RemoveAll(i => string.Equals(i.Key, item));
            }
        }

        /// <summary>
        /// Update the integer associated with the given string item
        /// </summary>
        /// <param name="item">String to find</param>
        /// <remarks>For large lists call Sort() prior to calling this function</remarks>
        /// <returns>True item was found and updated, false if the item does not exist</returns>
        public bool SetValueForItem(string item, int value)
        {
            string spannerKey = GetSpannerKey(item);

            if (mData.TryGetValue(spannerKey, out var subList))
            {
                var searchItem = new KeyValuePair<string, int>(item, 0);

                if (mDataIsSorted)
                {
                    // mSearchComparer uses StringComparison.Ordinal
                    int matchIndex = subList.BinarySearch(searchItem, mSearchComparer);
                    if (matchIndex >= 0)
                    {
                        subList[matchIndex] = new KeyValuePair<string, int>(item, value);
                        return true;
                    }
                }
                else
                {
                    // Use a brute-force search
                    int matchCount = 0;
                    for (int intIndex = 0; intIndex <= subList.Count - 1; intIndex++)
                    {
                        if (string.Equals(subList[intIndex].Key, item))
                        {
                            subList[intIndex] = new KeyValuePair<string, int>(item, value);
                            matchCount += 1;
                        }
                    }

                    if (matchCount > 0)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Sorts all of the stored items
        /// </summary>
        /// <remarks></remarks>
        public void Sort()
        {
            if (mDataIsSorted)
                return;
            foreach (var subList in mData.Values)
                // mSearchComparer uses StringComparison.Ordinal
                subList.Sort(mSearchComparer);

            mDataIsSorted = true;
        }

        private string GetSpannerKey(string key)
        {
            if (key == null)
            {
                throw new ArgumentNullException(key, "Key cannot be null");
            }

            if (key.Length <= SpannerCharLength)
            {
                return key;
            }

            return key.Substring(0, SpannerCharLength);
        }

        private class clsKeySearchComparer : IComparer<KeyValuePair<string, int>>, IEqualityComparer<KeyValuePair<string, int>>
        {
            public int Compare(KeyValuePair<string, int> x, KeyValuePair<string, int> y)
            {
                return string.Compare(x.Key, y.Key, StringComparison.Ordinal);
            }

            public bool Equals(KeyValuePair<string, int> x, KeyValuePair<string, int> y)
            {
                return string.Equals(x.Key, y.Key);
            }

            public int GetHashCode(KeyValuePair<string, int> obj)
            {
                return obj.Key.GetHashCode();
            }
        }
    }
}