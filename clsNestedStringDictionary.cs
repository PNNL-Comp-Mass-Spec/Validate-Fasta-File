using System;
using System.Collections.Generic;
using System.Linq;

namespace ValidateFastaFile
{

    /// <summary>
    /// This class implements a dictionary where keys are strings and values are type T (for example string or integer)
    /// Internally it uses a set of dictionaries to track the data, binning the data into separate dictionaries
    /// based on the first few letters of the keys of an added key/value pair
    /// </summary>
    /// <typeparam name="T">Type for values</typeparam>
    /// <remarks></remarks>
    public class clsNestedStringDictionary<T>
    {
        private readonly Dictionary<string, Dictionary<string, T>> mData;
        private readonly StringComparer mComparer;

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
                foreach (var subDictionary in mData.Values)
                    totalItems += subDictionary.Count;

                return totalItems;
            }
        }

        // ReSharper disable once UnusedMember.Global
        /// <summary>
        /// True when we are ignoring case for stored keys
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public bool IgnoreCase { get; }

        /// <summary>
        /// The number of characters at the start of keyStrings to use when adding items to clsNestedStringDictionary instances
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
        /// <param name="ignoreCaseForKeys">True to create case-insensitive dictionaries (and thus ignore differences between uppercase and lowercase letters)</param>
        /// <param name="spannerCharLength"></param>
        /// <remarks>
        /// If spannerCharLength is too small, all of the items added to this class instance using Add() will be
        /// tracked by the same dictionary, which could result in a dictionary surpassing the 2 GB boundary
        /// </remarks>
        public clsNestedStringDictionary(bool ignoreCaseForKeys = false, byte spannerCharLength = 1)
        {
            IgnoreCase = ignoreCaseForKeys;
            if (IgnoreCase)
            {
                mComparer = StringComparer.OrdinalIgnoreCase;
            }
            else
            {
                mComparer = StringComparer.Ordinal;
            }

            mData = new Dictionary<string, Dictionary<string, T>>(mComparer);

            if (spannerCharLength < 1)
            {
                SpannerCharLength = 1;
            }
            else
            {
                SpannerCharLength = spannerCharLength;
            }
        }

        /// <summary>
        /// Store a key and its associated value
        /// </summary>
        /// <param name="key">String to store</param>
        /// <param name="value">Value of type T</param>
        /// <remarks></remarks>
        /// <exception cref="System.ArgumentException">Thrown if the key has already been stored</exception>
        public void Add(string key, T value)
        {
            string spannerKey = GetSpannerKey(key);

            if (!mData.TryGetValue(spannerKey, out var subDictionary))
            {
                subDictionary = new Dictionary<string, T>(mComparer);
                mData.Add(spannerKey, subDictionary);
            }

            subDictionary.Add(key, value);
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
        }

        /// <summary>
        /// Check for the existence of a key
        /// </summary>
        /// <param name="key"></param>
        /// <returns>True if the key exists, otherwise false</returns>
        /// <remarks></remarks>
        public bool ContainsKey(string key)
        {
            string spannerKey = GetSpannerKey(key);

            if (mData.TryGetValue(spannerKey, out var subDictionary))
            {
                return subDictionary.ContainsKey(key);
            }

            return false;
        }

        /// <summary>
        /// Return a string summarizing the number of items in the dictionary associated with each spanning key
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

            keyNames.Sort(mComparer);

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
            string keyDescription = "'" + keyName + "' with " + mData[keyName].Values.Count + " item";
            if (mData[keyName].Values.Count == 1)
            {
                return keyDescription;
            }
            else
            {
                return keyDescription + "s";
            }
        }

        // ReSharper disable once UnusedMember.Global
        /// <summary>
        /// Retrieve the dictionary associated with the given spanner key
        /// </summary>
        /// <param name="keyName"></param>
        /// <returns>The dictionary, or nothing if the key is not found</returns>
        /// <remarks></remarks>
        public Dictionary<string, T> GetDictionaryForSpanningKey(string keyName)
        {
            if (mData.TryGetValue(keyName, out var subDictionary))
            {
                return subDictionary;
            }

            return null;
        }

        // ReSharper disable once UnusedMember.Global
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
        /// Try to get the value associated with the key
        /// </summary>
        /// <param name="key">Key to find</param>
        /// <param name="value">Value found, or nothing if no match</param>
        /// <returns>True if a match was found, otherwise nothing</returns>
        /// <remarks></remarks>
        public bool TryGetValue(string key, out T value)
        {
            string spannerKey = GetSpannerKey(key);

            if (mData.TryGetValue(spannerKey, out var subDictionary))
            {
                return subDictionary.TryGetValue(key, out value);
            }

            value = default;
            return false;
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
    }
}