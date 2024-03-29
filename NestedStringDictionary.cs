﻿using System;
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
    public class NestedStringDictionary<T>
    {
        private readonly Dictionary<string, Dictionary<string, T>> mData;
        private readonly StringComparer mComparer;

        /// <summary>
        /// Number of items stored with Add()
        /// </summary>
        public int Count
        {
            get
            {
                return mData.Values.Sum(subDictionary => subDictionary.Count);
            }
        }

        // ReSharper disable once UnusedMember.Global
        /// <summary>
        /// True when we are ignoring case for stored keys
        /// </summary>
        public bool IgnoreCase { get; }

        /// <summary>
        /// The number of characters at the start of keyStrings to use when adding items to NestedStringDictionary instances
        /// </summary>
        /// <remarks>
        /// If this value is too short, all of the items added to the NestedStringDictionary instance
        /// will be tracked by the same dictionary, which could result in a dictionary surpassing the 2 GB boundary
        /// </remarks>
        public byte SpannerCharLength { get; }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <remarks>
        /// If spannerCharLength is too small, all of the items added to this class instance using Add() will be
        /// tracked by the same dictionary, which could result in a dictionary surpassing the 2 GB boundary
        /// </remarks>
        /// <param name="ignoreCaseForKeys">True to create case-insensitive dictionaries (and thus ignore differences between uppercase and lowercase letters)</param>
        /// <param name="spannerCharLength"></param>
        public NestedStringDictionary(bool ignoreCaseForKeys = false, byte spannerCharLength = 1)
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
        /// <exception cref="System.ArgumentException">Thrown if the key has already been stored</exception>
        public void Add(string key, T value)
        {
            var spannerKey = GetSpannerKey(key);

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
        public void Clear()
        {
            foreach (var item in mData)
            {
                item.Value.Clear();
            }

            mData.Clear();
        }

        /// <summary>
        /// Check for the existence of a key
        /// </summary>
        /// <param name="key"></param>
        /// <returns>True if the key exists, otherwise false</returns>
        public bool ContainsKey(string key)
        {
            var spannerKey = GetSpannerKey(key);

            if (mData.TryGetValue(spannerKey, out var subDictionary))
            {
                return subDictionary.ContainsKey(key);
            }

            return false;
        }

        /// <summary>
        /// Return a string summarizing the number of items in the dictionary associated with each spanning key
        /// </summary>
        /// <remarks>
        /// Example return strings:
        /// 1 spanning key:  'a' with 1 item
        /// 2 spanning keys: 'a' with 1 item and 'o' with 1 item
        /// 3 spanning keys: including 'a' with 1 item, 'o' with 1 item, and 'p' with 1 item
        /// 5 spanning keys: including 'a' with 2 items, 'p' with 2 items, and 'w' with 1 item
        /// </remarks>
        /// <returns>String description of the stored data</returns>
        public string GetSizeSummary()
        {
            var summary = mData.Keys.Count + " spanning keys";

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
                var midPoint = keyNames.Count / 2;

                summary += ": including " +
                    GetSpanningKeyDescription(keyNames[0]) + ", " +
                    GetSpanningKeyDescription(keyNames[midPoint]) + ", and " +
                    GetSpanningKeyDescription(keyNames[keyNames.Count - 1]);
            }

            return summary;
        }

        private string GetSpanningKeyDescription(string keyName)
        {
            var keyDescription = "'" + keyName + "' with " + mData[keyName].Values.Count + " item";
            if (mData[keyName].Values.Count == 1)
            {
                return keyDescription;
            }

            return keyDescription + "s";
        }

        // ReSharper disable once UnusedMember.Global
        /// <summary>
        /// Retrieve the dictionary associated with the given spanner key
        /// </summary>
        /// <param name="keyName"></param>
        /// <returns>The dictionary, or nothing if the key is not found</returns>
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
        public bool TryGetValue(string key, out T value)
        {
            var spannerKey = GetSpannerKey(key);

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
                throw new ArgumentNullException(nameof(key), "Key cannot be null");
            }

            if (key.Length <= SpannerCharLength)
            {
                return key;
            }

            return key.Substring(0, SpannerCharLength);
        }
    }
}
