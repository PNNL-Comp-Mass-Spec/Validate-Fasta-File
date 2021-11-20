using System;
using System.Collections.Generic;

namespace ValidateFastaFile
{
    /// <summary>
    /// Container for sequence hashes and associated protein names
    /// </summary>
    public class ProteinHashInfo
    {
        /// <summary>
        /// Additional protein names
        /// </summary>
        /// <remarks>ProteinNameFirst is not stored here; only additional proteins</remarks>
        private readonly SortedSet<string> mAdditionalProteins;

        /// <summary>
        /// SHA-1 has of the protein sequence
        /// </summary>
        public string SequenceHash { get; }

        /// <summary>
        /// Number of residues in the protein sequence
        /// </summary>
        public int SequenceLength { get; }

        /// <summary>
        /// The first 20 residues of the protein sequence
        /// </summary>
        public string SequenceStart { get; }

        /// <summary>
        /// First protein associated with this hash value
        /// </summary>
        public string ProteinNameFirst { get; }

        /// <summary>
        /// Additional protein names for this sequence
        /// </summary>
        public IEnumerable<string> AdditionalProteins => mAdditionalProteins;

        /// <summary>
        /// Greater than 0 if multiple entries have the same name and same sequence
        /// </summary>
        public int DuplicateProteinNameCount { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="seqHash"></param>
        /// <param name="residues"></param>
        /// <param name="proteinName"></param>
        public ProteinHashInfo(string seqHash, System.Text.StringBuilder residues, string proteinName)
        {
            SequenceHash = seqHash;

            SequenceLength = residues.Length;
            SequenceStart = residues.ToString().Substring(0, Math.Min(residues.Length, 20));

            ProteinNameFirst = proteinName;
            mAdditionalProteins = new SortedSet<string>();
            DuplicateProteinNameCount = 0;
        }

        /// <summary>
        /// Add an additional protein name
        /// </summary>
        /// <param name="proteinName"></param>
        public void AddAdditionalProtein(string proteinName)
        {
            if (!mAdditionalProteins.Contains(proteinName))
            {
                mAdditionalProteins.Add(proteinName);
            }
        }

        /// <summary>
        /// Show the protein name and sequence hash
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return ProteinNameFirst + ": " + SequenceHash;
        }
    }
}