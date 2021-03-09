using System;
using System.Collections.Generic;

namespace ValidateFastaFile
{
    public class ProteinHashInfo
    {
        /// <summary>
        /// Additional protein names
        /// </summary>
        /// <remarks>mProteinNameFirst is not stored here; only additional proteins</remarks>
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
        /// <param name="sbResidues"></param>
        /// <param name="proteinName"></param>
        public ProteinHashInfo(string seqHash, System.Text.StringBuilder sbResidues, string proteinName)
        {
            SequenceHash = seqHash;

            SequenceLength = sbResidues.Length;
            SequenceStart = sbResidues.ToString().Substring(0, Math.Min(sbResidues.Length, 20));

            ProteinNameFirst = proteinName;
            mAdditionalProteins = new SortedSet<string>();
            DuplicateProteinNameCount = 0;
        }

        public void AddAdditionalProtein(string proteinName)
        {
            if (!mAdditionalProteins.Contains(proteinName))
            {
                mAdditionalProteins.Add(proteinName);
            }
        }

        public override string ToString()
        {
            return ProteinNameFirst + ": " + SequenceHash;
        }
    }
}