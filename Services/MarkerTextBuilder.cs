using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace AudioMarkerReview.Services
{
    /// <summary>
    /// Rebuilds the "[ms]word[ms]word..." marker syntax from the ORIGINAL script
    /// text (preserving its exact whitespace/line breaks) plus a boundaries array.
    /// This mirrors align_engine/align_batch.py's build_marker_text() exactly, so
    /// files exported here paste into Word identically to files produced by the
    /// batch aligner.
    ///
    /// boundariesMs must have length == words.Count + 1:
    ///   boundariesMs[0]   = start of word 0
    ///   boundariesMs[i]   = shared boundary: end of word (i-1) AND start of word i
    ///   boundariesMs[N]   = end of the last word
    /// </summary>
    public static class MarkerTextBuilder
    {
        private static readonly Regex TokenRegex = new Regex(@"\S+|\s+", RegexOptions.Compiled);

        public static string Build(string originalText, IReadOnlyList<string> words, IReadOnlyList<int> boundariesMs)
        {
            var insertions = ComputeInsertions(originalText, words, boundariesMs);
            var sb = new StringBuilder(originalText);
            // Insert from the highest offset down so earlier offsets stay valid.
            for (int i = insertions.Count - 1; i >= 0; i--)
                sb.Insert(insertions[i].Offset, insertions[i].Marker);
            return sb.ToString();
        }

        /// <summary>
        /// Returns, for the SAME originalText, the (character offset, "[ms]" marker
        /// string) pairs that Build() would splice in — but as insertion points into
        /// the ORIGINAL string rather than a rebuilt one. DocxScriptService uses this
        /// to insert markers directly into the existing docx Text nodes so every
        /// RunProperties (bold/italic/underline/shading/font) is left untouched.
        /// </summary>
        public static List<(int Offset, string Marker)> ComputeInsertions(
            string originalText, IReadOnlyList<string> words, IReadOnlyList<int> boundariesMs)
        {
            if (words.Count == 0)
                throw new ArgumentException("No words to build markers from.");
            if (boundariesMs.Count != words.Count + 1)
                throw new ArgumentException(
                    $"boundariesMs must have {words.Count + 1} entries (words.Count + 1), got {boundariesMs.Count}.");

            var tokens = TokenRegex.Matches(originalText);
            var insertions = new List<(int Offset, string Marker)>();
            int wordIdx = 0;
            int pos = 0;

            foreach (Match m in tokens)
            {
                string tok = m.Value;
                if (string.IsNullOrWhiteSpace(tok))
                {
                    pos += tok.Length;
                    continue;
                }

                if (wordIdx >= words.Count)
                {
                    // More non-space tokens in the raw text than aligned words —
                    // text/words mismatch. Skip rather than crash here; the count
                    // check below reports it clearly once the whole pass is done.
                    pos += tok.Length;
                    continue;
                }

                if (wordIdx == 0)
                    insertions.Add((pos, $"[{boundariesMs[0]}]"));

                pos += tok.Length;
                insertions.Add((pos, $"[{boundariesMs[wordIdx + 1]}]"));
                wordIdx++;
            }

            if (wordIdx != words.Count)
                throw new InvalidOperationException(
                    $"Text/words mismatch: text has {wordIdx} whitespace-separated tokens but " +
                    $"{words.Count} words were loaded. Re-check that the source text matches the draft JSON.");

            return insertions;
        }

        /// <summary>
        /// Splits raw text into whitespace-separated words, used to validate that a
        /// loaded .txt file actually matches a loaded draft JSON's word list, and as
        /// the source of words when no draft JSON was loaded (manual-only mode).
        /// </summary>
        public static List<string> SplitWords(string originalText)
        {
            var result = new List<string>();
            foreach (Match m in TokenRegex.Matches(originalText))
            {
                if (!string.IsNullOrWhiteSpace(m.Value))
                    result.Add(m.Value);
            }
            return result;
        }
    }
}
