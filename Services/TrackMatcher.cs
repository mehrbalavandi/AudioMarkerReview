using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace AudioMarkerReview.Services
{
    /// <summary>Result of matching one audio file against the qualifying tables of a .docx.</summary>
    public sealed class TrackMatchResult
    {
        /// <summary>Index into the tracks list, or -1 when there was nothing to match against.</summary>
        public int Index { get; set; } = -1;

        /// <summary>Similarity of the winner, 0..1.</summary>
        public double Score { get; set; }

        /// <summary>Similarity of the second-best candidate (0 when there is only one table).</summary>
        public double RunnerUpScore { get; set; }

        /// <summary>True when the winner is both good enough AND clearly ahead of the runner-up.</summary>
        public bool IsConfident { get; set; }

        /// <summary>Short human-readable explanation, for the status bar / warning dialog.</summary>
        public string Reason { get; set; } = "";

        /// <summary>All candidates sorted best-first — used to show the user the alternatives.</summary>
        public List<(int Index, double Score)> Ranked { get; set; } = new();
    }

    /// <summary>
    /// Decides WHICH table of a multi-track .docx belongs to the audio file that is
    /// currently loaded, by comparing the audio file name against each table's row[0]
    /// caption (the AudioTrackName that CustomLayoutGenerator uses).
    ///
    /// The comparison is deliberately forgiving about separators, casing and common
    /// abbreviations ("T1S2.mp3" == "Test 1 - Section 2") but strict about numbers,
    /// because in IELTS material the numbers are what actually tell two tracks apart:
    /// "Test 1 Section 4" must never win against "Test 11 Section 4", and a number that
    /// follows the same keyword on both sides ("section 2" vs "section 3") is treated as
    /// a hard contradiction rather than a small difference.
    /// </summary>
    public static class TrackMatcher
    {
        /// <summary>Below this similarity we don't trust the winner at all.</summary>
        public const double AcceptThreshold = 0.55;

        /// <summary>The winner must beat the runner-up by at least this much to count as confident.</summary>
        public const double AmbiguityMargin = 0.08;

        // Words that show up in captions/filenames without identifying anything.
        private static readonly HashSet<string> Noise = new(StringComparer.Ordinal)
        {
            "audio", "track", "tracks", "recording", "recordings", "script", "scripts",
            "file", "files", "final", "copy", "new", "old", "mp3", "wav", "m4a", "flac",
            "the", "of", "and", "listening"
        };

        // Common short forms, so "Sec2" and "Section 2" meet in the middle.
        private static readonly Dictionary<string, string> Synonyms = new(StringComparer.Ordinal)
        {
            ["sec"] = "section", ["sect"] = "section", ["sctn"] = "section",
            ["tst"] = "test", ["pt"] = "part", ["prt"] = "part",
            ["rec"] = "recording", ["trk"] = "track",
            ["ex"] = "exercise", ["exr"] = "exercise",
            ["ch"] = "chapter", ["chap"] = "chapter",
            ["les"] = "lesson", ["lsn"] = "lesson", ["unt"] = "unit"
        };

        // Single letters are expanded ONLY when a number follows, so "T1S2" becomes
        // "test 1 section 2" while the stray "s" of "Jack's" stays a stray "s".
        private static readonly Dictionary<string, string> ShortSynonyms = new(StringComparer.Ordinal)
        {
            ["t"] = "test", ["s"] = "section", ["p"] = "part",
            ["u"] = "unit", ["l"] = "lesson", ["c"] = "chapter"
        };

        /// <summary>
        /// Finds the table that best matches the given audio file. Never throws. When the
        /// document holds a single table that table always wins — there is nothing to
        /// confuse it with — which preserves the old "just load track 0" behaviour.
        /// </summary>
        public static TrackMatchResult Match(string audioPath, IReadOnlyList<DocxTrackInfo> tracks)
        {
            var result = new TrackMatchResult();
            if (tracks == null || tracks.Count == 0) return result;

            if (tracks.Count == 1)
            {
                result.Index = 0;
                result.Score = 1.0;
                result.IsConfident = true;
                result.Reason = "the document only has one table";
                result.Ranked.Add((0, 1.0));
                return result;
            }

            string fileName = SafeFileNameWithoutExtension(audioPath);
            string folderName = SafeFolderName(audioPath);

            var scored = new List<(int Index, double Score)>(tracks.Count);
            for (int i = 0; i < tracks.Count; i++)
            {
                string caption = tracks[i].AudioTrackName ?? "";
                double best = Similarity(fileName, caption);

                // A file called "02.mp3" inside a folder called "Test 3" carries its identity
                // in the folder name, so try the qualified name too and keep whichever wins.
                if (folderName.Length > 0)
                    best = Math.Max(best, Similarity(folderName + " " + fileName, caption));

                scored.Add((i, best));
            }

            result.Ranked = scored.OrderByDescending(s => s.Score).ThenBy(s => s.Index).ToList();
            result.Index = result.Ranked[0].Index;
            result.Score = result.Ranked[0].Score;
            result.RunnerUpScore = result.Ranked.Count > 1 ? result.Ranked[1].Score : 0.0;

            double margin = result.Score - result.RunnerUpScore;
            result.IsConfident = result.Score >= AcceptThreshold && margin >= AmbiguityMargin;
            result.Reason = result.IsConfident
                ? $"matched \"{tracks[result.Index].AudioTrackName}\""
                : result.Score < AcceptThreshold
                    ? "no table caption resembles this audio file name"
                    : "two or more tables match about equally well";
            return result;
        }

        /// <summary>Similarity of two names in 0..1. Public so it can be unit-tested / retuned.</summary>
        public static double Similarity(string a, string b)
        {
            var ta = Tokenize(Normalize(a));
            var tb = Tokenize(Normalize(b));
            if (ta.Count == 0 || tb.Count == 0) return 0.0;

            string ca = string.Concat(ta);
            string cb = string.Concat(tb);
            if (ca == cb) return 1.0;

            int maxLen = Math.Max(ca.Length, cb.Length);
            double token = TokenScore(ta, tb);
            double edit = 1.0 - (double)Levenshtein(ca, cb) / maxLen;

            double contain = 0.0;
            if (ca.Contains(cb) || cb.Contains(ca))
                contain = 0.80 + 0.20 * (Math.Min(ca.Length, cb.Length) / (double)maxLen);

            double score = Math.Max(0.70 * token + 0.30 * edit, contain);

            // ---- numbers decide -------------------------------------------------
            // "section 2" vs "section 3" is a contradiction, not a near miss; an
            // edit-distance metric alone would rate those two names as almost identical.
            var keyedA = KeyedNumbers(ta);
            var keyedB = KeyedNumbers(tb);
            int agree = 0, disagree = 0;
            foreach (var kv in keyedA)
            {
                if (!keyedB.TryGetValue(kv.Key, out string? other)) continue;
                if (string.Equals(other, kv.Value, StringComparison.Ordinal)) agree++;
                else disagree++;
            }

            var da = ta.Where(IsNumber).ToList();
            var db = tb.Where(IsNumber).ToList();
            var setA = new HashSet<string>(da, StringComparer.Ordinal);
            var setB = new HashSet<string>(db, StringComparer.Ordinal);

            if (disagree > 0)
            {
                score *= 0.30;                                   // same keyword, different number
            }
            else if (agree > 0)
            {
                score += 0.20;                                   // e.g. both say "section 2"
            }
            else if (da.Count > 0 && db.Count > 0)
            {
                if (setA.SetEquals(setB)) score += 0.15;          // same numbers, different order
                else if (setA.IsSubsetOf(setB) || setB.IsSubsetOf(setA)) { /* one side is terser */ }
                else if (!setA.Overlaps(setB)) score *= 0.35;     // nothing in common
                else score *= 0.75;                              // partial overlap
            }

            return Math.Clamp(score, 0.0, 1.0);
        }

        // ---------------------------------------------------------------

        private static bool IsNumber(string token) => token.Length > 0 && token.All(char.IsDigit);

        private static string SafeFileNameWithoutExtension(string path)
        {
            try { return Path.GetFileNameWithoutExtension(path) ?? ""; }
            catch { return path ?? ""; }
        }

        private static string SafeFolderName(string path)
        {
            try
            {
                string? dir = Path.GetDirectoryName(path);
                return string.IsNullOrEmpty(dir) ? "" : new DirectoryInfo(dir).Name;
            }
            catch { return ""; }
        }

        /// <summary>Lower-cases, turns every separator into a space, and splits
        /// letter/digit boundaries so "Test3" becomes "test 3".</summary>
        private static string Normalize(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            var sb = new StringBuilder(s.Length + 8);
            char prev = '\0';
            foreach (char raw in s)
            {
                char c = char.ToLowerInvariant(raw);
                if (char.IsLetterOrDigit(c))
                {
                    bool boundary = prev != '\0' && char.IsLetter(prev) != char.IsLetter(c);
                    if (boundary && sb.Length > 0 && sb[sb.Length - 1] != ' ') sb.Append(' ');
                    sb.Append(c);
                    prev = c;
                }
                else
                {
                    if (sb.Length > 0 && sb[sb.Length - 1] != ' ') sb.Append(' ');
                    prev = '\0';
                }
            }
            return sb.ToString().Trim();
        }

        /// <summary>Splits into tokens, expands abbreviations, strips leading zeros from
        /// numbers, and drops filler words (unless that would leave nothing at all).</summary>
        private static List<string> Tokenize(string normalized)
        {
            var raw = normalized.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            var all = new List<string>(raw.Length);
            for (int i = 0; i < raw.Length; i++)
            {
                string t = raw[i];
                bool numberFollows = i + 1 < raw.Length && IsNumber(raw[i + 1]);

                if (numberFollows && ShortSynonyms.TryGetValue(t, out string? shortSyn)) t = shortSyn;
                else if (Synonyms.TryGetValue(t, out string? syn)) t = syn;

                if (IsNumber(t))
                {
                    t = t.TrimStart('0');
                    if (t.Length == 0) t = "0";
                }
                all.Add(t);
            }
            var kept = all.Where(x => !Noise.Contains(x)).ToList();
            return kept.Count > 0 ? kept : all;
        }

        /// <summary>Maps "the word before a number" -> "that number", e.g.
        /// ["test","3","section","2"] -> { test: 3, section: 2 }.</summary>
        private static Dictionary<string, string> KeyedNumbers(List<string> tokens)
        {
            var map = new Dictionary<string, string>(StringComparer.Ordinal);
            for (int i = 1; i < tokens.Count; i++)
            {
                if (!IsNumber(tokens[i]) || IsNumber(tokens[i - 1])) continue;
                if (!map.ContainsKey(tokens[i - 1])) map[tokens[i - 1]] = tokens[i];
            }
            return map;
        }

        /// <summary>Coverage of the shorter name by the longer one, with number tokens
        /// counting double (a caption is usually wordier than a file name, so scoring
        /// against the smaller set avoids punishing "Test3.mp3" for being terse).</summary>
        private static double TokenScore(List<string> a, List<string> b)
        {
            var sa = new HashSet<string>(a, StringComparer.Ordinal);
            var sb = new HashSet<string>(b, StringComparer.Ordinal);
            if (sa.Count == 0 || sb.Count == 0) return 0.0;

            var smaller = sa.Count <= sb.Count ? sa : sb;
            double matched = 0.0, total = 0.0;
            foreach (string t in smaller)
            {
                double w = t.Any(char.IsDigit) ? 2.0 : 1.0;
                total += w;
                if (sa.Contains(t) && sb.Contains(t)) matched += w;
            }
            return total == 0.0 ? 0.0 : Math.Min(1.0, matched / total);
        }

        private static int Levenshtein(string a, string b)
        {
            if (a.Length == 0) return b.Length;
            if (b.Length == 0) return a.Length;

            var prev = new int[b.Length + 1];
            var cur = new int[b.Length + 1];
            for (int j = 0; j <= b.Length; j++) prev[j] = j;

            for (int i = 1; i <= a.Length; i++)
            {
                cur[0] = i;
                for (int j = 1; j <= b.Length; j++)
                {
                    int cost = a[i - 1] == b[j - 1] ? 0 : 1;
                    cur[j] = Math.Min(Math.Min(cur[j - 1] + 1, prev[j] + 1), prev[j - 1] + cost);
                }
                (prev, cur) = (cur, prev);
            }
            return prev[b.Length];
        }
    }
}
