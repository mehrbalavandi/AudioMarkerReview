using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AudioMarkerReview.Services
{
    /// <summary>One qualifying 2-row audio-script table found in a .docx.</summary>
    public class DocxTrackInfo
    {
        public int TrackIndex { get; set; }       // position among qualifying tables, in document order
        public string AudioTrackName { get; set; } = "";
        public string FullText { get; set; } = ""; // row[1]'s flattened text — same words/order MarkerTextBuilder expects
        public int WordCount { get; set; }        // words in FullText, shown in the track picker

        /// <summary>
        /// 🌟 مارکرهایی که از قبل در همین جدول نوشته شده‌اند (ترکی که قبلاً بازبینی و با «To Word»
        /// ذخیره شده). فقط وقتی پُر است که دقیقاً یک مارکر سرِ هر مرز باشد (WordCount + 1 تا)؛
        /// وگرنه null و ترک مثلِ ترکِ بی‌مارکر بار می‌شود. FullText همیشه بدونِ مارکر است.
        /// </summary>
        public List<int>? ExistingBoundariesMs { get; set; }

        /// <summary>تعدادِ مارکرهای [ms] که در جدول پیدا شد (برای نمایش، حتی اگر ناقص باشند).</summary>
        public int ExistingMarkerCount { get; set; }

        /// <summary>
        /// 🌟 وضعیتِ همین ترک در سندِ مقصد (_markers.docx) — MainWindow پُرش می‌کند. وقتی خودِ سندِ
        /// _markers باز شده، همان ExistingBoundariesMs است.
        /// </summary>
        public List<int>? ReviewedBoundariesMs { get; set; }
        public bool PartlyMarkedInTarget { get; set; }

        /// <summary>Label shown in the Track combo box.</summary>
        public string Display =>
            $"{TrackIndex + 1}. {AudioTrackName}" + (WordCount > 0 ? $"  ({WordCount} words)" : "") +
            (ReviewedBoundariesMs != null ? "  ✓ marked" : PartlyMarkedInTarget ? "  ⚠ partly marked" : "");
    }

    /// <summary>
    /// Reads/writes the exact table shape CustomLayoutGenerator.ProcessAudioScriptWordFile
    /// expects: any Descendants&lt;Table&gt; with &gt;= 2 rows and a non-empty row[0] (the
    /// AudioTrackName). Row[1]'s cells/paragraphs/runs hold the script text — on export we
    /// splice "[ms]" marker strings directly into the EXISTING Run/Text nodes' text and leave
    /// every RunProperties (bold/italic/underline/shading/font/size) completely untouched, so
    /// CustomLayoutGenerator's own style extraction (ExtractRunMarkers) re-derives styling from
    /// the edited document exactly as it would have from the un-marked one.
    /// </summary>
    public static class DocxScriptService
    {
        /// <summary>
        /// 🌟 مارکرهای زمانی که CustomLayoutGenerator می‌شناسد (MainForm.cs، markerPattern):
        /// [1234] (میلی‌ثانیه، همان که این برنامه می‌نویسد) و [mm:ss.fff] / [h:mm:ss]. پیش از
        /// شمردنِ کلمه‌ها و پیش از نوشتنِ مارکرهای تازه از متنِ جدول برداشته می‌شوند.
        /// </summary>
        private static readonly Regex AnyMarkerRegex = new Regex(
            @"\[(?:(?:\d{1,2}):)?\d{1,2}:\d{2}(?:\.\d+)?\]|\[\d+\]", RegexOptions.Compiled);
        private static readonly Regex MsMarkerRegex = new Regex(@"\[(\d+)\]", RegexOptions.Compiled);
        private static readonly Regex TokenRegex = new Regex(@"\S+", RegexOptions.Compiled);

        /// <summary>
        /// یک «[3]» تنها ممکن است بخشی از خودِ متن باشد (شماره‌ی بند، ارجاع)، پس متنی فقط وقتی
        /// «مارک‌خورده» حساب می‌شود که دست‌کم دو مارکر داشته باشد.
        /// </summary>
        private const int MinMarkersToTreatAsMarked = 2;

        /// <summary>
        /// 🌟 مثلِ <see cref="ExtractTracks"/> ولی از روی یک کپیِ موقت، تا سندی که در Word باز است هم
        /// خوانده شود (Word فایل را قفل می‌کند).
        /// </summary>
        public static List<DocxTrackInfo> ExtractTracksShared(string path)
        {
            string temp = Path.Combine(Path.GetTempPath(), "amr_" + Guid.NewGuid().ToString("N") + ".docx");
            using (var src = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var dst = File.Create(temp))
                src.CopyTo(dst);
            try { return ExtractTracks(temp); }
            finally { try { File.Delete(temp); } catch { /* موقت */ } }
        }

        /// <summary>Read-only pass: lists every qualifying track with its plain text, for
        /// picking a track and for word-matching against a draft JSON / the aligner's input.
        /// Opens and closes the file immediately — never holds a lock on it.</summary>
        public static List<DocxTrackInfo> ExtractTracks(string path)
        {
            using var doc = WordprocessingDocument.Open(path, false);
            var body = doc.MainDocumentPart?.Document?.Body
                       ?? throw new InvalidDataException("This .docx has no document body.");

            var tracks = new List<DocxTrackInfo>();
            int index = 0;
            foreach (var table in FindQualifyingTables(body))
            {
                var rows = table.Elements<TableRow>().ToList();
                string trackName = rows[0].InnerText.Trim();
                string rawText = FlattenRow1Text(rows[1], out _);
                var (fullText, existing, markerCount) = ParseExistingMarkers(rawText);
                tracks.Add(new DocxTrackInfo
                {
                    TrackIndex = index,
                    AudioTrackName = trackName,
                    FullText = fullText,
                    WordCount = MarkerTextBuilder.SplitWords(fullText).Count,
                    ExistingBoundariesMs = existing,
                    ExistingMarkerCount = markerCount,
                });
                index++;
            }
            return tracks;
        }

        /// <summary>
        /// 🌟 مارکرهای یک ترک را در سندِ مقصد می‌نویسد و <b>بقیه‌ی جدول‌ها را دست نمی‌زند</b>.
        ///
        /// - اگر <paramref name="targetPath"/> هست (مثلاً «Mindset1- Audio Scripts_markers.docx» که
        ///   ترک‌های دیگرش قبلاً مارک خورده‌اند)، همان فایل به‌روز می‌شود: فقط جدولِ همین ترک عوض
        ///   می‌شود و مارکرهای قبلیِ همین جدول جایشان را به مارکرهای تازه می‌دهند.
        /// - اگر نیست، اول از <paramref name="sourcePath"/> کپی می‌شود.
        /// - مقصد می‌تواند خودِ سندِ باز‌شده باشد (وقتی خودِ فایلِ _markers را باز کرده‌اید).
        ///
        /// امن: روی یک کپیِ موقت در همان پوشه کار می‌کند و آخر با File.Replace جایگزین می‌کند؛ نسخه‌ی
        /// قبلی در «‎&lt;نام&gt;.bak» می‌ماند. اگر سند در Word باز باشد، فایلِ اصلی دست‌نخورده می‌ماند و
        /// <see cref="DocumentLockedException"/> پرتاب می‌شود.
        ///
        /// جدولِ ترک در مقصد با نامش (row[0]) پیدا می‌شود، نه فقط با شماره، تا اگر ترتیبِ جدول‌های
        /// مقصد با سندِ منبع فرق داشت، مارکرها در جدولِ اشتباه نروند.
        /// </summary>
        /// <returns>شماره‌ی جدول در مقصد (از صفر) و اینکه فایلِ مقصد تازه ساخته شد یا نه.</returns>
        public static (int TargetTrackIndex, bool Created) WriteMarkers(string sourcePath, string targetPath,
            int trackIndex, string expectedTrackName, IReadOnlyList<string> words, IReadOnlyList<int> boundariesMs)
        {
            string fullTarget = Path.GetFullPath(targetPath);
            string dir = Path.GetDirectoryName(fullTarget) ?? ".";
            bool created = !File.Exists(fullTarget);
            string seed = created ? sourcePath : fullTarget;

            // کپیِ موقت کنارِ مقصد (همان درایو، تا File.Replace اتمی بماند)
            string temp = Path.Combine(dir, "~amr_" + Guid.NewGuid().ToString("N") + ".docx");
            try
            {
                try
                {
                    using var src = new FileStream(seed, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
                    using var dst = File.Create(temp);
                    src.CopyTo(dst);
                }
                catch (IOException ex) { throw new DocumentLockedException(seed, ex); }

                int targetIndex;
                using (var doc = WordprocessingDocument.Open(temp, true))
                {
                    var body = doc.MainDocumentPart?.Document?.Body
                               ?? throw new InvalidDataException("This .docx has no document body.");
                    var qualifying = FindQualifyingTables(body).ToList();
                    targetIndex = ResolveTargetTable(qualifying, trackIndex, expectedTrackName, fullTarget);
                    var rows = qualifying[targetIndex].Elements<TableRow>().ToList();

                    // ۱) مارکرهای قبلیِ همین جدول برداشته می‌شوند (ترکی که دوباره بازبینی شده).
                    //    روی متنِ کلِ ردیف پیدا می‌شوند، چون Word گاهی یک «[1234]» را میانِ دو Run می‌شکند.
                    string rawBefore = FlattenRow1Text(rows[1], out var runsBefore);
                    if (CountMsMarkers(rawBefore) >= MinMarkersToTreatAsMarked || HasClockMarkers(rawBefore))
                        RemoveRanges(runsBefore, AnyMarkerRegex.Matches(rawBefore).Select(m => (m.Index, m.Length)).ToList());

                    // ۲) متنِ تمیز باید همان کلمه‌هایی را داشته باشد که بازبینی شده‌اند.
                    string fullText = FlattenRow1Text(rows[1], out var textRuns);
                    var literal = AnyMarkerRegex.Match(fullText);
                    if (literal.Success)
                        throw new InvalidOperationException(
                            $"Table \"{expectedTrackName}\" contains \"{literal.Value}\" as part of its text. CustomLayoutGenerator " +
                            "reads that as a time marker, so the timings would be wrong. Change it in Word (e.g. \"(3)\") and try again.");
                    var freshWords = MarkerTextBuilder.SplitWords(fullText);
                    if (freshWords.Count != words.Count)
                        throw new InvalidOperationException(
                            $"Table \"{expectedTrackName}\" in {Path.GetFileName(fullTarget)} has {freshWords.Count} words " +
                            $"but you reviewed {words.Count}. The text differs between the two documents — " +
                            "open the _markers document itself in the review tool and review this track again.");

                    // ۳) درجِ مارکرهای تازه داخلِ همان Text nodeها (استایل‌ها دست نمی‌خورند).
                    var insertions = MarkerTextBuilder.ComputeInsertions(fullText, words, boundariesMs);
                    foreach (var (offset, marker) in insertions.OrderByDescending(i => i.Offset))
                    {
                        var target = FindInsertionTarget(textRuns, offset);
                        int localOffset = Math.Max(0, Math.Min(target.Length, offset - target.StartOffset));
                        string current = target.Element.Text ?? "";
                        target.Element.Text = current.Insert(localOffset, marker);
                        // guard against Word/OpenXml trimming the whitespace next to the markers on save
                        target.Element.Space = new EnumValue<SpaceProcessingModeValues>(SpaceProcessingModeValues.Preserve);
                    }

                    doc.MainDocumentPart!.Document.Save();
                }

                // ۴) جایگزینیِ اتمی؛ نسخه‌ی قبلی در .bak
                try
                {
                    if (created) File.Move(temp, fullTarget);
                    else File.Replace(temp, fullTarget, fullTarget + ".bak", ignoreMetadataErrors: true);
                }
                catch (IOException ex) { throw new DocumentLockedException(fullTarget, ex); }
                catch (UnauthorizedAccessException ex) { throw new DocumentLockedException(fullTarget, ex); }

                return (targetIndex, created);
            }
            finally
            {
                try { if (File.Exists(temp)) File.Delete(temp); } catch { /* فایلِ موقت؛ مهم نیست */ }
            }
        }

        /// <summary>
        /// 🌟 جدولِ ترک در سندِ مقصد: اول همان شماره اگر نامش می‌خورد، وگرنه تنها جدولی با همین نام.
        /// اگر نبود یا چند جدول هم‌نام بود، خطای روشن.
        /// </summary>
        private static int ResolveTargetTable(List<Table> qualifying, int trackIndex, string expectedName, string targetPath)
        {
            string want = expectedName.Trim();
            string NameAt(int i) => qualifying[i].Elements<TableRow>().First().InnerText.Trim();

            if (trackIndex >= 0 && trackIndex < qualifying.Count && string.Equals(NameAt(trackIndex), want, StringComparison.Ordinal))
                return trackIndex;

            var byName = Enumerable.Range(0, qualifying.Count)
                .Where(i => string.Equals(NameAt(i), want, StringComparison.Ordinal)).ToList();
            if (byName.Count == 1) return byName[0];
            if (byName.Count == 0)
                throw new InvalidOperationException(
                    $"{Path.GetFileName(targetPath)} has no table named \"{want}\". " +
                    "Its tables don't match the document you reviewed — check the file, or Shift+click \"To Word\" to pick another one.");
            throw new InvalidOperationException(
                $"{Path.GetFileName(targetPath)} has {byName.Count} tables named \"{want}\" and table {trackIndex + 1} isn't one of them, " +
                "so it isn't clear which one to update. Give the tables unique names (row 1) and try again.");
        }

        /// <summary>
        /// 🌟 متنِ خامِ ردیفِ دوم ← (متنِ بدونِ مارکر، مرزهای موجود اگر کامل‌اند، تعدادِ مارکرها).
        /// مرزها کامل‌اند وقتی دقیقاً یک [ms] پیشِ کلمه‌ی اول و یکی پسِ هر کلمه باشد — همان شکلی که
        /// <see cref="MarkerTextBuilder.ComputeInsertions"/> می‌نویسد.
        /// </summary>
        public static (string Clean, List<int>? Boundaries, int MarkerCount) ParseExistingMarkers(string raw)
        {
            int msCount = CountMsMarkers(raw);
            bool clock = HasClockMarkers(raw);
            if (msCount < MinMarkersToTreatAsMarked && !clock) return (raw, null, 0);

            // متنِ تمیز + جای هر مارکر در متنِ تمیز
            var sb = new StringBuilder();
            var markers = new List<(int Offset, int Ms, bool IsMs)>();
            int last = 0;
            foreach (Match m in AnyMarkerRegex.Matches(raw))
            {
                sb.Append(raw, last, m.Index - last);
                var ms = MsMarkerRegex.Match(m.Value);
                markers.Add((sb.Length, ms.Success && ms.Length == m.Length ? int.Parse(ms.Groups[1].Value) : -1, ms.Success && ms.Length == m.Length));
                last = m.Index + m.Length;
            }
            sb.Append(raw, last, raw.Length - last);
            string clean = sb.ToString();

            var tokens = TokenRegex.Matches(clean);
            if (tokens.Count == 0 || markers.Count != tokens.Count + 1 || markers.Any(x => !x.IsMs))
                return (clean, null, markers.Count);

            var boundaries = new List<int>(tokens.Count + 1);
            if (markers[0].Offset != tokens[0].Index) return (clean, null, markers.Count);
            boundaries.Add(markers[0].Ms);
            for (int i = 0; i < tokens.Count; i++)
            {
                var mk = markers[i + 1];
                if (mk.Offset != tokens[i].Index + tokens[i].Length) return (clean, null, markers.Count);
                boundaries.Add(mk.Ms);
            }
            for (int i = 1; i < boundaries.Count; i++)
                if (boundaries[i] < boundaries[i - 1]) return (clean, null, markers.Count);

            return (clean, boundaries, markers.Count);
        }

        /// <summary>بازه‌های [start, start+len) از متنِ تخت‌شده‌ی ردیف را از Text nodeهای مربوط حذف می‌کند.</summary>
        private static void RemoveRanges(List<TextRunRef> runs, List<(int Start, int Length)> ranges)
        {
            if (ranges.Count == 0) return;
            foreach (var run in runs)
            {
                string t = run.Element.Text ?? "";
                var keep = new StringBuilder(t.Length);
                for (int i = 0; i < t.Length; i++)
                {
                    int flat = run.StartOffset + i;
                    bool drop = ranges.Any(r => flat >= r.Start && flat < r.Start + r.Length);
                    if (!drop) keep.Append(t[i]);
                }
                if (keep.Length != t.Length)
                {
                    run.Element.Text = keep.ToString();
                    run.Element.Space = new EnumValue<SpaceProcessingModeValues>(SpaceProcessingModeValues.Preserve);
                }
            }
        }

        private static int CountMsMarkers(string text) => MsMarkerRegex.Matches(text).Count;

        private static bool HasClockMarkers(string text) =>
            AnyMarkerRegex.Matches(text).Cast<Match>().Any(m => m.Value.Contains(':'));

        // ---------------------------------------------------------------

        private static IEnumerable<Table> FindQualifyingTables(Body body)
        {
            foreach (var table in body.Descendants<Table>())
            {
                var rows = table.Elements<TableRow>().ToList();
                if (rows.Count < 2) continue;
                if (string.IsNullOrEmpty(rows[0].InnerText.Trim())) continue;
                yield return table;
            }
        }

        private class TextRunRef
        {
            public Text Element = null!;
            public int StartOffset;
            public int Length;
        }

        /// <summary>Concatenates row[1]'s paragraph/run text (paragraph breaks -> "\n",
        /// &lt;w:br/&gt; -> " "), and records each real Text node's offset range for splicing.</summary>
        private static string FlattenRow1Text(TableRow row1, out List<TextRunRef> textRuns)
        {
            var sb = new StringBuilder();
            var runs = new List<TextRunRef>();

            foreach (var cell in row1.Elements<TableCell>())
            {
                foreach (var paragraph in cell.Elements<Paragraph>())
                {
                    foreach (var run in paragraph.Elements<Run>())
                    {
                        foreach (var child in run.ChildElements)
                        {
                            if (child is Text text)
                            {
                                int start = sb.Length;
                                sb.Append(text.Text);
                                runs.Add(new TextRunRef { Element = text, StartOffset = start, Length = text.Text.Length });
                            }
                            else if (child is Break || child is TabChar)
                            {
                                sb.Append(' ');
                            }
                        }
                    }
                    sb.Append('\n');
                }
            }

            textRuns = runs;
            // Row content ends with a trailing "\n" from the loop above; trim exactly one
            // to match how MarkerTextBuilder/align_batch.py treat a plain-text file's contents
            // (no forced trailing newline).
            string result = sb.ToString();
            if (result.EndsWith("\n")) result = result[..^1];
            return result;
        }

        private static TextRunRef FindInsertionTarget(List<TextRunRef> textRuns, int offset)
        {
            if (textRuns.Count == 0)
                throw new InvalidOperationException("No text runs found to insert markers into.");

            TextRunRef best = textRuns[0];
            foreach (var run in textRuns)
            {
                if (run.StartOffset <= offset) best = run;
                if (run.StartOffset > offset) break;
            }
            return best;
        }
    }

    /// <summary>🌟 سندِ مقصد قفل است (معمولاً چون در Word باز است). فایلِ اصلی دست نخورده.</summary>
    public sealed class DocumentLockedException : IOException
    {
        public string FilePath { get; }

        public DocumentLockedException(string path, Exception inner)
            : base($"{Path.GetFileName(path)} is open in another program (probably Word). " +
                   "Close it there and press \"To Word\" again — nothing was changed.", inner)
        {
            FilePath = path;
        }
    }
}
