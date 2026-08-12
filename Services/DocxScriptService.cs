using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
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

        /// <summary>Label shown in the Track combo box.</summary>
        public string Display =>
            $"{TrackIndex + 1}. {AudioTrackName}" + (WordCount > 0 ? $"  ({WordCount} words)" : "");
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
                string fullText = FlattenRow1Text(rows[1], out _);
                tracks.Add(new DocxTrackInfo
                {
                    TrackIndex = index,
                    AudioTrackName = trackName,
                    FullText = fullText,
                    WordCount = MarkerTextBuilder.SplitWords(fullText).Count
                });
                index++;
            }
            return tracks;
        }

        /// <summary>
        /// Writes a NEW .docx at destPath: a copy of sourcePath with "[ms]" markers spliced
        /// into the same track's row[1] runs. words/boundariesMs must be the reviewed/corrected
        /// values (boundariesMs.Count == words.Count + 1, see MarkerTextBuilder).
        /// </summary>
        /// <param name="expectedTrackName">Optional guard for multi-table documents: the row[0]
        /// caption the caller believes sits at <paramref name="trackIndex"/>. If the document was
        /// edited between loading and exporting (a table added, removed or reordered), the index
        /// would silently point at a different track — this turns that into an error instead.</param>
        public static void ExportWithMarkers(string sourcePath, int trackIndex,
            IReadOnlyList<string> words, IReadOnlyList<int> boundariesMs, string destPath,
            string? expectedTrackName = null)
        {
            if (!string.Equals(Path.GetFullPath(sourcePath), Path.GetFullPath(destPath), StringComparison.OrdinalIgnoreCase))
                File.Copy(sourcePath, destPath, overwrite: true);
            else
                throw new IOException("Choose a different file name than the source .docx (Save As, not overwrite).");

            using var doc = WordprocessingDocument.Open(destPath, true);
            var body = doc.MainDocumentPart?.Document?.Body
                       ?? throw new InvalidDataException("This .docx has no document body.");

            var qualifying = FindQualifyingTables(body).ToList();
            if (trackIndex < 0 || trackIndex >= qualifying.Count)
                throw new InvalidOperationException(
                    $"Track index {trackIndex} no longer exists in this document ({qualifying.Count} qualifying table(s) found).");

            var rows = qualifying[trackIndex].Elements<TableRow>().ToList();

            if (expectedTrackName != null)
            {
                string actualName = rows[0].InnerText.Trim();
                if (!string.Equals(actualName, expectedTrackName.Trim(), StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        $"Table {trackIndex + 1} is now \"{actualName}\" but you reviewed \"{expectedTrackName}\". " +
                        "The document changed since you opened it — re-open it in the review tool and try again.");
            }

            string fullText = FlattenRow1Text(rows[1], out var textRuns);

            var freshWords = MarkerTextBuilder.SplitWords(fullText);
            if (freshWords.Count != words.Count)
                throw new InvalidOperationException(
                    $"The document's text has changed since you loaded it: {freshWords.Count} words now vs " +
                    $"{words.Count} when reviewed. Re-open the document in the review tool and try again.");

            var insertions = MarkerTextBuilder.ComputeInsertions(fullText, words, boundariesMs);
            foreach (var (offset, marker) in insertions.OrderByDescending(i => i.Offset))
            {
                var target = FindInsertionTarget(textRuns, offset);
                int localOffset = Math.Max(0, Math.Min(target.Length, offset - target.StartOffset));
                string current = target.Element.Text ?? "";
                target.Element.Text = current.Insert(localOffset, marker);
                // guard against Word/OpenXml trimming the whitespace next to the markers on save
                target.Element.Space = new DocumentFormat.OpenXml.EnumValue<SpaceProcessingModeValues>(SpaceProcessingModeValues.Preserve);
            }

            doc.MainDocumentPart!.Document.Save();
        }

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
}
