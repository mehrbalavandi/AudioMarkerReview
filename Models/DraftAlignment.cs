using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace AudioMarkerReview.Models
{
    /// <summary>
    /// Mirrors the *_draft.json files produced by align_engine/align_batch.py.
    /// boundariesMs.Count must equal words.Count + 1; scores.Count must equal words.Count.
    /// </summary>
    public class DraftAlignment
    {
        [JsonPropertyName("audioFile")]
        public string AudioFile { get; set; } = "";

        [JsonPropertyName("words")]
        public List<string> Words { get; set; } = new();

        [JsonPropertyName("boundariesMs")]
        public List<int> BoundariesMs { get; set; } = new();

        [JsonPropertyName("scores")]
        public List<double> Scores { get; set; } = new();
    }

    /// <summary>
    /// One word plus the two boundary indices that bracket it, used for data-binding
    /// the word list panel. Index i's EndBoundaryIndex == index i+1's StartBoundaryIndex
    /// — dragging updates the shared boundaries array in MainWindow, not a copy here.
    /// </summary>
    public class WordItem
    {
        public int Index { get; set; }
        public string Word { get; set; } = "";
        public double Score { get; set; }
        public int StartMs { get; set; }
        public int EndMs { get; set; }
    }
}
