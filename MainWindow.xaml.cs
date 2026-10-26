using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using NAudio.Wave;
using SoundTouch.Net.NAudioSupport;
using AudioMarkerReview.Models;
using AudioMarkerReview.Services;

namespace AudioMarkerReview
{
    public partial class MainWindow : Window
    {
        // ---- loaded content -------------------------------------------------
        private DecodedAudio? _audio;
        private string? _audioPath;
        private string _originalText = "";
        private List<string> _words = new();
        private List<int> _boundariesMs = new();   // length == _words.Count + 1
        private List<double> _scores = new();      // length == _words.Count (0 = unknown/manual)

        // ---- Word-document source (when the text came from a .docx track
        // instead of a plain .txt file) --------------------------------------
        private string? _docxSourcePath;
        private int _docxTrackIndex = -1;
        private List<DocxTrackInfo> _docxTracks = new();
        private bool _suppressTrackSelectionEvent;
        private string _trackMatchNote = "";   // how the current track got picked, for the status bar
        private bool _autoMatchEnabled = true; // mirrors AutoMatchCheckBox, kept as a plain field so
                                               // the matching logic never touches a XAML-generated name

        // 🌟 سندی که «To Word» مارکرها را در آن می‌نویسد (پیش‌فرض: ‎<نامِ سند>_markers.docx کنارِ سند؛
        // اگر خودِ سندِ _markers باز شده، همان). هر بار فقط جدولِ همین ترک عوض می‌شود.
        private string? _docxTargetPath;
        private List<DocxTrackInfo> _targetTracks = new();
        private string _targetNote = "";       // وضعیتِ سندِ مقصد برای نوارِ وضعیت
        private bool _reviewedMarkersLoaded;   // مارکرهای این ترک از سندِ مقصد آمده‌اند؛ پیش‌نویس رویشان نمی‌نشیند

        // ---- playback ---------------------------------------------------------
        private AudioFileReader? _playbackReader;
        private SoundTouchWaveProvider? _soundTouch; // wraps _playbackReader for pitch-preserving speed control
        private WaveOutEvent? _waveOut;
        private readonly DispatcherTimerWrapper _playbackTimer;
        private int _currentWordIndex = -1;

        // ---- waveform / markers ------------------------------------------------
        private double _pixelsPerSecond = 60;
        private const double MinPixelsPerSecond = 8;
        private const double MaxPixelsPerSecond = 500;
        private const double MinMarkerGapPixels = 3;
        private const double WaveHeight = 180;
        private readonly List<Thumb> _thumbs = new();
        private readonly List<Border> _wordChips = new();

        // ---- marker multi-selection / batch shift ------------------------------
        private readonly HashSet<int> _selectedMarkerIndices = new();
        private int? _selectionAnchorIndex;
        private bool _isGroupDrag;
        private double _groupDragAccumulatedDeltaPx;
        private double _groupDragLowerBoundPx;
        private double _groupDragUpperBoundPx;
        private readonly Dictionary<int, double> _groupDragStartLeftPx = new();

        // ---- undo / redo (boundariesMs snapshots, one per completed edit) ------
        private readonly Stack<List<int>> _undoStack = new();
        /// <summary>Status-bar note about the draft JSON (auto-loaded / not found / mismatch).</summary>
        private string _draftNote = "";
        private readonly Stack<List<int>> _redoStack = new();
        private const int MaxUndoDepth = 100;

        public MainWindow()
        {
            InitializeComponent();
            _playbackTimer = new DispatcherTimerWrapper(TimeSpan.FromMilliseconds(50), PlaybackTimer_Tick);
            Closing += (_, _) => CleanupPlayback();
            PreviewKeyDown += MainWindow_PreviewKeyDown;
        }

        private void MainWindow_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (Keyboard.Modifiers != ModifierKeys.Control) return;
            if (e.Key == Key.Z && UndoButton.IsEnabled) { Undo(); e.Handled = true; }
            else if (e.Key == Key.Y && RedoButton.IsEnabled) { Redo(); e.Handled = true; }
        }

        // =====================================================================
        // File loading
        // =====================================================================

        private void OpenAudioButton_Click(object sender, RoutedEventArgs e)
        {
            // 🌟 در پوشه‌ی فایلِ جاری باز می‌شود و تا خودِ آن اسکرول می‌کند (↓ + Enter = فایلِ بعدی)
            string? picked = AudioFilePicker.Pick(this, _audioPath);
            if (picked == null) return;

            try
            {
                CleanupPlayback();
                _audio = AudioLoader.Load(picked);
                _audioPath = picked;

                _playbackReader = new AudioFileReader(picked);
                _soundTouch = new SoundTouchWaveProvider(_playbackReader);
                _soundTouch.OptimizeForSpeech();
                ApplyCurrentSpeedSelection();

                _waveOut = new WaveOutEvent();
                _waveOut.Init(_soundTouch);
                _waveOut.PlaybackStopped += (_, _) => Dispatcher.Invoke(() => PlayPauseButton.Content = "Play");

                PlayPauseButton.IsEnabled = true;

                // A document can hold many tracks — now that we know which audio file is
                // being reviewed, jump to the table whose caption names it.
                TryAutoSelectTrackForAudio(announce: true);
                // align_engine writes <audio>_draft.json into <audio folder>\output — pick it
                // up by itself instead of asking for "Load Draft JSON..." every time.
                TryAutoLoadDraft();

                EnsureBoundariesInitialized();
                RebuildWaveform();
                UpdateStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Could not load audio:\n{ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenTextButton_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { Filter = "Text files|*.txt|All files|*.*" };
            if (dlg.ShowDialog() != true) return;

            try
            {
                // Loading a plain .txt directly severs any prior association with a
                // .docx track — "Export to Word" only makes sense when the text came
                // from a docx track, since that's what carries the source path/table.
                _docxSourcePath = null;
                _docxTrackIndex = -1;
                _docxTracks = new List<DocxTrackInfo>();
                _docxTargetPath = null;
                _targetTracks = new List<DocxTrackInfo>();
                _targetNote = "";
                _reviewedMarkersLoaded = false;
                _trackMatchNote = "";
                _suppressTrackSelectionEvent = true;
                TrackComboBox.ItemsSource = null;
                TrackComboBox.IsEnabled = false;
                _suppressTrackSelectionEvent = false;

                LoadOriginalText(File.ReadAllText(dlg.FileName));
                TryAutoLoadDraft();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Could not load text:\n{ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenDocxButton_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { Filter = "Word documents|*.docx|All files|*.*" };
            if (dlg.ShowDialog() != true) return;


            string originalFile = dlg.FileName;

            try
            {
                // ساخت فایل موقت: ExtractTracksShared از یک کپیِ موقت می‌خواند (حتی اگر سند در Word باز و
                // قفل باشد) و کپی را بعد از خواندن پاک می‌کند.
                var tracks = DocxScriptService.ExtractTracksShared(originalFile);
                if (tracks.Count == 0)
                {
                    MessageBox.Show(this,
                        "No qualifying audio-script table found in this document (needs a table with " +
                        "at least 2 rows and a non-empty row 1, same as CustomLayoutGenerator expects).",
                        "No tracks found", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                _docxSourcePath = dlg.FileName;
                _docxTracks = tracks;
                _trackMatchNote = "";
                _docxTrackIndex = -1;   // an index from the previous document means nothing here

                // 🌟 مقصدِ «To Word» و مارکرهایی که قبلاً در آن نوشته شده‌اند
                _docxTargetPath = DefaultTargetFor(dlg.FileName);
                RefreshTargetTracks();

                _suppressTrackSelectionEvent = true;
                TrackComboBox.ItemsSource = _docxTracks;
                TrackComboBox.IsEnabled = true;
                _suppressTrackSelectionEvent = false;

                // If an audio file is already open, land on ITS table instead of the first
                // one; otherwise show track 1 and re-match as soon as audio arrives.
                if (!TryAutoSelectTrackForAudio(announce: true))
                    SelectTrack(0);
                TryAutoLoadDraft();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Could not read Word document:\n{ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void TrackComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_suppressTrackSelectionEvent || TrackComboBox.SelectedIndex < 0) return;
            _trackMatchNote = "chosen manually";
            LoadDocxTrack(TrackComboBox.SelectedIndex);
            TryAutoLoadDraft();
        }

        private void AutoMatchCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            // Read the state off the sender rather than the x:Name field: this handler runs
            // once while the XAML is still being parsed, before the generated field is assigned.
            _autoMatchEnabled = sender is CheckBox box && box.IsChecked == true;
            if (!IsLoaded) return;

            if (_autoMatchEnabled) TryAutoSelectTrackForAudio(announce: false);
            else _trackMatchNote = "auto-match off";
            UpdateStatus();
        }

        /// <summary>
        /// Picks the .docx table whose row[0] caption best names the currently loaded audio
        /// file, and loads it. Returns true when a track ended up selected, so the caller
        /// knows whether it still has to fall back to track 1. Does nothing when there is no
        /// audio yet, no document yet, or auto-match is switched off.
        /// </summary>
        private bool TryAutoSelectTrackForAudio(bool announce)
        {
            if (_docxTracks.Count == 0 || string.IsNullOrEmpty(_audioPath)) return false;
            if (!_autoMatchEnabled)
            {
                _trackMatchNote = "auto-match off";
                return false;
            }

            var match = TrackMatcher.Match(_audioPath!, _docxTracks);
            if (match.Index < 0) return false;

            if (match.IsConfident)
            {
                _trackMatchNote = $"auto-matched {match.Score:P0}";
            }
            else
            {
                _trackMatchNote = $"UNCERTAIN match {match.Score:P0} — check the Track list";
                if (announce)
                    MessageBox.Show(this,
                        $"Couldn't confidently tell which table belongs to \"{Path.GetFileName(_audioPath)}\".\n" +
                        $"Reason: {match.Reason}.\n\n" +
                        BuildCandidateList(match) + "\n" +
                        "The closest one has been selected — use the Track list to correct it if that's wrong.",
                        "Track match uncertain", MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            // Already on the right track: leave the loaded words/markers alone so that
            // re-opening the same audio doesn't throw away the edits made so far.
            if (match.Index == _docxTrackIndex) return true;

            SelectTrack(match.Index);
            return true;
        }

        /// <summary>Moves the combo box AND loads the track, without the selection event
        /// firing back and stamping the choice as "manual".</summary>
        private void SelectTrack(int index)
        {
            if (index < 0 || index >= _docxTracks.Count) return;
            _suppressTrackSelectionEvent = true;
            TrackComboBox.SelectedIndex = index;
            _suppressTrackSelectionEvent = false;
            LoadDocxTrack(index);
        }

        private string BuildCandidateList(TrackMatchResult match)
        {
            var sb = new System.Text.StringBuilder("Closest tables:\n");
            foreach (var (index, score) in match.Ranked.Take(3))
                sb.AppendLine($"   {score:P0}  —  {_docxTracks[index].Display}");
            return sb.ToString();
        }

        private void LoadDocxTrack(int index)
        {
            _docxTrackIndex = index;
            // Switching tracks means any previously loaded draft/boundaries belonged
            // to the OTHER track's words — drop them rather than risk silently
            // pairing track B's text with track A's timings.
            _words = new List<string>();
            _boundariesMs = new List<int>();
            _scores = new List<double>();
            _draftNote = "";
            _reviewedMarkersLoaded = false;

            // 🌟 ترکی که قبلاً بازبینی و در سندِ مقصد نوشته شده، با همان مارکرها باز می‌شود
            // (نه با پیش‌نویسِ align_engine و نه با مارکرهای یکنواخت).
            var track = _docxTracks[index];
            var reviewed = track.ReviewedBoundariesMs;
            if (reviewed != null)
            {
                var words = MarkerTextBuilder.SplitWords(track.FullText);
                if (reviewed.Count == words.Count + 1)
                {
                    _words = words;
                    _boundariesMs = new List<int>(reviewed);
                    _scores = new List<double>(new double[words.Count]);
                    _selectedMarkerIndices.Clear();
                    _selectionAnchorIndex = null;
                    ResetUndoHistory();
                    _reviewedMarkersLoaded = true;
                    _draftNote = $"reviewed markers from {Path.GetFileName(_docxTargetPath)}";
                }
            }
            LoadOriginalText(track.FullText);
        }

        /// <summary>🌟 ‎«X.docx» ← «X_markers.docx»؛ اگر خودِ سندِ _markers باز شده، همان.</summary>
        private static string DefaultTargetFor(string docxPath)
        {
            string stem = Path.GetFileNameWithoutExtension(docxPath);
            if (stem.EndsWith("_markers", StringComparison.OrdinalIgnoreCase)) return docxPath;
            return Path.Combine(Path.GetDirectoryName(docxPath) ?? "", stem + "_markers.docx");
        }

        private bool TargetIsSource =>
            _docxTargetPath != null && _docxSourcePath != null &&
            string.Equals(Path.GetFullPath(_docxTargetPath), Path.GetFullPath(_docxSourcePath), StringComparison.OrdinalIgnoreCase);

        /// <summary>
        /// 🌟 سندِ مقصد را (اگر هست) می‌خواند و برای هر ترکِ سندِ باز‌شده مشخص می‌کند که آن‌جا
        /// مارکرِ کامل دارد یا نه — هم برای علامتِ ✓ در فهرستِ ترک‌ها و هم برای باز کردنِ ترک با
        /// مارکرهای بازبینی‌شده. ترک با نام (و در صورتِ تکرارِ نام، با شماره) پیدا می‌شود.
        /// </summary>
        private void RefreshTargetTracks()
        {
            _targetTracks = new List<DocxTrackInfo>();
            _targetNote = "";
            if (_docxTargetPath == null) return;

            if (TargetIsSource) _targetTracks = _docxTracks;
            else if (File.Exists(_docxTargetPath))
            {
                try { _targetTracks = DocxScriptService.ExtractTracksShared(_docxTargetPath); }
                catch (Exception ex) { _targetNote = $"could not read {Path.GetFileName(_docxTargetPath)}: {ex.Message}"; }
            }

            int marked = 0;
            foreach (var t in _docxTracks)
            {
                var match = _targetTracks.Where(x => x.AudioTrackName == t.AudioTrackName)
                    .OrderBy(x => x.TrackIndex == t.TrackIndex ? 0 : 1).FirstOrDefault();
                bool full = match?.ExistingBoundariesMs != null && match.WordCount == t.WordCount;
                t.ReviewedBoundariesMs = full ? match!.ExistingBoundariesMs : null;
                t.PartlyMarkedInTarget = !full && match != null && match.ExistingMarkerCount > 0;
                if (full) marked++;
            }

            if (_targetNote.Length == 0)
                _targetNote = File.Exists(_docxTargetPath)
                    ? $"To Word → {Path.GetFileName(_docxTargetPath)} ({marked}/{_docxTracks.Count} tracks marked)"
                    : $"To Word → new {Path.GetFileName(_docxTargetPath)}";

            // فهرستِ ترک‌ها علامت‌ها را نشان دهد (Display محاسبه‌ای است، پس دوباره بایند می‌شود)
            _suppressTrackSelectionEvent = true;
            int sel = TrackComboBox.SelectedIndex;
            TrackComboBox.ItemsSource = null;
            TrackComboBox.ItemsSource = _docxTracks;
            TrackComboBox.SelectedIndex = sel < _docxTracks.Count ? sel : -1;
            _suppressTrackSelectionEvent = false;
        }

        /// <summary>
        /// Shared by both the .txt and .docx-track loading paths: validates against
        /// an already-loaded draft (if any), falls back to evenly-spaced markers if
        /// there's none yet, and refreshes the UI.
        /// </summary>
        private void LoadOriginalText(string text)
        {
            _originalText = text;
            if (_words.Count > 0)
            {
                var textWords = MarkerTextBuilder.SplitWords(_originalText);
                if (textWords.Count != _words.Count)
                {
                    MessageBox.Show(this,
                        $"This text has {textWords.Count} words but the loaded draft JSON has " +
                        $"{_words.Count}. Load the matching draft JSON, or continue to start from " +
                        "evenly-spaced markers instead.",
                        "Word count mismatch", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            EnsureBoundariesInitialized();
            RebuildWaveform();
            RebuildWordPanel();
            UpdateStatus();
        }

        private void LoadDraftButton_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { Filter = "Draft JSON|*.json|All files|*.*" };
            string? guessDir = DraftFolderFor(_audioPath);
            if (guessDir != null && Directory.Exists(guessDir)) dlg.InitialDirectory = guessDir;
            if (dlg.ShowDialog() != true) return;
            LoadDraftFile(dlg.FileName, interactive: true);
        }

        /// <summary>
        /// Loads one *_draft.json (align_engine output). interactive=true is the "Load Draft
        /// JSON..." button: errors and word-count warnings are shown and the draft is loaded
        /// anyway. interactive=false is the automatic pick-up: it stays silent and refuses a
        /// draft whose word count doesn't match the loaded text (it belongs to another track
        /// or an older version of the text). Returns true when the draft was loaded.
        /// </summary>
        private bool LoadDraftFile(string path, bool interactive)
        {
            try
            {
                var json = File.ReadAllText(path);
                var opts = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                var draft = JsonSerializer.Deserialize<DraftAlignment>(json, opts)
                            ?? throw new InvalidDataException("Empty or invalid draft JSON.");

                if (draft.BoundariesMs.Count != draft.Words.Count + 1)
                    throw new InvalidDataException(
                        $"boundariesMs should have {draft.Words.Count + 1} entries, found {draft.BoundariesMs.Count}.");
                if (draft.Scores.Count != draft.Words.Count)
                    throw new InvalidDataException(
                        $"scores should have {draft.Words.Count} entries, found {draft.Scores.Count}.");

                if (!string.IsNullOrEmpty(_originalText))
                {
                    var textWords = MarkerTextBuilder.SplitWords(_originalText);
                    if (textWords.Count != draft.Words.Count)
                    {
                        if (!interactive)
                        {
                            _draftNote = $"{Path.GetFileName(path)} NOT loaded: {draft.Words.Count} words vs {textWords.Count} in the text";
                            UpdateStatus();
                            return false;
                        }
                        MessageBox.Show(this,
                            $"The loaded text has {textWords.Count} words but this draft has " +
                            $"{draft.Words.Count}. Loading it anyway, but export will fail until " +
                            "you load the matching text file.",
                            "Word count mismatch", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }

                _words = draft.Words;
                _boundariesMs = draft.BoundariesMs;
                _scores = draft.Scores;
                _reviewedMarkersLoaded = false;
                _selectedMarkerIndices.Clear();
                _selectionAnchorIndex = null;
                ResetUndoHistory();
                _draftNote = $"draft: {Path.GetFileName(path)}" + (interactive ? "" : " (auto)");

                RebuildWaveform();
                RebuildWordPanel();
                UpdateStatus();
                return true;
            }
            catch (Exception ex)
            {
                if (interactive)
                    MessageBox.Show(this, $"Could not load draft JSON:\n{ex.Message}", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                else
                    _draftNote = $"{Path.GetFileName(path)} could not be read: {ex.Message}";
                UpdateStatus();
                return false;
            }
        }

        /// <summary>Where align_engine puts the drafts by default: &lt;audio folder&gt;\output.</summary>
        private static string? DraftFolderFor(string? audioPath)
        {
            if (string.IsNullOrEmpty(audioPath)) return null;
            string? dir = Path.GetDirectoryName(audioPath);
            return dir == null ? null : Path.Combine(dir, "output");
        }

        /// <summary>
        /// Once both the audio and its text are loaded, looks for &lt;audio name&gt;_draft.json
        /// in &lt;audio folder&gt;\output (align_engine's default) and then next to the audio,
        /// and loads it silently. Never replaces markers you have already moved (undo history
        /// not empty) — use "Load Draft JSON..." for that.
        /// </summary>
        private void TryAutoLoadDraft()
        {
            if (_audio == null || string.IsNullOrEmpty(_audioPath) || string.IsNullOrEmpty(_originalText)) return;
            if (_undoStack.Count > 0) return;
            // 🌟 مارکرهای بازبینی‌شده‌ی سندِ مقصد از پیش‌نویسِ خام بهترند؛ «Load Draft JSON...» هنوز کار می‌کند.
            if (_reviewedMarkersLoaded) return;

            string stem = Path.GetFileNameWithoutExtension(_audioPath);
            string? audioDir = Path.GetDirectoryName(_audioPath);
            if (audioDir == null) return;
            var candidates = new[]
            {
                Path.Combine(audioDir, "output", stem + "_draft.json"),
                Path.Combine(audioDir, stem + "_draft.json"),
            };
            foreach (var candidate in candidates)
            {
                if (File.Exists(candidate) && LoadDraftFile(candidate, interactive: false))
                    return;
            }
            if (!candidates.Any(File.Exists))
            {
                _draftNote = "no draft found — run align_engine for this track, or use Load Draft JSON...";
                UpdateStatus();
            }
        }

        /// <summary>
        /// Fallback for when no draft JSON is loaded (or hasn't been loaded yet): once
        /// both audio and text are available, evenly space markers across the audio's
        /// duration so there's always something on the waveform to drag, instead of
        /// requiring the automated aligner to be run first.
        /// </summary>
        private void EnsureBoundariesInitialized()
        {
            if (_audio == null || string.IsNullOrEmpty(_originalText)) return;
            if (_words.Count > 0 && _boundariesMs.Count == _words.Count + 1) return; // already have real data

            _words = MarkerTextBuilder.SplitWords(_originalText);
            if (_words.Count == 0) return;

            int totalMs = (int)Math.Round(_audio.DurationSeconds * 1000);
            _boundariesMs = new List<int>(_words.Count + 1);
            for (int i = 0; i <= _words.Count; i++)
                _boundariesMs.Add((int)Math.Round(totalMs * (i / (double)_words.Count)));
            _scores = new List<double>(new double[_words.Count]); // all 0 => "unknown" grey chips
            _selectedMarkerIndices.Clear();
            _selectionAnchorIndex = null;
            ResetUndoHistory();

            RebuildWordPanel();
        }

        // =====================================================================
        // Waveform + marker rendering
        // =====================================================================

        private void RebuildWaveform()
        {
            if (_audio == null) return;

            double totalWidth = Math.Max(WaveScrollViewer.ActualWidth, _audio.DurationSeconds * _pixelsPerSecond);
            WaveHostGrid.Width = totalWidth;
            WaveformPath.Data = WaveformRenderer.BuildPeaksGeometry(_audio.Samples, totalWidth, WaveHeight);
            RebuildMarkers();
        }

        private void RebuildMarkers()
        {
            MarkerCanvas.Children.Clear();
            _thumbs.Clear();
            if (_boundariesMs.Count == 0) return;

            for (int i = 0; i < _boundariesMs.Count; i++)
            {
                var thumb = new Thumb
                {
                    Style = (Style)FindResource("MarkerThumbStyle"),
                    Height = WaveHeight,
                    Tag = i,
                };
                Canvas.SetTop(thumb, 0);
                Canvas.SetLeft(thumb, MsToPixels(_boundariesMs[i]));
                thumb.PreviewMouseLeftButtonDown += MarkerThumb_PreviewMouseLeftButtonDown;
                thumb.DragStarted += MarkerThumb_DragStarted;
                thumb.DragDelta += MarkerThumb_DragDelta;
                MarkerCanvas.Children.Add(thumb);
                _thumbs.Add(thumb);
            }

            // indices in _selectedMarkerIndices are still valid here (this method
            // only replaces the visuals, never the boundaries array itself) — just
            // repaint which thumbs should show as selected and refresh the buttons.
            RefreshMarkerSelectionVisuals();
        }

        /// <summary>Ctrl+Click toggles one marker; Shift+Click selects the contiguous
        /// range from the last-touched marker to this one. A plain click is left
        /// alone here so the Thumb's own drag-start still fires normally.</summary>
        private void MarkerThumb_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            int i = (int)((Thumb)sender).Tag;
            var mods = Keyboard.Modifiers;

            if (mods.HasFlag(ModifierKeys.Control))
            {
                if (!_selectedMarkerIndices.Remove(i))
                    _selectedMarkerIndices.Add(i);
                _selectionAnchorIndex = i;
                RefreshMarkerSelectionVisuals();
                e.Handled = true;
            }
            else if (mods.HasFlag(ModifierKeys.Shift) && _selectionAnchorIndex.HasValue)
            {
                int a = Math.Min(_selectionAnchorIndex.Value, i);
                int b = Math.Max(_selectionAnchorIndex.Value, i);
                _selectedMarkerIndices.Clear();
                for (int k = a; k <= b; k++) _selectedMarkerIndices.Add(k);
                RefreshMarkerSelectionVisuals();
                e.Handled = true;
            }
            // else: plain click — handled by DragStarted below, not here.
        }

        /// <summary>Decides, once per drag gesture, whether this is a batch move
        /// (dragging a marker that's part of a >1-marker selection) or the original
        /// single-marker drag — and for batch mode, precomputes each selected
        /// marker's start position plus the shared clamp range for the whole group.</summary>
        private void MarkerThumb_DragStarted(object sender, DragStartedEventArgs e)
        {
            int i = (int)((Thumb)sender).Tag;
            PushUndoSnapshot(); // one snapshot per gesture, whether it ends up single or group

            if (_selectedMarkerIndices.Contains(i) && _selectedMarkerIndices.Count > 1)
            {
                _isGroupDrag = true;
                _groupDragAccumulatedDeltaPx = 0;
                _groupDragStartLeftPx.Clear();
                foreach (int idx in _selectedMarkerIndices)
                    _groupDragStartLeftPx[idx] = MsToPixels(_boundariesMs[idx]);

                int minSelected = _selectedMarkerIndices.Min();
                int maxSelected = _selectedMarkerIndices.Max();
                _groupDragLowerBoundPx = (minSelected > 0)
                    ? MsToPixels(_boundariesMs[minSelected - 1]) + MinMarkerGapPixels : 0;
                _groupDragUpperBoundPx = (maxSelected < _boundariesMs.Count - 1)
                    ? MsToPixels(_boundariesMs[maxSelected + 1]) - MinMarkerGapPixels : WaveHostGrid.Width;
            }
            else
            {
                _isGroupDrag = false;
                // Plain drag on a marker outside the current selection: that marker
                // becomes the new (single) selection, same as clicking it alone would.
                _selectedMarkerIndices.Clear();
                _selectedMarkerIndices.Add(i);
                _selectionAnchorIndex = i;
                RefreshMarkerSelectionVisuals();
            }
        }

        private void MarkerThumb_DragDelta(object sender, DragDeltaEventArgs e)
        {
            var thumb = (Thumb)sender;
            int i = (int)thumb.Tag;

            if (_isGroupDrag)
            {
                _groupDragAccumulatedDeltaPx += e.HorizontalChange;
                double minAllowed = _groupDragLowerBoundPx - _groupDragStartLeftPx[_selectedMarkerIndices.Min()];
                double maxAllowed = _groupDragUpperBoundPx - _groupDragStartLeftPx[_selectedMarkerIndices.Max()];
                _groupDragAccumulatedDeltaPx = Math.Max(minAllowed, Math.Min(maxAllowed, _groupDragAccumulatedDeltaPx));

                foreach (int idx in _selectedMarkerIndices)
                {
                    double left = _groupDragStartLeftPx[idx] + _groupDragAccumulatedDeltaPx;
                    Canvas.SetLeft(_thumbs[idx], left);
                    _boundariesMs[idx] = PixelsToMs(left);
                }
                return;
            }

            double newLeft = Canvas.GetLeft(thumb) + e.HorizontalChange;

            double lowerBound = (i > 0) ? MsToPixels(_boundariesMs[i - 1]) + MinMarkerGapPixels : 0;
            double upperBound = (i < _boundariesMs.Count - 1)
                ? MsToPixels(_boundariesMs[i + 1]) - MinMarkerGapPixels
                : WaveHostGrid.Width;

            newLeft = Math.Max(lowerBound, Math.Min(upperBound, newLeft));
            Canvas.SetLeft(thumb, newLeft);
            _boundariesMs[i] = PixelsToMs(newLeft);
        }

        private double MsToPixels(int ms) => ms / 1000.0 * _pixelsPerSecond;
        private int PixelsToMs(double px) => (int)Math.Round(px / _pixelsPerSecond * 1000.0);

        // =====================================================================
        // Marker multi-selection / batch shift
        // =====================================================================

        private void RefreshMarkerSelectionVisuals()
        {
            for (int idx = 0; idx < _thumbs.Count; idx++)
                _thumbs[idx].Style = (Style)FindResource(
                    _selectedMarkerIndices.Contains(idx) ? "MarkerThumbSelectedStyle" : "MarkerThumbStyle");

            bool hasSelection = _selectedMarkerIndices.Count > 0;
            SelectionInfoText.Text = _selectedMarkerIndices.Count.ToString();
            NudgeMinus50Button.IsEnabled = hasSelection;
            NudgeMinus10Button.IsEnabled = hasSelection;
            NudgePlus10Button.IsEnabled = hasSelection;
            NudgePlus50Button.IsEnabled = hasSelection;
            ClearSelectionButton.IsEnabled = hasSelection;
        }

        private void ClearSelectionButton_Click(object sender, RoutedEventArgs e)
        {
            _selectedMarkerIndices.Clear();
            _selectionAnchorIndex = null;
            RefreshMarkerSelectionVisuals();
        }

        private void NudgeMinus50Button_Click(object sender, RoutedEventArgs e) => NudgeSelection(-50);
        private void NudgeMinus10Button_Click(object sender, RoutedEventArgs e) => NudgeSelection(-10);
        private void NudgePlus10Button_Click(object sender, RoutedEventArgs e) => NudgeSelection(10);
        private void NudgePlus50Button_Click(object sender, RoutedEventArgs e) => NudgeSelection(50);

        /// <summary>Shifts every selected marker by the same requestedDeltaMs, clamped
        /// as a whole group so the selection never crosses its nearest unselected
        /// neighbor on either side (same guarantee the drag path gives).</summary>
        private void NudgeSelection(int requestedDeltaMs)
        {
            if (_selectedMarkerIndices.Count == 0 || _boundariesMs.Count == 0) return;

            int minSelected = _selectedMarkerIndices.Min();
            int maxSelected = _selectedMarkerIndices.Max();
            int lowerBoundMs = (minSelected > 0) ? _boundariesMs[minSelected - 1] + 1 : 0;
            int upperBoundMs = (maxSelected < _boundariesMs.Count - 1)
                ? _boundariesMs[maxSelected + 1] - 1
                : (_audio != null ? (int)Math.Round(_audio.DurationSeconds * 1000) : int.MaxValue);

            int delta = requestedDeltaMs;
            if (delta < 0)
                delta = Math.Max(delta, lowerBoundMs - _boundariesMs[minSelected]);
            else if (delta > 0)
                delta = Math.Min(delta, upperBoundMs - _boundariesMs[maxSelected]);
            if (delta == 0) return;

            PushUndoSnapshot();
            foreach (int idx in _selectedMarkerIndices)
                _boundariesMs[idx] += delta;

            RebuildMarkers(); // repaints positions; selection itself is untouched
        }

        // =====================================================================
        // Undo / redo
        // =====================================================================

        /// <summary>Call BEFORE applying an edit that's actually going to change
        /// _boundariesMs (drag start, or a nudge that survived clamping) — records
        /// what to go back to and clears the now-stale redo history.</summary>
        private void PushUndoSnapshot()
        {
            _undoStack.Push(new List<int>(_boundariesMs));
            if (_undoStack.Count > MaxUndoDepth)
            {
                var kept = _undoStack.ToArray(); // top (most recent) first
                _undoStack.Clear();
                for (int k = kept.Length - 2; k >= 0; k--) // drop the oldest one
                    _undoStack.Push(kept[k]);
            }
            _redoStack.Clear();
            UpdateUndoRedoButtons();
        }

        /// <summary>Wholesale-replacing _boundariesMs (new draft/track/fallback
        /// regeneration) makes prior undo/redo history meaningless — call this
        /// alongside clearing marker selection at those same points.</summary>
        private void ResetUndoHistory()
        {
            _undoStack.Clear();
            _redoStack.Clear();
            UpdateUndoRedoButtons();
        }

        private void Undo()
        {
            if (_undoStack.Count == 0) return;
            _redoStack.Push(new List<int>(_boundariesMs));
            _boundariesMs = _undoStack.Pop();
            _selectedMarkerIndices.Clear();
            _selectionAnchorIndex = null;
            RebuildMarkers();
            UpdateUndoRedoButtons();
        }

        private void Redo()
        {
            if (_redoStack.Count == 0) return;
            _undoStack.Push(new List<int>(_boundariesMs));
            _boundariesMs = _redoStack.Pop();
            _selectedMarkerIndices.Clear();
            _selectionAnchorIndex = null;
            RebuildMarkers();
            UpdateUndoRedoButtons();
        }

        private void UndoButton_Click(object sender, RoutedEventArgs e) => Undo();
        private void RedoButton_Click(object sender, RoutedEventArgs e) => Redo();

        private void UpdateUndoRedoButtons()
        {
            UndoButton.IsEnabled = _undoStack.Count > 0;
            RedoButton.IsEnabled = _redoStack.Count > 0;
        }

        private void ZoomInButton_Click(object sender, RoutedEventArgs e) => ChangeZoom(1.25);
        private void ZoomOutButton_Click(object sender, RoutedEventArgs e) => ChangeZoom(1 / 1.25);

        private void ChangeZoom(double factor)
        {
            if (_audio == null) return;
            _pixelsPerSecond = Math.Max(MinPixelsPerSecond, Math.Min(MaxPixelsPerSecond, _pixelsPerSecond * factor));
            RebuildWaveform();
        }

        private void WaveHostGrid_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (_playbackReader == null) return;
            double x = e.GetPosition(WaveHostGrid).X;
            int ms = PixelsToMs(x);
            SeekTo(ms);
        }

        // =====================================================================
        // Word panel
        // =====================================================================

        private void RebuildWordPanel()
        {
            WordsItemsControl.Items.Clear();
            _wordChips.Clear();
            _currentWordIndex = -1;

            for (int i = 0; i < _words.Count; i++)
            {
                double score = (i < _scores.Count) ? _scores[i] : 0;
                var border = new Border
                {
                    Margin = new Thickness(2),
                    Padding = new Thickness(5, 2, 5, 2),
                    CornerRadius = new CornerRadius(3),
                    Background = ScoreToBrush(score),
                    BorderThickness = new Thickness(1),
                    BorderBrush = Brushes.Transparent,
                    Cursor = System.Windows.Input.Cursors.Hand,
                    Tag = i,
                    ToolTip = score == 0 ? "no confidence score (manual)" : $"confidence score: {score:0.00}",
                    Child = new TextBlock { Text = _words[i] },
                };
                border.MouseLeftButtonDown += (s, e) =>
                {
                    int idx = (int)((Border)s!).Tag;
                    SeekTo(_boundariesMs[idx]);
                };
                WordsItemsControl.Items.Add(border);
                _wordChips.Add(border);
            }
        }

        private static SolidColorBrush ScoreToBrush(double score)
        {
            if (score == 0) return new SolidColorBrush(Color.FromRgb(0xE8, 0xE8, 0xE8));   // unknown / manual
            if (score >= -0.8) return new SolidColorBrush(Color.FromRgb(0xC8, 0xF0, 0xC8)); // confident
            if (score >= -1.8) return new SolidColorBrush(Color.FromRgb(0xFA, 0xE8, 0xB0)); // check it
            return new SolidColorBrush(Color.FromRgb(0xF6, 0xC6, 0xC6));                    // likely wrong
        }

        private void HighlightCurrentWord(int index)
        {
            if (_currentWordIndex == index) return;
            if (_currentWordIndex >= 0 && _currentWordIndex < _wordChips.Count)
                _wordChips[_currentWordIndex].BorderBrush = Brushes.Transparent;
            if (index >= 0 && index < _wordChips.Count)
            {
                _wordChips[index].BorderBrush = Brushes.DodgerBlue;
                _wordChips[index].BringIntoView();
            }
            _currentWordIndex = index;
        }

        // =====================================================================
        // Playback
        // =====================================================================

        private void PlayPauseButton_Click(object sender, RoutedEventArgs e)
        {
            if (_waveOut == null) return;
            if (_waveOut.PlaybackState == PlaybackState.Playing)
            {
                _waveOut.Pause();
                PlayPauseButton.Content = "Play";
                _playbackTimer.Stop();
            }
            else
            {
                _waveOut.Play();
                PlayPauseButton.Content = "Pause";
                _playbackTimer.Start();
            }
        }

        private void SeekTo(int ms)
        {
            if (_playbackReader == null || _audio == null) return;
            ms = Math.Max(0, Math.Min(ms, (int)(_audio.DurationSeconds * 1000)));
            _playbackReader.CurrentTime = TimeSpan.FromMilliseconds(ms);
            _soundTouch?.Clear(); // flush internal time-stretch buffers so playback doesn't resume mid-stale-chunk
            UpdatePlayheadPosition(ms);
        }

        private void SpeedComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e) =>
            ApplyCurrentSpeedSelection();

        /// <summary>Reads the "1.25x" style text of the selected ComboBoxItem and applies
        /// it as SoundTouch's Tempo (speed changes, pitch stays put — no chipmunk effect).</summary>
        private void ApplyCurrentSpeedSelection()
        {
            if (_soundTouch == null) return;
            if (SpeedComboBox.SelectedItem is ComboBoxItem item && item.Content is string label &&
                double.TryParse(label.TrimEnd('x', 'X'), NumberStyles.Float, CultureInfo.InvariantCulture, out double speed))
            {
                _soundTouch.Tempo = speed;
            }
        }

        private void PlaybackTimer_Tick()
        {
            if (_playbackReader == null) return;
            int ms = (int)_playbackReader.CurrentTime.TotalMilliseconds;
            UpdatePlayheadPosition(ms);
        }

        private void UpdatePlayheadPosition(int ms)
        {
            double x = MsToPixels(ms);
            PlayheadLine.X1 = x;
            PlayheadLine.X2 = x;

            // keep the playhead in view while playing
            double viewLeft = WaveScrollViewer.HorizontalOffset;
            double viewRight = viewLeft + WaveScrollViewer.ViewportWidth;
            if (x < viewLeft || x > viewRight - 20)
                WaveScrollViewer.ScrollToHorizontalOffset(Math.Max(0, x - 40));

            int idx = FindWordIndexAt(ms);
            HighlightCurrentWord(idx);
        }

        private int FindWordIndexAt(int ms)
        {
            if (_boundariesMs.Count < 2) return -1;
            int lo = 0, hi = _boundariesMs.Count - 2;
            if (ms < _boundariesMs[0] || ms > _boundariesMs[^1]) return -1;
            while (lo <= hi)
            {
                int mid = (lo + hi) / 2;
                if (ms < _boundariesMs[mid]) hi = mid - 1;
                else if (ms >= _boundariesMs[mid + 1]) lo = mid + 1;
                else return mid;
            }
            return -1;
        }

        private void CleanupPlayback()
        {
            _playbackTimer.Stop();
            _waveOut?.Stop();
            _waveOut?.Dispose();
            _waveOut = null;
            _soundTouch = null; // no Dispose exposed — it holds no unmanaged resources of its own
            _playbackReader?.Dispose();
            _playbackReader = null;
        }

        // =====================================================================
        // Export
        // =====================================================================

        private void CopyMarkersButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string text = MarkerTextBuilder.Build(_originalText, _words, _boundariesMs);
                Clipboard.SetText(text);
                StatusText.Text = "Marker text copied to clipboard — paste into the Word table row.";
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Export failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveMarkersButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string text = MarkerTextBuilder.Build(_originalText, _words, _boundariesMs);
                var dlg = new SaveFileDialog { Filter = "Text files|*.txt|All files|*.*", FileName = "markers.txt" };
                if (dlg.ShowDialog() == true)
                {
                    File.WriteAllText(dlg.FileName, text);
                    StatusText.Text = $"Saved to {dlg.FileName}";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Export failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 🌟 مارکرهای همین ترک را در سندِ مقصد (پیش‌فرض ‎«نامِ سند»_markers.docx) می‌نویسد و بقیه‌ی
        /// جدول‌ها و مارکرهایشان را دست نمی‌زند؛ پس یک سندِ _markers ترک‌به‌ترک کامل می‌شود و دیگر
        /// لازم نیست خروجی‌ها دستی در هم ادغام شوند. Shift+کلیک: انتخابِ سندِ مقصدِ دیگر.
        /// </summary>
        private void ExportDocxButton_Click(object sender, RoutedEventArgs e)
        {
            if (_docxSourcePath == null || _docxTrackIndex < 0)
            {
                MessageBox.Show(this, "Open a Word document (not a .txt file) first — markers are written into " +
                    "the track you're reviewing, in that document's _markers copy.",
                    "Nothing to export", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            try
            {
                string target = _docxTargetPath ?? DefaultTargetFor(_docxSourcePath);
                if ((Keyboard.Modifiers & ModifierKeys.Shift) != 0)
                {
                    var dlg = new SaveFileDialog
                    {
                        Title = "Markers document to update (its other tables are kept)",
                        Filter = "Word documents|*.docx",
                        FileName = Path.GetFileName(target),
                        InitialDirectory = Path.GetDirectoryName(target),
                        OverwritePrompt = false,   // به‌روزرسانی است، نه جایگزینی
                    };
                    if (dlg.ShowDialog() != true) return;
                    target = dlg.FileName;
                }

                string trackName = _docxTracks[_docxTrackIndex].AudioTrackName;
                var (targetIndex, created) = DocxScriptService.WriteMarkers(_docxSourcePath, target,
                    _docxTrackIndex, trackName, _words, _boundariesMs);

                _docxTargetPath = target;
                if (TargetIsSource)
                {
                    // خودِ سندِ باز‌شده عوض شد؛ فهرستِ ترک‌ها از نو خوانده می‌شود (متنِ تمیز همان است)
                    _docxTracks = DocxScriptService.ExtractTracksShared(_docxSourcePath);
                }
                RefreshTargetTracks();
                // حالا مارکرهای این ترک در سندِ مقصد ثبت شده‌اند؛ پیش‌نویسِ خودکار دیگر رویشان نمی‌نشیند
                _reviewedMarkersLoaded = true;

                UpdateStatus();
                StatusText.Text = $"{(created ? "Created" : "Updated")} {Path.GetFileName(target)}: table {targetIndex + 1} " +
                                  $"\"{trackName}\" now has {_boundariesMs.Count} markers. Other tables untouched" +
                                  (created ? "." : $"; previous version kept as {Path.GetFileName(target)}.bak.");
            }
            catch (DocumentLockedException ex)
            {
                MessageBox.Show(this, ex.Message, "Document is open", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Export failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UpdateStatus()
        {
            bool ready = _audio != null && _words.Count > 0 && !string.IsNullOrEmpty(_originalText);
            CopyMarkersButton.IsEnabled = ready;
            SaveMarkersButton.IsEnabled = ready;
            ExportDocxButton.IsEnabled = ready && _docxSourcePath != null;

            int lowConf = _scores.FindAll(s => s < -1.8).Count;
            string sourceInfo = _docxSourcePath != null && _docxTrackIndex >= 0 && _docxTrackIndex < _docxTracks.Count
                ? $"Word doc track {_docxTrackIndex + 1}/{_docxTracks.Count}: {_docxTracks[_docxTrackIndex].AudioTrackName}" +
                  (string.IsNullOrEmpty(_trackMatchNote) ? "" : $" [{_trackMatchNote}]") + "  |  "
                : "";
            StatusText.Text = _audio == null
                ? "No files loaded yet."
                : $"{sourceInfo}Audio: {Path.GetFileName(_audioPath)} ({_audio.DurationSeconds:0.0}s)  |  " +
                  $"Words: {_words.Count}  |  Low-confidence: {lowConf}" +
                  (string.IsNullOrEmpty(_draftNote) ? "" : $"  |  {_draftNote}") +
                  (string.IsNullOrEmpty(_targetNote) ? "" : $"  |  {_targetNote}");
        }
    }

    /// <summary>Tiny wrapper so MainWindow doesn't need a XAML-declared DispatcherTimer field.</summary>
    internal sealed class DispatcherTimerWrapper
    {
        private readonly System.Windows.Threading.DispatcherTimer _timer;

        public DispatcherTimerWrapper(TimeSpan interval, Action onTick)
        {
            _timer = new System.Windows.Threading.DispatcherTimer { Interval = interval };
            _timer.Tick += (_, _) => onTick();
        }

        public void Start() => _timer.Start();
        public void Stop() => _timer.Stop();
    }
}