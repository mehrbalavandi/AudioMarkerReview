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

        public MainWindow()
        {
            InitializeComponent();
            _playbackTimer = new DispatcherTimerWrapper(TimeSpan.FromMilliseconds(50), PlaybackTimer_Tick);
            Closing += (_, _) => CleanupPlayback();
        }

        // =====================================================================
        // File loading
        // =====================================================================

        private void OpenAudioButton_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { Filter = "Audio files|*.mp3;*.wav;*.m4a;*.flac|All files|*.*" };
            if (dlg.ShowDialog() != true) return;

            try
            {
                CleanupPlayback();
                _audio = AudioLoader.Load(dlg.FileName);
                _audioPath = dlg.FileName;

                _playbackReader = new AudioFileReader(dlg.FileName);
                _soundTouch = new SoundTouchWaveProvider(_playbackReader);
                _soundTouch.OptimizeForSpeech();
                ApplyCurrentSpeedSelection();

                _waveOut = new WaveOutEvent();
                _waveOut.Init(_soundTouch);
                _waveOut.PlaybackStopped += (_, _) => Dispatcher.Invoke(() => PlayPauseButton.Content = "Play");

                PlayPauseButton.IsEnabled = true;
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
                _suppressTrackSelectionEvent = true;
                TrackComboBox.ItemsSource = null;
                TrackComboBox.IsEnabled = false;
                _suppressTrackSelectionEvent = false;

                LoadOriginalText(File.ReadAllText(dlg.FileName));
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

            try
            {
                var tracks = DocxScriptService.ExtractTracks(dlg.FileName);
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

                _suppressTrackSelectionEvent = true;
                TrackComboBox.ItemsSource = _docxTracks;
                TrackComboBox.SelectedIndex = 0;
                TrackComboBox.IsEnabled = true;
                _suppressTrackSelectionEvent = false;

                LoadDocxTrack(0);
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
            LoadDocxTrack(TrackComboBox.SelectedIndex);
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
            LoadOriginalText(_docxTracks[index].FullText);
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
            if (dlg.ShowDialog() != true) return;

            try
            {
                var json = File.ReadAllText(dlg.FileName);
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
                _selectedMarkerIndices.Clear();
                _selectionAnchorIndex = null;

                RebuildWaveform();
                RebuildWordPanel();
                UpdateStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Could not load draft JSON:\n{ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
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

            foreach (int idx in _selectedMarkerIndices)
                _boundariesMs[idx] += delta;

            RebuildMarkers(); // repaints positions; selection itself is untouched
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

        private void ExportDocxButton_Click(object sender, RoutedEventArgs e)
        {
            if (_docxSourcePath == null || _docxTrackIndex < 0)
            {
                MessageBox.Show(this, "Open a Word document (not a .txt file) first — this exports a new " +
                    "copy of that same document with markers spliced into the track you're reviewing.",
                    "Nothing to export", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            try
            {
                var dlg = new SaveFileDialog
                {
                    Filter = "Word documents|*.docx",
                    FileName = Path.GetFileNameWithoutExtension(_docxSourcePath) + "_markers.docx",
                    InitialDirectory = Path.GetDirectoryName(_docxSourcePath),
                };
                if (dlg.ShowDialog() != true) return;

                DocxScriptService.ExportWithMarkers(_docxSourcePath, _docxTrackIndex, _words, _boundariesMs, dlg.FileName);
                StatusText.Text = $"Saved marked-up Word document to {dlg.FileName} " +
                                   "(styles/formatting carried over untouched — only markers were added).";
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
                ? $"Word doc track: {_docxTracks[_docxTrackIndex].AudioTrackName}  |  "
                : "";
            StatusText.Text = _audio == null
                ? "No files loaded yet."
                : $"{sourceInfo}Audio: {Path.GetFileName(_audioPath)} ({_audio.DurationSeconds:0.0}s)  |  " +
                  $"Words: {_words.Count}  |  Low-confidence: {lowConf}";
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
