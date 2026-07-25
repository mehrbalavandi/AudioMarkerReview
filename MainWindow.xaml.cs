using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using Microsoft.Win32;
using NAudio.Wave;
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

        // ---- playback ---------------------------------------------------------
        private AudioFileReader? _playbackReader;
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
                _waveOut = new WaveOutEvent();
                _waveOut.Init(_playbackReader);
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
                _originalText = File.ReadAllText(dlg.FileName);
                // If a draft JSON was already loaded, verify the text still matches it;
                // otherwise this becomes manual mode using this text's own word list.
                if (_words.Count > 0)
                {
                    var textWords = MarkerTextBuilder.SplitWords(_originalText);
                    if (textWords.Count != _words.Count)
                    {
                        MessageBox.Show(this,
                            $"This text file has {textWords.Count} words but the loaded draft JSON has " +
                            $"{_words.Count}. Load the matching draft JSON, or load this text first and " +
                            "skip the draft to start from evenly-spaced markers instead.",
                            "Word count mismatch", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                EnsureBoundariesInitialized();
                RebuildWaveform();
                RebuildWordPanel();
                UpdateStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Could not load text:\n{ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
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
                thumb.DragDelta += MarkerThumb_DragDelta;
                MarkerCanvas.Children.Add(thumb);
                _thumbs.Add(thumb);
            }
        }

        private void MarkerThumb_DragDelta(object sender, DragDeltaEventArgs e)
        {
            var thumb = (Thumb)sender;
            int i = (int)thumb.Tag;

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
            UpdatePlayheadPosition(ms);
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

        private void UpdateStatus()
        {
            bool ready = _audio != null && _words.Count > 0 && !string.IsNullOrEmpty(_originalText);
            CopyMarkersButton.IsEnabled = ready;
            SaveMarkersButton.IsEnabled = ready;

            int lowConf = _scores.FindAll(s => s < -1.8).Count;
            StatusText.Text = _audio == null
                ? "No files loaded yet."
                : $"Audio: {Path.GetFileName(_audioPath)} ({_audio.DurationSeconds:0.0}s)  |  " +
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
