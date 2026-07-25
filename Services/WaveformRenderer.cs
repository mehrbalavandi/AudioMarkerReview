using System;
using System.Windows;
using System.Windows.Media;
using NAudio.Wave;

namespace AudioMarkerReview.Services
{
    public class DecodedAudio
    {
        public float[] Samples { get; set; } = Array.Empty<float>(); // mono, [-1, 1]
        public int SampleRate { get; set; }
        public double DurationSeconds { get; set; }
    }

    public static class AudioLoader
    {
        /// <summary>
        /// Decodes an audio file (mp3/wav/whatever NAudio's default readers support on
        /// Windows) to a mono float array entirely in memory. Fine for the few-minutes-
        /// long listening-passage files this tool targets; not meant for hour-long audio.
        /// </summary>
        public static DecodedAudio Load(string path)
        {
            using var reader = new AudioFileReader(path);
            int sampleRate = reader.WaveFormat.SampleRate;
            int channels = reader.WaveFormat.Channels;

            var buffer = new float[reader.Length / 4]; // 4 bytes per float sample frame slot (upper bound)
            int read = reader.Read(buffer, 0, buffer.Length);

            float[] mono;
            if (channels <= 1)
            {
                mono = new float[read];
                Array.Copy(buffer, mono, read);
            }
            else
            {
                int frames = read / channels;
                mono = new float[frames];
                for (int i = 0; i < frames; i++)
                {
                    float sum = 0f;
                    for (int c = 0; c < channels; c++)
                        sum += buffer[i * channels + c];
                    mono[i] = sum / channels;
                }
            }

            return new DecodedAudio
            {
                Samples = mono,
                SampleRate = sampleRate,
                DurationSeconds = mono.Length / (double)sampleRate,
            };
        }
    }

    public static class WaveformRenderer
    {
        /// <summary>
        /// Builds a single filled Path geometry showing min/max envelope peaks, sized
        /// totalWidthPixels x height. Caller sets totalWidthPixels = duration * pixelsPerSecond
        /// so it lines up with the marker/playhead overlay using the same pixels-per-second.
        /// </summary>
        public static PathGeometry BuildPeaksGeometry(float[] samples, double totalWidthPixels, double height)
        {
            int width = Math.Max(1, (int)Math.Round(totalWidthPixels));
            int samplesPerPixel = Math.Max(1, samples.Length / width);
            double mid = height / 2.0;

            var top = new PointCollection(width);
            var bottom = new PointCollection(width);

            for (int px = 0; px < width; px++)
            {
                int start = px * samplesPerPixel;
                int end = Math.Min(samples.Length, start + samplesPerPixel);
                float min = 0f, max = 0f;
                for (int i = start; i < end; i++)
                {
                    float v = samples[i];
                    if (v > max) max = v;
                    if (v < min) min = v;
                }
                top.Add(new Point(px, mid - max * mid));
                bottom.Add(new Point(px, mid - min * mid));
            }

            var figure = new PathFigure { StartPoint = top.Count > 0 ? top[0] : new Point(0, mid) };
            for (int i = 1; i < top.Count; i++)
                figure.Segments.Add(new LineSegment(top[i], false));
            for (int i = bottom.Count - 1; i >= 0; i--)
                figure.Segments.Add(new LineSegment(bottom[i], false));
            figure.IsClosed = true;
            figure.IsFilled = true;

            var geometry = new PathGeometry();
            geometry.Figures.Add(figure);
            return geometry;
        }
    }
}
