using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AudioMarkerReview.Services
{
    /// <summary>
    /// 🌟 A-4b: شماره‌ای که Word برای هر بندِ لیست می‌کشد («1:»، «a)»، «iv.»، بولت). Word آن را در
    /// متن ذخیره نمی‌کند؛ از numbering.xml و شمارنده‌ها حساب می‌شود. نسخه‌ی کوچکِ NumberingResolverِ
    /// CustomLayoutGenerator (همان قاعده‌ها: شمارنده به‌ازای numId، startOverride، از نو شروع شدنِ
    /// سطوحِ عمیق‌تر، قالبِ lvlText مثلِ «%1:» و numFmt) و docx_tracks.ListNumberingِ align_engine.
    /// برای هر پاراگراف به ترتیبِ سند <see cref="Next"/> صدا زده می‌شود.
    /// </summary>
    public sealed class ListNumbering
    {
        private readonly Dictionary<int, AbstractNum> _absByNum = new Dictionary<int, AbstractNum>();
        private readonly Dictionary<int, Dictionary<int, int>> _startOverride = new Dictionary<int, Dictionary<int, int>>();
        private readonly Dictionary<int, Dictionary<int, int>> _counters = new Dictionary<int, Dictionary<int, int>>();
        private readonly Dictionary<string, int> _styleNum = new Dictionary<string, int>();

        public ListNumbering(MainDocumentPart mainPart)
        {
            var styles = mainPart.StyleDefinitionsPart?.Styles;
            if (styles != null)
                foreach (var st in styles.Elements<Style>())
                {
                    int? num = st.StyleParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value;
                    if (num != null && st.StyleId?.Value != null) _styleNum[st.StyleId.Value] = num.Value;
                }

            var numbering = mainPart.NumberingDefinitionsPart?.Numbering;
            if (numbering == null) return;
            var absById = new Dictionary<int, AbstractNum>();
            foreach (var a in numbering.Elements<AbstractNum>())
                if (a.AbstractNumberId?.Value is int id) absById[id] = a;

            foreach (var num in numbering.Elements<NumberingInstance>())
            {
                if (num.NumberID?.Value is not int numId) continue;
                var absRef = num.GetFirstChild<AbstractNumId>();
                if (absRef?.Val?.Value is int absId && absById.TryGetValue(absId, out var abs))
                    _absByNum[numId] = abs;
                foreach (var ov in num.Elements<LevelOverride>())
                {
                    var start = ov.StartOverrideNumberingValue?.Val?.Value;
                    if (ov.LevelIndex?.Value is int lvl && start.HasValue)
                    {
                        if (!_startOverride.TryGetValue(numId, out var m))
                            _startOverride[numId] = m = new Dictionary<int, int>();
                        m[lvl] = start.Value;
                    }
                }
            }
        }

        /// <summary>شمارنده را جلو می‌برد و شماره‌ی این بند را برمی‌گرداند؛ null اگر بندِ لیست نیست.</summary>
        public string? Next(Paragraph p)
        {
            var numPr = p.ParagraphProperties?.NumberingProperties;
            int? numId = numPr?.NumberingId?.Val?.Value;
            int level = numPr?.NumberingLevelReference?.Val?.Value ?? 0;
            if (numId == null)
            {
                string? styleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                if (styleId != null && _styleNum.TryGetValue(styleId, out int sn)) numId = sn;
            }
            if (numId == null || numId == 0) return null;
            if (!_absByNum.TryGetValue(numId.Value, out var abs)) return null;

            if (!_counters.TryGetValue(numId.Value, out var counters))
                _counters[numId.Value] = counters = new Dictionary<int, int>();
            if (!counters.ContainsKey(level))
                counters[level] = StartOf(numId.Value, abs, level) - 1;
            counters[level]++;
            foreach (int deeper in counters.Keys.Where(k => k > level).ToList())
                counters[deeper] = StartOf(numId.Value, abs, deeper) - 1;

            var lvlDef = LevelOf(abs, level);
            var fmt = lvlDef?.NumberingFormat?.Val?.Value ?? NumberFormatValues.Decimal;
            if (fmt == NumberFormatValues.Bullet) return "•";

            string template = lvlDef?.LevelText?.Val?.Value ?? $"%{level + 1}.";
            var sb = new StringBuilder();
            for (int i = 0; i < template.Length; i++)
            {
                if (template[i] == '%' && i + 1 < template.Length && char.IsDigit(template[i + 1]))
                {
                    int refLevel = (template[i + 1] - '0') - 1;
                    int val = counters.TryGetValue(refLevel, out var v) ? v : StartOf(numId.Value, abs, refLevel);
                    var f = LevelOf(abs, refLevel)?.NumberingFormat?.Val?.Value ?? NumberFormatValues.Decimal;
                    sb.Append(Format(val, f));
                    i++;
                }
                else sb.Append(template[i]);
            }
            return sb.ToString();
        }

        private int StartOf(int numId, AbstractNum abs, int level)
        {
            if (_startOverride.TryGetValue(numId, out var m) && m.TryGetValue(level, out var s)) return s;
            return LevelOf(abs, level)?.StartNumberingValue?.Val?.Value ?? 1;
        }

        private static Level? LevelOf(AbstractNum abs, int level) =>
            abs.Elements<Level>().FirstOrDefault(l => l.LevelIndex?.Value == level);

        private static string Format(int n, NumberFormatValues fmt)
        {
            if (n < 1) n = 1;
            if (fmt == NumberFormatValues.LowerLetter || fmt == NumberFormatValues.UpperLetter)
            {
                var sb = new StringBuilder();
                while (n > 0) { n--; sb.Insert(0, (char)('a' + (n % 26))); n /= 26; }
                return fmt == NumberFormatValues.UpperLetter ? sb.ToString().ToUpperInvariant() : sb.ToString();
            }
            if (fmt == NumberFormatValues.LowerRoman || fmt == NumberFormatValues.UpperRoman)
            {
                int[] vals = { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
                string[] syms = { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
                var sb = new StringBuilder();
                for (int i = 0; i < vals.Length && n > 0; i++)
                    while (n >= vals[i]) { sb.Append(syms[i]); n -= vals[i]; }
                return fmt == NumberFormatValues.LowerRoman ? sb.ToString().ToLowerInvariant() : sb.ToString();
            }
            return n.ToString();
        }
    }
}
