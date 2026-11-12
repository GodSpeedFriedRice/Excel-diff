using System.Globalization;
using ExcelDiff.Core.Workbooks;

namespace ExcelDiff.Desktop.Views;

internal sealed class PreviewAxis
{
    private readonly int defaultSize;
    private readonly int[] indices;
    private readonly long[] deltas;

    internal PreviewAxis(Sheet sheet, bool rows, double zoom)
    {
        Limit = rows ? 1048576 : 16384;
        var prefix = rows ? "行 " : "列 ";
        var sizeKey = rows ? "ht" : "width";
        var defaultKey = rows ? "defaultRowHeight" : "defaultColWidth";
        var defaultUnits = Units(sheet.Settings.GetValueOrDefault("既定レイアウト / " + defaultKey), rows ? 15 : 8.43);
        int Pixels(double units) => (int)Math.Round(Math.Clamp((rows ? units * 96 / 72 : units * 7 + 5) * zoom, 1, 4096));

        defaultSize = Pixels(defaultUnits);
        var settings = new Dictionary<int, (double Size, bool Hidden)>();
        foreach (var (key, value) in sheet.Settings)
        {
            if (!key.StartsWith(prefix, StringComparison.Ordinal)) continue;
            var separator = key.IndexOf(" / ", StringComparison.Ordinal);
            if (separator < 0) continue;
            var property = key.AsSpan(separator + 3);
            if (!property.SequenceEqual(sizeKey) && !property.SequenceEqual("hidden")) continue;
            var label = key.AsSpan(prefix.Length, separator - prefix.Length);
            if (!(rows ? int.TryParse(label, out var index) : Address.TryColumn(label, out index))) continue;
            var current = settings.GetValueOrDefault(index, (defaultUnits, false));
            settings[index] = property.SequenceEqual("hidden")
                ? (current.Item1, value is "1" or "true") : (Units(value, defaultUnits), current.Item2);
        }

        indices = settings.Keys.Order().ToArray();
        deltas = new long[indices.Length + 1];
        for (var i = 0; i < indices.Length; i++)
        {
            var (size, hidden) = settings[indices[i]];
            deltas[i + 1] = deltas[i] + (hidden || size == 0 ? 0 : Pixels(size)) - defaultSize;
        }
        LastSized = indices.LastOrDefault();
    }

    internal int Limit { get; }
    internal int LastSized { get; }

    internal long Position(int index)
    {
        var match = Array.BinarySearch(indices, index);
        return (long)(index - 1) * defaultSize + deltas[match < 0 ? ~match : match];
    }

    internal int Offset(int origin, int index, int header) =>
        (int)Math.Clamp(Position(index) - Position(origin) + header, -100000, 100000);

    internal int AtPixel(long position)
    {
        var low = 1;
        var high = Limit;
        while (low < high)
        {
            var middle = (low + high) / 2;
            if (Position(middle + 1) <= position) low = middle + 1;
            else high = middle;
        }
        return low;
    }

    internal int[] Visible(int first, int last, int pixels)
    {
        var visible = new List<int>();
        var end = Position(first) + pixels;
        for (var index = first; index <= last && Position(index) < end;)
        {
            if (Position(index + 1) > Position(index))
            {
                visible.Add(index++);
                continue;
            }

            // Jump past contiguous hidden rows or columns without walking the range.
            var low = index + 1;
            var high = last + 1;
            var start = Position(index);
            while (low < high)
            {
                var middle = (low + high) / 2;
                if (Position(middle + 1) <= start) low = middle + 1;
                else high = middle;
            }
            index = low;
        }
        return visible.ToArray();
    }

    private static double Units(string? value, double fallback) =>
        double.TryParse(value, CultureInfo.InvariantCulture, out var parsed) && double.IsFinite(parsed) ? parsed : fallback;
}
