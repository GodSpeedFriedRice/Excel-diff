using System.Xml.Linq;

namespace ExcelDiff.Core.Comparison;

public enum XmlLineKind { Context, Removed, Added }

public sealed record XmlDiffLine(XmlLineKind Kind, int? Before, int? After, string Text);

public sealed record XmlLineDiff(IReadOnlyList<XmlDiffLine> Lines, bool HasUnalignedBlocks)
{
    public static XmlLineDiff Compare(XDocument? before, XDocument? after, CancellationToken token = default)
    {
        token.ThrowIfCancellationRequested();
        var left = LinesOf(before);
        token.ThrowIfCancellationRequested();
        var right = LinesOf(after);
        var lines = new List<XmlDiffLine>();
        var budget = 4_000_000L;
        var unaligned = false;
        CompareRange(0, left.Length, 0, right.Length, 0);
        return new XmlLineDiff(lines, unaligned);

        void Emit(XmlLineKind kind, int? l, int? r) =>
            lines.Add(new XmlDiffLine(kind, l + 1, r + 1, l is int i ? left[i] : right[r!.Value]));

        void CompareRange(int l, int le, int r, int re, int depth)
        {
            token.ThrowIfCancellationRequested();
            while (l < le && r < re && left[l] == right[r])
            {
                Emit(XmlLineKind.Context, l++, r++);
                if ((l & 4095) == 0) token.ThrowIfCancellationRequested();
            }
            var suffix = 0;
            while (l < le - suffix && r < re - suffix && left[le - suffix - 1] == right[re - suffix - 1])
                suffix++;
            le -= suffix;
            re -= suffix;

            if (l == le || r == re)
                EmitBlock(l, le, r, re);
            else if ((long)(le - l + 1) * (re - r + 1) is var size && size <= 500_000 && size <= budget)
            {
                budget -= size;
                CompareSmall(l, le, r, re);
            }
            else
            {
                var anchors = depth < 16 ? UniqueAnchors(left, l, le, right, r, re, token) : [];
                if (anchors.Length == 0)
                {
                    unaligned = true;
                    EmitBlock(l, le, r, re);
                }
                else
                {
                    foreach (var (li, ri) in anchors)
                    {
                        CompareRange(l, li, r, ri, depth + 1);
                        Emit(XmlLineKind.Context, li, ri);
                        l = li + 1;
                        r = ri + 1;
                    }
                    CompareRange(l, le, r, re, depth + 1);
                }
            }
            for (var i = 0; i < suffix; i++) Emit(XmlLineKind.Context, le + i, re + i);
        }

        void EmitBlock(int l, int le, int r, int re)
        {
            for (; l < le; l++)
            {
                if ((l & 4095) == 0) token.ThrowIfCancellationRequested();
                Emit(XmlLineKind.Removed, l, null);
            }
            for (; r < re; r++)
            {
                if ((r & 4095) == 0) token.ThrowIfCancellationRequested();
                Emit(XmlLineKind.Added, null, r);
            }
        }

        void CompareSmall(int l, int le, int r, int re)
        {
            var height = le - l;
            var width = re - r;
            var lengths = new int[height + 1, width + 1];
            for (var i = height - 1; i >= 0; i--)
            {
                if ((i & 63) == 0) token.ThrowIfCancellationRequested();
                for (var j = width - 1; j >= 0; j--)
                    lengths[i, j] = left[l + i] == right[r + j] ? lengths[i + 1, j + 1] + 1 :
                        Math.Max(lengths[i + 1, j], lengths[i, j + 1]);
            }

            var x = 0;
            var y = 0;
            var startX = 0;
            var startY = 0;
            while (x < height && y < width)
            {
                if (left[l + x] == right[r + y])
                {
                    EmitBlock(l + startX, l + x, r + startY, r + y);
                    Emit(XmlLineKind.Context, l + x++, r + y++);
                    startX = x;
                    startY = y;
                }
                else if (lengths[x + 1, y] >= lengths[x, y + 1]) x++;
                else y++;
            }
            EmitBlock(l + startX, le, r + startY, re);
        }
    }

    private static string[] LinesOf(XDocument? document) =>
        document?.ToString().Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n') ?? [];

    private static (int Left, int Right)[] UniqueAnchors(string[] left, int l, int le,
        string[] right, int r, int re, CancellationToken token)
    {
        static Dictionary<string, int> Positions(string[] lines, int start, int end, CancellationToken token)
        {
            var positions = new Dictionary<string, int>(StringComparer.Ordinal);
            for (var i = start; i < end; i++)
            {
                if ((i & 4095) == 0) token.ThrowIfCancellationRequested();
                if (!positions.TryAdd(lines[i], i)) positions[lines[i]] = -1;
            }
            return positions;
        }

        var lpos = Positions(left, l, le, token);
        var rpos = Positions(right, r, re, token);
        var pairs = new List<(int Left, int Right)>();
        for (var i = l; i < le; i++)
        {
            if ((i & 4095) == 0) token.ThrowIfCancellationRequested();
            if (lpos[left[i]] == i && rpos.TryGetValue(left[i], out var match) && match >= 0)
                pairs.Add((i, match));
        }

        // Keep only anchors that occur in the same order on both sides.
        var tails = new int[pairs.Count];
        var predecessor = new int[pairs.Count];
        var count = 0;
        for (var i = 0; i < pairs.Count; i++)
        {
            var low = 0;
            var high = count;
            while (low < high)
            {
                var middle = (low + high) / 2;
                if (pairs[tails[middle]].Right < pairs[i].Right) low = middle + 1;
                else high = middle;
            }
            predecessor[i] = low == 0 ? -1 : tails[low - 1];
            tails[low] = i;
            if (low == count) count++;
        }
        var result = new (int, int)[count];
        if (count == 0) return result;
        var cursor = tails[count - 1];
        for (var i = count - 1; i >= 0; i--)
        {
            result[i] = pairs[cursor];
            cursor = predecessor[cursor];
        }
        return result;
    }
}
