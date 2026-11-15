namespace ExcelDiff.Core.Comparison;

public sealed record ThreeWayXmlLine(XmlDiffLine? Base, XmlDiffLine? Production, XmlDiffLine? Incoming)
{
    public bool Changed => Base?.Kind != XmlLineKind.Context || Production?.Kind != XmlLineKind.Context ||
        Incoming?.Kind != XmlLineKind.Context;
}

public static class ThreeWayXmlDiff
{
    public static IReadOnlyList<ThreeWayXmlLine> Align(XmlLineDiff production, XmlLineDiff incoming,
        CancellationToken token = default)
    {
        var first = Index(production, token);
        var second = Index(incoming, token);
        var rows = new List<ThreeWayXmlLine>();
        AppendInsertions(0);
        for (var number = 1; number <= first.Lines.Count; number++)
        {
            token.ThrowIfCancellationRequested();
            var ours = first.Lines[number];
            var theirs = second.Lines[number];
            var original = ours.Original;
            if (theirs.Original.Kind == XmlLineKind.Removed) original = theirs.Original;
            rows.Add(new(original, ours.Version, theirs.Version));
            AppendInsertions(number);
        }
        return rows;

        void AppendInsertions(int afterLine)
        {
            var ours = first.Insertions.GetValueOrDefault(afterLine);
            var theirs = second.Insertions.GetValueOrDefault(afterLine);
            for (var i = 0; i < Math.Max(ours?.Count ?? 0, theirs?.Count ?? 0); i++)
            {
                token.ThrowIfCancellationRequested();
                rows.Add(new(null, ours is not null && i < ours.Count ? ours[i] : null,
                    theirs is not null && i < theirs.Count ? theirs[i] : null));
            }
        }
    }

    private static Alignment Index(XmlLineDiff diff, CancellationToken token)
    {
        var result = new Alignment();
        var removed = new List<XmlDiffLine>();
        var added = new List<XmlDiffLine>();
        var previous = 0;
        foreach (var line in diff.Lines)
        {
            token.ThrowIfCancellationRequested();
            switch (line.Kind)
            {
                case XmlLineKind.Context:
                    Flush();
                    previous = line.Before!.Value;
                    result.Lines.Add(previous, (line, line));
                    break;
                case XmlLineKind.Removed:
                    removed.Add(line);
                    break;
                case XmlLineKind.Added:
                    added.Add(line);
                    break;
            }
        }
        Flush();
        return result;

        void Flush()
        {
            for (var i = 0; i < removed.Count; i++)
            {
                var old = removed[i];
                previous = old.Before!.Value;
                result.Lines.Add(previous, (old, i < added.Count ? added[i] : null));
            }
            if (added.Count > removed.Count)
                result.Insertions.Add(previous, added.GetRange(removed.Count, added.Count - removed.Count));
            removed.Clear();
            added.Clear();
        }
    }

    private sealed class Alignment
    {
        internal Dictionary<int, (XmlDiffLine Original, XmlDiffLine? Version)> Lines { get; } = [];
        internal Dictionary<int, List<XmlDiffLine>> Insertions { get; } = [];
    }
}
