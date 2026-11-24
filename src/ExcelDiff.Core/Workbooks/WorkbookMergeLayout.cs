namespace ExcelDiff.Core.Workbooks;

public static partial class WorkbookMerger
{
    private static MergeIndex AddMergeEdits(Sheet? baselineSheet, Sheet productionSheet, Sheet incomingSheet,
        MergeWorksheet? baseline, MergeWorksheet production, MergeWorksheet incoming,
        List<WorkbookMergePlan.Edit> edits, CancellationToken token)
    {
        var original = baselineSheet?.Merges.Ranges ?? [];
        var ours = productionSheet.Merges.Ranges;
        var theirs = incomingSheet.Merges.Ranges;
        if (ours.SequenceEqual(theirs) && (baselineSheet is null || original.SequenceEqual(ours))) return new([]);

        var ranges = original.Concat(ours).Concat(theirs).Distinct().ToArray();
        var originalSet = original.ToHashSet();
        var ourSet = ours.ToHashSet();
        var theirSet = theirs.ToHashSet();
        var index = new MergeIndex(ranges);
        var visited = new HashSet<MergeRange>();
        var covered = new List<MergeRange>();
        var beforeVersion = baselineSheet is null || baseline is null ? null : new LayoutVersion(baseline);
        var currentVersion = new LayoutVersion(production);
        var incomingVersion = new LayoutVersion(incoming);
        foreach (var seed in index.Ranges)
        {
            token.ThrowIfCancellationRequested();
            if (!visited.Add(seed)) continue;
            var group = new List<MergeRange> { seed };
            for (var next = 0; next < group.Count; next++)
            {
                token.ThrowIfCancellationRequested();
                var range = group[next];
                foreach (var other in index.InRows(range.FirstRow, range.LastRow))
                    if (OverlapsColumns(range, other) && visited.Add(other)) group.Add(other);
            }

            var scope = new MergeIndex(group);
            var oursInScope = scope.Ranges.Where(ourSet.Contains).ToArray();
            var theirsInScope = scope.Ranges.Where(theirSet.Contains).ToArray();
            var originalInScope = scope.Ranges.Where(originalSet.Contains).ToArray();
            if (oursInScope.SequenceEqual(theirsInScope) &&
                (baselineSheet is null || originalInScope.SequenceEqual(oursInScope))) continue;

            covered.AddRange(scope.Ranges);
            var before = beforeVersion?.Read(scope, originalInScope, token);
            var current = currentVersion.Read(scope, oursInScope, token);
            var source = incomingVersion.Read(scope, theirsInScope, token);
            var allowsChoice = current.AllowsChoice && source.AllowsChoice;
            MergeDisposition disposition;
            if (before is not null && source.Matches(before)) disposition = MergeDisposition.KeepProduction;
            else if (current.Matches(source)) disposition = MergeDisposition.AlreadyMatches;
            else if (!allowsChoice) disposition = MergeDisposition.Unsupported;
            else disposition = before is not null && current.Matches(before) ? MergeDisposition.TakeIncoming : MergeDisposition.Conflict;

            var anchor = scope.Ranges[0];
            var reason = disposition == MergeDisposition.Unsupported
                ? "結合範囲に特殊な数式やセル内の文字装飾などがあります。Excel で解消してから再比較してください。"
                : "結合範囲と範囲内のセルの内容をまとめて採用します。";
            var change = new MergeCellChange(new(productionSheet.Name, anchor.FirstRow, anchor.FirstColumn),
                before?.Display ?? "", current.Display, source.Display, disposition, reason)
            {
                AllowsChoice = allowsChoice,
                MergeRanges = scope.Ranges
            };
            edits.Add(new(change, productionSheet.XmlPartPath!, MergeContent.Blank)
            {
                Layout = new(scope, source.Ranges, source.Cells)
            });
        }
        return new(covered);
    }

    private static bool OverlapsColumns(MergeRange left, MergeRange right) =>
        left.FirstColumn <= right.LastColumn && right.FirstColumn <= left.LastColumn;

    private sealed class LayoutVersion(MergeWorksheet worksheet)
    {
        private readonly int[] rows = worksheet.Rows.Keys.Order().ToArray();

        internal LayoutState Read(MergeIndex scope, MergeRange[] merges, CancellationToken token)
        {
            var cells = new Dictionary<(int Row, int Column), MergeContent>();
            var allowsChoice = true;
            foreach (var range in scope.Ranges)
            {
                token.ThrowIfCancellationRequested();
                var first = Array.BinarySearch(rows, range.FirstRow);
                if (first < 0) first = ~first;
                for (var i = first; i < rows.Length && rows[i] <= range.LastRow; i++)
                {
                    token.ThrowIfCancellationRequested();
                    foreach (var (column, cell) in worksheet.Rows[rows[i]])
                    {
                        if (!range.ContainsColumn(column)) continue;
                        cells[(rows[i], column)] = cell;
                        if (cell.Special || cell.Key.Kind == "formula" && ExternalBookReference().IsMatch(cell.Key.Text))
                            allowsChoice = false;
                    }
                }
                if (worksheet.FormulaRanges.InRows(range.FirstRow, range.LastRow).Any(other => OverlapsColumns(range, other)))
                    allowsChoice = false;
            }
            return new(merges, cells, allowsChoice);
        }
    }

    private sealed record LayoutState(MergeRange[] Ranges, Dictionary<(int Row, int Column), MergeContent> Cells, bool AllowsChoice)
    {
        internal bool Matches(LayoutState other) => Ranges.SequenceEqual(other.Ranges) && Cells.Count == other.Cells.Count &&
            Cells.All(cell => other.Cells.TryGetValue(cell.Key, out var content) && cell.Value.Key == content.Key);

        internal string Display
        {
            get
            {
                var merges = Ranges.Length == 0 ? "結合なし" : "結合 " + string.Join(" / ", Ranges.Select(range => range.ToString()));
                var contents = Cells.OrderBy(cell => cell.Key.Row).ThenBy(cell => cell.Key.Column)
                    .Select(cell => $"{Workbooks.Address.Format(cell.Key.Row, cell.Key.Column)}={cell.Value.Display}");
                return string.Join(Environment.NewLine, new[] { merges }.Concat(contents));
            }
        }
    }
}
