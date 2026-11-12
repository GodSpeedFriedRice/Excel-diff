using ExcelDiff.Core.Workbooks;

namespace ExcelDiff.Core.Comparison;

public enum ChangeKind { Value, Formula, Format, Setting, Added, Removed, Merge }

public sealed record DifferencePart(string Item, string? Left, string? Right);

public sealed record Difference(ChangeKind Kind, string Item, string? Left, string? Right,
    int? LeftRow = null, int? RightRow = null, int? Column = null,
    MergeRange? LeftMerge = null, MergeRange? RightMerge = null, IReadOnlyList<DifferencePart>? Parts = null)
{
    public string LeftAddress => LeftMerge?.ToString() ?? Address.Format(LeftRow, Column);
    public string RightAddress => RightMerge?.ToString() ?? Address.Format(RightRow, Column);
}

public sealed record AlignedRow(int? Left, int? Right);
public sealed record SheetDiff(string Name, Sheet? Left, Sheet? Right, IReadOnlyList<AlignedRow> Rows,
    IReadOnlyList<Difference> Differences, bool UsesKey, string? KeyWarning);

public static class WorkbookComparer
{
    public static IReadOnlyList<SheetDiff> Compare(Workbook left, Workbook right, int? keyColumn = null,
        CancellationToken token = default)
    {
        var names = left.Sheets.Select(s => s.Name).Union(right.Sheets.Select(s => s.Name), StringComparer.OrdinalIgnoreCase);
        var leftSheets = left.Sheets.ToDictionary(s => s.Name, StringComparer.OrdinalIgnoreCase);
        var rightSheets = right.Sheets.ToDictionary(s => s.Name, StringComparer.OrdinalIgnoreCase);
        var leftOrder = left.Sheets.Where(s => rightSheets.ContainsKey(s.Name)).Select((s, i) => (s.Name, Position: i + 1))
            .ToDictionary(x => x.Name, x => x.Position, StringComparer.OrdinalIgnoreCase);
        var rightOrder = right.Sheets.Where(s => leftSheets.ContainsKey(s.Name)).Select((s, i) => (s.Name, Position: i + 1))
            .ToDictionary(x => x.Name, x => x.Position, StringComparer.OrdinalIgnoreCase);

        var result = new List<SheetDiff>();
        var compareDateSystem = left.Date1904 != right.Date1904;
        foreach (var name in names)
        {
            token.ThrowIfCancellationRequested();
            var a = leftSheets.GetValueOrDefault(name);
            var b = rightSheets.GetValueOrDefault(name);
            var changes = new List<Difference>();
            var rows = Align(a, b, keyColumn, out var warning);

            if (a is null || b is null)
                changes.Add(new Difference(a is null ? ChangeKind.Added : ChangeKind.Removed, "シート", a?.Name, b?.Name));
            else
            {
                Add(changes, ChangeKind.Setting, "シート名", a.Name, b.Name);
                Add(changes, ChangeKind.Setting, "シート順", leftOrder[name].ToString(), rightOrder[name].ToString());
                Add(changes, ChangeKind.Setting, "シートの表示状態", a.Visibility, b.Visibility);
                if (compareDateSystem)
                {
                    Add(changes, ChangeKind.Setting, "日付システム", left.Date1904 ? "1904" : "1900", right.Date1904 ? "1904" : "1900");
                    compareDateSystem = false;
                }
                CompareSettings(changes, a.Settings, b.Settings, skipRows: true);
                var leftRowSettings = GroupRowSettings(a);
                var rightRowSettings = GroupRowSettings(b);
                if (keyColumn is not null && warning is null)
                {
                    var matched = rows.Where(r => r.Left.HasValue && r.Right.HasValue).ToArray();
                    if (!matched.OrderBy(r => r.Left).SequenceEqual(matched.OrderBy(r => r.Right)))
                        changes.Add(new Difference(ChangeKind.Setting, "キー列で対応する行の順序", "変更前の順序", "変更後の順序"));
                }

                CompareRows(changes, a, b, rows, leftRowSettings, rightRowSettings,
                    keyColumn.HasValue && warning is null, token);
            }
            result.Add(new SheetDiff(name, a, b, rows, changes, keyColumn is not null && warning is null, warning));
        }
        return result;
    }

    private static void CompareRows(List<Difference> changes, Sheet left, Sheet right, AlignedRow[] rows,
        Dictionary<int, IReadOnlyDictionary<string, string>> leftRowSettings,
        Dictionary<int, IReadOnlyDictionary<string, string>> rightRowSettings, bool usesKey, CancellationToken token)
    {
        var merges = new MergeComparison(left, right, rows, usesKey, token);
        foreach (var row in rows)
        {
            token.ThrowIfCancellationRequested();
            var leftRanges = RangesAt(left, row.Left);
            var rightRanges = RangesAt(right, row.Right);
            var leftCells = row.Left is int l ? left.Rows.GetValueOrDefault(l) : null;
            var rightCells = row.Right is int r ? right.Rows.GetValueOrDefault(r) : null;

            if (row.Left is null || row.Right is null)
            {
                var cells = leftCells ?? rightCells;
                var ranges = row.Left is null ? rightRanges : leftRanges;
                if (cells is not null && cells.Keys.Any(column => FindRange(ranges, column) is null))
                    changes.Add(new Difference(row.Left is null ? ChangeKind.Added : ChangeKind.Removed,
                        "行", row.Left?.ToString(), row.Right?.ToString(), row.Left, row.Right));
                continue;
            }

            CompareSettings(changes, leftRowSettings.GetValueOrDefault(row.Left.Value, Empty),
                rightRowSettings.GetValueOrDefault(row.Right.Value, Empty), row);
            foreach (var column in (leftCells?.Keys ?? []).Union(rightCells?.Keys ?? []).Order())
            {
                var before = leftCells?.GetValueOrDefault(column);
                var after = rightCells?.GetValueOrDefault(column);
                var a = FindRange(leftRanges, column);
                var b = FindRange(rightRanges, column);
                if (a is not null || b is not null)
                {
                    if (merges.Collect(a, b, before, after, row, column, left.DefaultFormat, right.DefaultFormat)) continue;
                    before = a is null ? before : null;
                    after = b is null ? after : null;
                }
                CompareCell(changes, left, right, before, after, row, column);
            }
        }
        merges.AppendTo(changes);
    }

    private static void CompareCell(List<Difference> changes, Sheet left, Sheet right, Cell? before, Cell? after,
        AlignedRow row, int column)
    {
        var oldValue = before?.Value is { Length: > 0 } ? before.Display : null;
        var newValue = after?.Value is { Length: > 0 } ? after.Display : null;
        if (oldValue is not null && newValue is not null && before?.Kind != after?.Kind && oldValue == newValue)
        {
            oldValue = $"{oldValue} ({KindName(before!.Kind)})";
            newValue = $"{newValue} ({KindName(after!.Kind)})";
        }
        Add(changes, ChangeKind.Value, "値", oldValue, newValue, row, column);
        Add(changes, ChangeKind.Formula, "数式", before?.Formula, after?.Formula, row, column);
        CompareSettings(changes, before?.Format ?? left.DefaultFormat,
            after?.Format ?? right.DefaultFormat, row, column);
    }

    private static IReadOnlyList<MergeRange> RangesAt(Sheet sheet, int? row)
    {
        if (row is not int number || sheet.Merges.Ranges.Count == 0) return [];
        var ranges = sheet.Merges.InRows(number, number);
        if (ranges.Count > 1) ranges.Sort((a, b) => a.FirstColumn.CompareTo(b.FirstColumn));
        return ranges;
    }

    private static MergeRange? FindRange(IReadOnlyList<MergeRange> ranges, int column)
    {
        var low = 0;
        var high = ranges.Count;
        while (low < high)
        {
            var middle = (low + high) / 2;
            var range = ranges[middle];
            if (column < range.FirstColumn) high = middle;
            else if (column > range.LastColumn) low = middle + 1;
            else return range;
        }
        return null;
    }

    private sealed class MergeComparison
    {
        private readonly Dictionary<MergeRange, Group> leftGroups = [];
        private readonly Dictionary<MergeRange, Group> rightGroups = [];
        private readonly List<Group> groups = [];

        internal MergeComparison(Sheet left, Sheet right, AlignedRow[] rows, bool usesKey, CancellationToken token)
        {
            if (left.Merges.Ranges.Count == 0 && right.Merges.Ranges.Count == 0) return;
            var rightByAnchor = right.Merges.Ranges.ToDictionary(range => (range.FirstRow, range.FirstColumn));
            var leftToRight = usesKey ? rows.Where(row => row.Left.HasValue)
                .ToDictionary(row => row.Left!.Value, row => row.Right) : null;
            var rightToLeft = usesKey ? rows.Where(row => row.Right.HasValue)
                .ToDictionary(row => row.Right!.Value, row => row.Left) : null;
            foreach (var a in left.Merges.Ranges)
            {
                token.ThrowIfCancellationRequested();
                var r = usesKey ? leftToRight!.GetValueOrDefault(a.FirstRow) : a.FirstRow;
                var b = r is int target ? rightByAnchor.GetValueOrDefault((target, a.FirstColumn)) : null;
                AddGroup(left, right, a, b, new AlignedRow(a.FirstRow, r), a.FirstColumn);
            }
            foreach (var b in right.Merges.Ranges)
            {
                token.ThrowIfCancellationRequested();
                if (rightGroups.ContainsKey(b)) continue;
                var l = usesKey ? rightToLeft!.GetValueOrDefault(b.FirstRow) : b.FirstRow;
                AddGroup(left, right, null, b, new AlignedRow(l, b.FirstRow), b.FirstColumn);
            }
        }

        private void AddGroup(Sheet left, Sheet right, MergeRange? a, MergeRange? b, AlignedRow row, int column)
        {
            var group = new Group(a, b, row, column);
            groups.Add(group);
            if (a is not null) leftGroups.Add(a, group);
            if (b is not null) rightGroups.Add(b, group);
            if (a is null || b is null || a.LastRow - a.FirstRow != b.LastRow - b.FirstRow ||
                a.LastColumn != b.LastColumn)
                group.Parts.Add(new DifferencePart("結合範囲", a?.ToString(), b?.ToString()));

            static Cell? Anchor(Sheet sheet, MergeRange? range, int? row, int column)
            {
                if (row is not int number) return null;
                if (range is null && FindRange(RangesAt(sheet, number), column) is not null) return null;
                return sheet.At(number, column);
            }

            var changes = new List<Difference>();
            CompareCell(changes, left, right, Anchor(left, a, row.Left, column),
                Anchor(right, b, row.Right, column), row, column);
            foreach (var change in changes)
                group.Parts.Add(new DifferencePart(change.Item, change.Left, change.Right));
        }

        internal bool Collect(MergeRange? a, MergeRange? b, Cell? before, Cell? after, AlignedRow row, int column,
            IReadOnlyDictionary<string, string> leftDefault, IReadOnlyDictionary<string, string> rightDefault)
        {
            if ((a is not null && leftGroups[a].Row.Right == row.Right && a.IsAnchor(row.Left!.Value, column)) ||
                (b is not null && rightGroups[b].Row.Left == row.Left && b.IsAnchor(row.Right!.Value, column)))
                return true;

            if (a is not null && b is not null && leftGroups[a] == rightGroups[b])
            {
                var beforeFormat = before?.Format ?? leftDefault;
                var afterFormat = after?.Format ?? rightDefault;
                if (!ReferenceEquals(beforeFormat, afterFormat))
                {
                    foreach (var key in beforeFormat.Keys.Union(afterFormat.Keys).Where(key => key.StartsWith("罫線/", StringComparison.Ordinal)))
                    {
                        var oldValue = beforeFormat.GetValueOrDefault(key);
                        var newValue = afterFormat.GetValueOrDefault(key);
                        if (oldValue != newValue) leftGroups[a].Parts.Add(new DifferencePart(key, oldValue, newValue));
                    }
                }
                return true;
            }

            // A cell uncovered by a merge change may now hold its own value or formula.
            return !(a is null && HasContent(before)) && !(b is null && HasContent(after));
        }

        private static bool HasContent(Cell? cell) => cell is { Value.Length: > 0 } || cell?.Formula is not null;

        internal void AppendTo(List<Difference> changes)
        {
            foreach (var group in groups.OrderBy(group => group.Row.Left ?? group.Row.Right).ThenBy(group => group.Column))
            {
                if (group.Parts.Count == 0) continue;
                var parts = group.Parts.Distinct().ToArray();
                changes.Add(new Difference(ChangeKind.Merge, "結合セル", parts[0].Left, parts[0].Right,
                    group.Row.Left, group.Row.Right, group.Column, group.Left, group.Right, parts));
            }
        }

        private sealed record Group(MergeRange? Left, MergeRange? Right, AlignedRow Row, int Column)
        {
            internal List<DifferencePart> Parts { get; } = [];
        }
    }

    private static AlignedRow[] Align(Sheet? left, Sheet? right, int? column, out string? warning)
    {
        warning = null;
        var l = left?.Rows.Keys ?? [];
        var r = right?.Rows.Keys ?? [];
        if (left is null || right is null)
            return l.Select(i => new AlignedRow(i, null)).Concat(r.Select(i => new AlignedRow(null, i))).ToArray();
        if (column is int key)
        {
            var before = KeyRows(left, key);
            var after = KeyRows(right, key);
            if (before is not null && after is not null)
            {
                // Keep existing rows in left order and insert right-only rows before their next match.
                var rows = new List<AlignedRow>();
                var insertions = new Dictionary<int, List<AlignedRow>>();
                var pending = new List<AlignedRow>();
                foreach (var (value, number) in after)
                {
                    if (before.TryGetValue(value, out var corresponding))
                    {
                        insertions.Add(corresponding, pending);
                        pending = [];
                    }
                    else pending.Add(new AlignedRow(null, number));
                }
                foreach (var (value, number) in before)
                {
                    if (insertions.TryGetValue(number, out var added)) rows.AddRange(added);
                    rows.Add(new AlignedRow(number, after.TryGetValue(value, out var matched) ? matched : null));
                }
                rows.AddRange(pending);
                return rows.ToArray();
            }
            warning = $"キー列 {Address.Column(key)} に空白または重複があるため、セル番地で比較しています。";
        }
        var leftNumbers = l.ToHashSet();
        var rightNumbers = r.ToHashSet();
        return leftNumbers.Union(rightNumbers).Union(GroupRowSettings(left).Keys).Union(GroupRowSettings(right).Keys)
            .Order().Select(i => leftNumbers.Contains(i) == rightNumbers.Contains(i) ? new AlignedRow(i, i) :
                leftNumbers.Contains(i) ? new AlignedRow(i, null) : new AlignedRow(null, i)).ToArray();
    }

    private static Dictionary<(string, string), int>? KeyRows(Sheet sheet, int column)
    {
        var result = new Dictionary<(string, string), int>();
        foreach (var (row, cells) in sheet.Rows)
        {
            var key = cells.GetValueOrDefault(column);
            if (key is null || string.IsNullOrWhiteSpace(key.Value) || !result.TryAdd((key.Kind, key.Value), row))
                return null;
        }
        return result;
    }

    private static readonly IReadOnlyDictionary<string, string> Empty = new Dictionary<string, string>();

    private static Dictionary<int, IReadOnlyDictionary<string, string>> GroupRowSettings(Sheet sheet) =>
        sheet.Settings.Where(item => item.Key.StartsWith("行 ", StringComparison.Ordinal))
            .Select(item => (Parts: item.Key.Split(" / ", 2), item.Value))
            .GroupBy(item => int.Parse(item.Parts[0].AsSpan(2)))
            .ToDictionary(group => group.Key, group => (IReadOnlyDictionary<string, string>)group.ToDictionary(
                item => item.Parts[1] == "ht" ? "行高" : item.Parts[1] == "hidden" ? "行の非表示" : item.Parts[1],
                item => item.Value));

    private static void CompareSettings(List<Difference> changes, IReadOnlyDictionary<string, string> left,
        IReadOnlyDictionary<string, string> right, AlignedRow? row = null, int? column = null, bool skipRows = false)
    {
        if (ReferenceEquals(left, right)) return;
        var kind = column is null ? ChangeKind.Setting : ChangeKind.Format;
        foreach (var key in left.Keys.Union(right.Keys).Order(StringComparer.Ordinal))
            if (!skipRows || !key.StartsWith("行 ", StringComparison.Ordinal))
                Add(changes, kind, FormatLabel(key), left.GetValueOrDefault(key), right.GetValueOrDefault(key), row, column);
    }

    private static void Add(List<Difference> changes, ChangeKind kind, string item, string? before, string? after,
        AlignedRow? row = null, int? column = null)
    {
        if (before != after) changes.Add(new Difference(kind, item, before, after, row?.Left, row?.Right, column));
    }

    private static string KindName(string kind) => kind switch
    {
        "n" => "数値", "text" => "文字列", "b" => "論理値", "e" => "エラー", "d" => "日付", _ => kind
    };

    private static string FormatLabel(string key) => key switch
    {
        "フォント/sz/val" => "フォントサイズ", "フォント/name/val" => "フォント名",
        "配置/wrapText" => "折り返して全体を表示", "配置/shrinkToFit" => "縮小して全体を表示",
        "印刷 / scale" => "印刷倍率", "印刷 / orientation" => "印刷の向き",
        "印刷 / paperSize" => "用紙", _ => key
    };
}
