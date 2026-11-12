using System.Globalization;
using ExcelDiff.Core.Workbooks;

namespace ExcelDiff.Core.Comparison;

public enum ChangeKind { Value, Formula, Format, Setting, Added, Removed, Merge, RowLayout, ColumnLayout }

public sealed record DifferencePart(string Item, string? Left, string? Right);

public sealed record Difference(ChangeKind Kind, string Item, string? Left, string? Right,
    int? LeftRow = null, int? RightRow = null, int? Column = null,
    MergeRange? LeftMerge = null, MergeRange? RightMerge = null, IReadOnlyList<DifferencePart>? Parts = null, int LayoutCount = 1)
{
    public string LeftAddress => LeftMerge?.ToString() ?? DescribeAddress(LeftRow);
    public string RightAddress => RightMerge?.ToString() ?? DescribeAddress(RightRow);

    private string DescribeAddress(int? row)
    {
        if (Kind == ChangeKind.ColumnLayout && Column is int column)
            return LayoutCount == 1 ? $"{Address.Column(column)} 列" : $"{Address.Column(column)}:{Address.Column(column + LayoutCount - 1)} 列";
        if (Kind == ChangeKind.RowLayout && row is int number && LayoutCount > 1)
            return $"{number}–{number + LayoutCount - 1} 行";
        return Address.Format(row, Column);
    }
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
                CompareSettings(changes, a.Settings, b.Settings);
                AddDefaultSizeChange(changes, a, b, rows: true);
                AddDefaultSizeChange(changes, a, b, rows: false);
                var leftRowSettings = GroupLayoutSettings(a, rows: true);
                var rightRowSettings = GroupLayoutSettings(b, rows: true);
                AddColumnChanges(changes, a, b);
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
        var layout = new List<Difference>();
        var leftDefaultHeight = DefaultSize(left, rows: true);
        var rightDefaultHeight = DefaultSize(right, rows: true);
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

            if (leftRowSettings.ContainsKey(row.Left.Value) || rightRowSettings.ContainsKey(row.Right.Value))
            {
                var parts = CompareLayout(leftRowSettings.GetValueOrDefault(row.Left.Value, Empty),
                    rightRowSettings.GetValueOrDefault(row.Right.Value, Empty), "ht", leftDefaultHeight, rightDefaultHeight);
                AppendLayout(layout, ChangeKind.RowLayout, row.Left, row.Right, null, parts);
            }
            foreach (var column in (leftCells?.Keys ?? []).Union(rightCells?.Keys ?? []).Order())
            {
                var before = leftCells?.GetValueOrDefault(column);
                var after = rightCells?.GetValueOrDefault(column);
                var a = FindRange(leftRanges, column);
                var b = FindRange(rightRanges, column);
                if (a is not null || b is not null)
                {
                    if (merges.Collect(a, b, before, after, row, column,
                        left.InheritedFormat(row.Left.Value, column), right.InheritedFormat(row.Right.Value, column))) continue;
                    before = a is null ? before : null;
                    after = b is null ? after : null;
                }
                CompareCell(changes, left, right, before, after, row, column);
            }
        }
        if (usesKey)
        {
            // A sized blank row has no key value and is absent from the aligned data rows.
            foreach (var number in leftRowSettings.Keys.Union(rightRowSettings.Keys).Order())
            {
                if (left.Rows.ContainsKey(number) || right.Rows.ContainsKey(number)) continue;
                var parts = CompareLayout(leftRowSettings.GetValueOrDefault(number, Empty),
                    rightRowSettings.GetValueOrDefault(number, Empty), "ht", leftDefaultHeight, rightDefaultHeight);
                AppendLayout(layout, ChangeKind.RowLayout, number, number, null, parts);
            }
        }
        changes.AddRange(layout);
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
        var leftBase = row.Left is int l ? left.InheritedFormat(l, column) : left.DefaultFormat;
        var rightBase = row.Right is int r ? right.InheritedFormat(r, column) : right.DefaultFormat;
        var leftFormat = before?.Format ?? leftBase;
        var rightFormat = after?.Format ?? rightBase;
        if (ReferenceEquals(leftFormat, rightFormat) ||
            (ReferenceEquals(leftFormat, leftBase) && ReferenceEquals(rightFormat, rightBase))) return;

        foreach (var key in leftFormat.Keys.Union(rightFormat.Keys).Order(StringComparer.Ordinal))
        {
            var oldValueForKey = leftFormat.GetValueOrDefault(key);
            var newValueForKey = rightFormat.GetValueOrDefault(key);
            if (oldValueForKey == leftBase.GetValueOrDefault(key) && newValueForKey == rightBase.GetValueOrDefault(key)) continue;
            Add(changes, ChangeKind.Format, FormatLabel(key), oldValueForKey, newValueForKey, row, column);
        }
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
                        if (oldValue != newValue &&
                            (oldValue != leftDefault.GetValueOrDefault(key) || newValue != rightDefault.GetValueOrDefault(key)))
                            leftGroups[a].Parts.Add(new DifferencePart(key, oldValue, newValue));
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
        return leftNumbers.Union(rightNumbers).Union(GroupLayoutSettings(left, rows: true).Keys).Union(GroupLayoutSettings(right, rows: true).Keys)
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

    private static Dictionary<int, IReadOnlyDictionary<string, string>> GroupLayoutSettings(Sheet sheet, bool rows)
    {
        var prefix = rows ? "行 " : "列 ";
        var settings = new Dictionary<int, Dictionary<string, string>>();
        foreach (var (name, value) in sheet.Settings)
        {
            if (!name.StartsWith(prefix, StringComparison.Ordinal)) continue;
            var separator = name.IndexOf(" / ", StringComparison.Ordinal);
            if (separator < 0) continue;
            var label = name.AsSpan(prefix.Length, separator - prefix.Length);
            if (!(rows ? int.TryParse(label, out var index) : Address.TryColumn(label, out index))) continue;
            if (!settings.TryGetValue(index, out var items)) settings[index] = items = new Dictionary<string, string>();
            items[name[(separator + 3)..]] = value;
        }
        return settings.ToDictionary(entry => entry.Key, entry => (IReadOnlyDictionary<string, string>)entry.Value);
    }

    private static string DefaultSize(Sheet sheet, bool rows)
    {
        var name = rows ? "defaultRowHeight" : "defaultColWidth";
        return NormalizeSize(sheet.Settings.GetValueOrDefault("既定レイアウト / " + name) ?? (rows ? "15" : "8.43"));
    }

    private static string NormalizeSize(string size) => decimal.TryParse(size, NumberStyles.Float,
        CultureInfo.InvariantCulture, out var value) ? value.ToString("G29", CultureInfo.InvariantCulture) : size;

    private static void AddDefaultSizeChange(List<Difference> changes, Sheet left, Sheet right, bool rows) =>
        Add(changes, ChangeKind.Setting, rows ? "既定の行高" : "既定の列幅", DefaultSize(left, rows), DefaultSize(right, rows));

    private static DifferencePart[] CompareLayout(IReadOnlyDictionary<string, string> left,
        IReadOnlyDictionary<string, string> right, string sizeKey, string leftDefault, string rightDefault)
    {
        var parts = new List<DifferencePart>();
        foreach (var key in left.Keys.Union(right.Keys).Order(StringComparer.Ordinal))
        {
            var before = left.GetValueOrDefault(key);
            var after = right.GetValueOrDefault(key);
            var label = FormatLabel(key);
            if (key == sizeKey)
            {
                before = NormalizeSize(before ?? leftDefault);
                after = NormalizeSize(after ?? rightDefault);
                if (before == leftDefault && after == rightDefault) continue;
                label = key == "ht" ? "行高" : "列幅";
            }
            else if (key == "hidden")
            {
                before = before is "1" or "true" ? "非表示" : "表示";
                after = after is "1" or "true" ? "非表示" : "表示";
                label = "表示状態";
            }
            if (before != after) parts.Add(new DifferencePart(label, before, after));
        }
        return parts.ToArray();
    }

    private static void AddColumnChanges(List<Difference> changes, Sheet left, Sheet right)
    {
        var a = GroupLayoutSettings(left, rows: false);
        var b = GroupLayoutSettings(right, rows: false);
        var before = DefaultSize(left, rows: false);
        var after = DefaultSize(right, rows: false);
        foreach (var column in a.Keys.Union(b.Keys).Order())
        {
            var parts = CompareLayout(a.GetValueOrDefault(column, Empty), b.GetValueOrDefault(column, Empty), "width", before, after);
            AppendLayout(changes, ChangeKind.ColumnLayout, null, null, column, parts);
        }
    }

    private static void AppendLayout(List<Difference> changes, ChangeKind kind, int? leftRow, int? rightRow,
        int? column, DifferencePart[] parts)
    {
        if (parts.Length == 0) return;
        var previous = changes.Count == 0 ? null : changes[^1];
        if (previous?.Kind == kind && previous.LeftRow + previous.LayoutCount == leftRow &&
            previous.RightRow + previous.LayoutCount == rightRow && previous.Column + previous.LayoutCount == column &&
            previous.Parts!.SequenceEqual(parts))
        {
            changes[^1] = previous with { LayoutCount = previous.LayoutCount + 1 };
            return;
        }
        changes.Add(new Difference(kind, kind == ChangeKind.RowLayout ? "行設定" : "列設定", parts[0].Left, parts[0].Right,
            leftRow, rightRow, column, Parts: parts));
    }

    private static void CompareSettings(List<Difference> changes, IReadOnlyDictionary<string, string> left,
        IReadOnlyDictionary<string, string> right)
    {
        if (ReferenceEquals(left, right)) return;
        foreach (var key in left.Keys.Union(right.Keys).Order(StringComparer.Ordinal))
            if (!key.StartsWith("行 ", StringComparison.Ordinal) && !key.StartsWith("列 ", StringComparison.Ordinal) &&
                key is not ("既定レイアウト / defaultRowHeight" or "既定レイアウト / defaultColWidth"))
                Add(changes, ChangeKind.Setting, FormatLabel(key), left.GetValueOrDefault(key), right.GetValueOrDefault(key));
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
