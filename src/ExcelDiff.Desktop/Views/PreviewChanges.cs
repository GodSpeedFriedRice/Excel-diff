using ExcelDiff.Core.Comparison;
using ExcelDiff.Core.Workbooks;

namespace ExcelDiff.Desktop.Views;

internal sealed class PreviewChanges
{
    private readonly Dictionary<int, Dictionary<int, Mark>> cells = [];
    private readonly Spans merges;
    private readonly Spans rows;
    private readonly Spans columns;

    internal PreviewChanges(IReadOnlyList<Difference> differences, bool before)
    {
        var merged = new Dictionary<MergeRange, Mark>();
        var changedRows = new Dictionary<MergeRange, Mark>();
        var changedColumns = new Dictionary<MergeRange, Mark>();
        for (var index = 0; index < differences.Count; index++)
        {
            var diff = differences[index];
            var row = before ? diff.LeftRow : diff.RightRow;
            var range = before ? diff.LeftMerge : diff.RightMerge;
            var column = diff.Column;
            var color = diff.Kind switch
            {
                ChangeKind.Added => Appearance.Added,
                ChangeKind.Removed => Appearance.Removed,
                _ => Appearance.Changed
            };
            if (range is not null)
                Add(merged, range, index, color);
            else if (diff.Kind == ChangeKind.ColumnLayout && column is int c)
                Add(changedColumns, new MergeRange(c, 1, c + diff.LayoutCount - 1, 1), index, color);
            else if (row is int r && column is int number)
            {
                if (!cells.TryGetValue(r, out var rowMarks)) cells[r] = rowMarks = [];
                Add(rowMarks, number, index, color);
            }
            else if (row is int first)
                Add(changedRows, new MergeRange(first, 1, first + diff.LayoutCount - 1, 1),
                    index, color, diff.Kind is ChangeKind.Added or ChangeKind.Removed);
        }
        merges = new Spans(merged);
        rows = new Spans(changedRows);
        columns = new Spans(changedColumns);
    }

    internal IReadOnlyDictionary<int, Mark>? CellsInRow(int row) => cells.GetValueOrDefault(row);
    internal IEnumerable<(MergeRange Range, Mark Mark)> Merges(int first, int last) => merges.Intersecting(first, last);
    internal IEnumerable<(MergeRange Range, Mark Mark)> Rows(int first, int last) => rows.Intersecting(first, last);
    internal IEnumerable<(MergeRange Range, Mark Mark)> Columns(int first, int last) => columns.Intersecting(first, last);

    internal Mark? Pick(int? row, int? column)
    {
        if (row is int r && column is int c)
            return CellsInRow(r)?.GetValueOrDefault(c) ?? merges.At(r, c) ?? rows.At(r, 1, bodyOnly: true);
        if (row is int rowHeader) return rows.At(rowHeader, 1);
        return column is int columnHeader ? columns.At(columnHeader, 1) : null;
    }

    private static void Add<TKey>(Dictionary<TKey, Mark> target, TKey key, int index, Color color, bool fillRow = false)
        where TKey : notnull
    {
        if (!target.TryGetValue(key, out var mark)) target[key] = mark = new Mark(color, fillRow);
        mark.Indices.Add(index);
    }

    internal sealed record Mark(Color Color, bool FillRow)
    {
        internal List<int> Indices { get; } = [];
    }

    private sealed class Spans(Dictionary<MergeRange, Mark> marks)
    {
        private readonly MergeIndex index = new(marks.Keys);

        internal IEnumerable<(MergeRange Range, Mark Mark)> Intersecting(int first, int last)
        {
            foreach (var range in index.InRows(first, last)) yield return (range, marks[range]);
        }

        internal Mark? At(int row, int column, bool bodyOnly = false)
        {
            foreach (var range in index.InRows(row, row))
                if (range.ContainsColumn(column) && (!bodyOnly || marks[range].FillRow)) return marks[range];
            return null;
        }
    }
}
