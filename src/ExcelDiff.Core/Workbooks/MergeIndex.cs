namespace ExcelDiff.Core.Workbooks;

public sealed record MergeRange(int FirstRow, int FirstColumn, int LastRow, int LastColumn)
{
    public bool ContainsColumn(int column) => column >= FirstColumn && column <= LastColumn;
    public bool IsAnchor(int row, int column) => row == FirstRow && column == FirstColumn;

    public override string ToString() => $"{Address.Format(FirstRow, FirstColumn)}:{Address.Format(LastRow, LastColumn)}";

    public static MergeRange Parse(string reference)
    {
        var ends = reference.Split(':', 2);
        var first = Address.Parse(ends[0]);
        var last = Address.Parse(ends[^1]);
        if (first.Row > last.Row || first.Column > last.Column)
            throw new InvalidDataException($"無効な結合範囲: {reference}");
        return new MergeRange(first.Row, first.Column, last.Row, last.Column);
    }
}

public sealed class MergeIndex
{
    private readonly MergeRange[] ranges;
    private readonly int[] maxEndRow;

    public MergeIndex(IEnumerable<MergeRange> ranges)
    {
        this.ranges = ranges.OrderBy(range => range.FirstRow).ThenBy(range => range.FirstColumn).ToArray();
        maxEndRow = new int[this.ranges.Length];
        Build(0, this.ranges.Length);
    }

    public IReadOnlyList<MergeRange> Ranges => ranges;

    public List<MergeRange> InRows(int firstRow, int lastRow)
    {
        var result = new List<MergeRange>();
        Query(0, ranges.Length, firstRow, lastRow, result);
        return result;
    }

    private int Build(int start, int end)
    {
        if (start >= end) return 0;
        var middle = (start + end) / 2;
        return maxEndRow[middle] = Math.Max(ranges[middle].LastRow,
            Math.Max(Build(start, middle), Build(middle + 1, end)));
    }

    private void Query(int start, int end, int firstRow, int lastRow, List<MergeRange> result)
    {
        if (start >= end || ranges[start].FirstRow > lastRow) return;
        var middle = (start + end) / 2;
        if (maxEndRow[middle] < firstRow) return;
        var range = ranges[middle];
        Query(start, middle, firstRow, lastRow, result);
        if (range.FirstRow <= lastRow && range.LastRow >= firstRow) result.Add(range);
        Query(middle + 1, end, firstRow, lastRow, result);
    }
}
