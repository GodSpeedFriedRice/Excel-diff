namespace ExcelDiff.Desktop.Views;

internal sealed class GridSort
{
    internal GridSort(DataGridView grid, Action refresh)
    {
        foreach (DataGridViewColumn column in grid.Columns)
            column.SortMode = DataGridViewColumnSortMode.Programmatic;
        grid.ColumnHeaderMouseClick += (_, e) =>
        {
            if (e.Button != MouseButtons.Left || e.ColumnIndex < 0) return;
            if (Column >= 0) grid.Columns[Column].HeaderCell.SortGlyphDirection = SortOrder.None;
            Direction = Column == e.ColumnIndex && Direction == SortOrder.Ascending
                ? SortOrder.Descending : SortOrder.Ascending;
            Column = e.ColumnIndex;
            grid.Columns[Column].HeaderCell.SortGlyphDirection = Direction;
            refresh();
        };
    }

    internal int Column { get; private set; } = -1;
    internal SortOrder Direction { get; private set; }

    internal IOrderedEnumerable<T> Order<T, TKey>(IEnumerable<T> rows, Func<T, TKey> key, IComparer<TKey>? comparer = null) =>
        Direction == SortOrder.Descending ? rows.OrderByDescending(key, comparer) : rows.OrderBy(key, comparer);
}
