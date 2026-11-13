using ExcelDiff.Core.Comparison;
using ExcelDiff.Core.Workbooks;

namespace ExcelDiff.Desktop.Views;

internal sealed class XmlDiffView : UserControl
{
    private const int ColumnsPerSide = 3;
    private const int CodeColumn = 2;
    private static readonly Color Added = Color.FromArgb(230, 255, 236);
    private static readonly Color Removed = Color.FromArgb(255, 235, 233);
    private static readonly Color Empty = Color.FromArgb(246, 248, 250);
    private static readonly Color Hunk = Color.FromArgb(221, 244, 255);
    private readonly DataGridView grid = Appearance.Grid();
    private readonly CheckBox showAll = new() { Text = "変更のない行も表示", AutoSize = true };
    private readonly Label summary = Appearance.Label("");
    private readonly Font codeFont = new(FontFamily.GenericMonospace, 9);
    private CancellationTokenSource? pending;
    private Sheet? before;
    private Sheet? after;
    private XmlLineDiff? diff;
    private readonly List<SplitRow> visible = [];
    private readonly List<int> changeStarts = [];

    internal XmlDiffView()
    {
        Dock = DockStyle.Fill;
        var layout = Appearance.Stack();
        var toolbar = Appearance.Toolbar();
        showAll.Margin = new Padding(8, 8, 6, 0);
        toolbar.Controls.AddRange([Appearance.Button("↑ 前の差分", () => MoveChange(-1)),
            Appearance.Button("↓ 次の差分", () => MoveChange(1)), showAll, summary]);
        Appearance.Row(layout, toolbar);

        grid.VirtualMode = true;
        grid.AccessibleName = "シートのXML差分";
        grid.DefaultCellStyle.Font = codeFont;
        grid.DefaultCellStyle.WrapMode = DataGridViewTriState.False;
        grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
        grid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
        grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
        grid.RowTemplate.Height = 24;
        AddSide("変更前のXML");
        AddSide("変更後のXML");
        grid.CellValueNeeded += FillLine;
        grid.CellFormatting += ColorLine;
        grid.CellPainting += DrawDivider;
        showAll.CheckedChanged += (_, _) => Render();
        Appearance.Row(layout, grid, fill: true);
        Controls.Add(layout);
    }

    private void AddSide(string title)
    {
        grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            Width = 54, SortMode = DataGridViewColumnSortMode.NotSortable,
            DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleRight }
        });
        grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            Width = 25, SortMode = DataGridViewColumnSortMode.NotSortable,
            DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter }
        });
        grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = title, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            FillWeight = 50, SortMode = DataGridViewColumnSortMode.NotSortable
        });
    }

    internal async void ShowSheets(Sheet? left, Sheet? right)
    {
        if (ReferenceEquals(before, left) && ReferenceEquals(after, right) && (diff is not null || pending is not null)) return;
        before = left;
        after = right;
        pending?.Cancel();
        using var request = new CancellationTokenSource();
        pending = request;
        diff = null;
        Render();
        summary.Text = "XMLを比較しています…";
        try
        {
            var result = await Task.Run(() => XmlLineDiff.Compare(left?.Xml, right?.Xml, request.Token), request.Token);
            if (IsDisposed || !ReferenceEquals(pending, request)) return;
            diff = result;
            Render();
        }
        catch (OperationCanceledException) when (request.IsCancellationRequested) { }
        catch (ArgumentException)
        {
            if (!IsDisposed && ReferenceEquals(pending, request)) summary.Text = "XML差分を表示できません。";
        }
        finally
        {
            if (ReferenceEquals(pending, request)) pending = null;
        }
    }

    private void Render()
    {
        grid.RowCount = 0;
        visible.Clear();
        changeStarts.Clear();
        if (diff is null) return;
        var lines = diff.Lines;
        var additions = 0;
        var removals = 0;
        var blocks = new List<(int First, int End)>();
        for (var i = 0; i < lines.Count; i++)
        {
            if (lines[i].Kind == XmlLineKind.Context) continue;
            if (lines[i].Kind == XmlLineKind.Added) additions++;
            else removals++;
            var first = Math.Max(0, i - 3);
            var end = Math.Min(lines.Count, i + 4);
            if (blocks.Count == 0 || blocks[^1].End < first) blocks.Add((first, end));
            else blocks[^1] = (blocks[^1].First, end);
        }

        summary.Text = additions == 0 && removals == 0 ? "XMLに差分はありません。" :
            $"+{additions} 行　−{removals} 行" + (diff.HasUnalignedBlocks ? "　大きな変更は範囲単位で表示" : "");
        if (showAll.Checked)
            AppendRows(lines, 0, lines.Count);
        else
        {
            foreach (var (first, end) in blocks)
            {
                int? oldLine = null;
                int? newLine = null;
                var oldCount = 0;
                var newCount = 0;
                for (var i = first; i < end; i++)
                {
                    oldLine ??= lines[i].Before;
                    newLine ??= lines[i].After;
                    if (lines[i].Before.HasValue) oldCount++;
                    if (lines[i].After.HasValue) newCount++;
                }
                visible.Add(new SplitRow(null, null, $"@@ -{oldLine ?? 0},{oldCount} +{newLine ?? 0},{newCount} @@"));
                AppendRows(lines, first, end);
            }
        }
        grid.RowCount = visible.Count;
        grid.ClearSelection();
    }

    private void AppendRows(IReadOnlyList<XmlDiffLine> lines, int first, int end)
    {
        var i = first;
        while (i < end)
        {
            if (lines[i].Kind == XmlLineKind.Context)
            {
                visible.Add(new SplitRow(lines[i], lines[i]));
                i++;
                continue;
            }

            // Pair removed and added lines in the same change. Extra lines leave the opposite side empty.
            changeStarts.Add(visible.Count);
            var removedStart = i;
            while (i < end && lines[i].Kind == XmlLineKind.Removed) i++;
            var addedStart = i;
            while (i < end && lines[i].Kind == XmlLineKind.Added) i++;
            var removedCount = addedStart - removedStart;
            var addedCount = i - addedStart;
            for (var offset = 0; offset < Math.Max(removedCount, addedCount); offset++)
                visible.Add(new SplitRow(
                    offset < removedCount ? lines[removedStart + offset] : null,
                    offset < addedCount ? lines[addedStart + offset] : null));
        }
    }

    private void FillLine(object? sender, DataGridViewCellValueEventArgs e)
    {
        if (e.RowIndex < 0 || e.RowIndex >= visible.Count) return;
        var row = visible[e.RowIndex];
        if (row.Heading is string heading)
        {
            if (e.ColumnIndex == CodeColumn) e.Value = heading;
            return;
        }
        var isBefore = e.ColumnIndex < ColumnsPerSide;
        var line = isBefore ? row.Before : row.After;
        if (line is null) return;
        e.Value = (e.ColumnIndex % ColumnsPerSide) switch
        {
            0 => isBefore ? line.Before : line.After,
            1 => line.Kind == XmlLineKind.Added ? "+" : line.Kind == XmlLineKind.Removed ? "−" : "",
            _ => line.Text
        };
    }

    private void ColorLine(object? sender, DataGridViewCellFormattingEventArgs e)
    {
        if (e.RowIndex < 0 || e.RowIndex >= visible.Count || e.CellStyle is not { } style) return;
        var row = visible[e.RowIndex];
        var line = e.ColumnIndex < ColumnsPerSide ? row.Before : row.After;
        style.BackColor = row.Heading is not null ? Hunk : line?.Kind switch
        {
            XmlLineKind.Added => Added,
            XmlLineKind.Removed => Removed,
            XmlLineKind.Context => Color.White,
            _ => Empty
        };
    }

    private void DrawDivider(object? sender, DataGridViewCellPaintingEventArgs e)
    {
        if (e.ColumnIndex != CodeColumn || e.Graphics is not { } graphics) return;
        e.Paint(e.ClipBounds, e.PaintParts);
        using var pen = new Pen(Appearance.Line);
        var edge = e.CellBounds.Right - 1;
        graphics.DrawLine(pen, edge, e.CellBounds.Top, edge, e.CellBounds.Bottom);
        e.Handled = true;
    }

    private void MoveChange(int direction)
    {
        if (changeStarts.Count == 0) return;
        var current = grid.CurrentCell?.RowIndex ?? -1;
        var found = changeStarts.BinarySearch(current);
        var target = found >= 0 ? found + direction : direction > 0 ? ~found : ~found - 1;
        var row = changeStarts[Math.Clamp(target, 0, changeStarts.Count - 1)];
        var column = visible[row].Before is null ? ColumnsPerSide + CodeColumn : CodeColumn;
        grid.CurrentCell = grid[column, row];
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing) pending?.Cancel();
        base.Dispose(disposing);
        if (disposing) codeFont.Dispose();
    }

    private readonly record struct SplitRow(XmlDiffLine? Before, XmlDiffLine? After, string? Heading = null);
}
