using ExcelDiff.Core.Comparison;
using ExcelDiff.Core.Workbooks;

namespace ExcelDiff.Desktop.Views;

internal sealed class XmlDiffView : UserControl
{
    private static readonly Color Added = Color.FromArgb(230, 255, 236);
    private static readonly Color Removed = Color.FromArgb(255, 235, 233);
    private readonly DataGridView grid = Appearance.Grid();
    private readonly CheckBox showAll = new() { Text = "変更のない行も表示", AutoSize = true };
    private readonly Label summary = Appearance.Label("");
    private readonly Font codeFont = new(FontFamily.GenericMonospace, 9);
    private CancellationTokenSource? pending;
    private Sheet? before;
    private Sheet? after;
    private XmlLineDiff? diff;
    private List<(int Index, string? Heading)> visible = [];
    private int[] changeStarts = [];

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
        grid.ShowCellToolTips = false;
        grid.AccessibleName = "シートのXML差分";
        grid.DefaultCellStyle.Font = codeFont;
        grid.DefaultCellStyle.WrapMode = DataGridViewTriState.False;
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        grid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
        foreach (var (title, width) in new[] { ("変更前", 68), ("変更後", 68), ("", 28), ("XML", 900) })
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = title, Width = width, SortMode = DataGridViewColumnSortMode.NotSortable
            });
        grid.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        grid.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        grid.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        grid.CellValueNeeded += FillLine;
        grid.CellFormatting += ColorLine;
        showAll.CheckedChanged += (_, _) => Render();
        Appearance.Row(layout, grid, fill: true);
        Controls.Add(layout);
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
        visible = [];
        changeStarts = [];
        if (diff is null) return;
        var lines = diff.Lines;
        var additions = 0;
        var removals = 0;
        var blocks = new List<(int First, int End)>();
        var starts = new List<int>();
        for (var i = 0; i < lines.Count; i++)
        {
            if (lines[i].Kind == XmlLineKind.Context) continue;
            if (lines[i].Kind == XmlLineKind.Added) additions++;
            else removals++;
            if (i == 0 || lines[i - 1].Kind == XmlLineKind.Context) starts.Add(i);
            var first = Math.Max(0, i - 3);
            var end = Math.Min(lines.Count, i + 4);
            if (blocks.Count == 0 || blocks[^1].End < first) blocks.Add((first, end));
            else blocks[^1] = (blocks[^1].First, end);
        }

        summary.Text = additions == 0 && removals == 0 ? "XMLに差分はありません。" :
            $"+{additions} 行　−{removals} 行" + (diff.HasUnalignedBlocks ? "　大きな変更は範囲単位で表示" : "");
        if (showAll.Checked)
        {
            for (var i = 0; i < lines.Count; i++) visible.Add((i, null));
            changeStarts = starts.ToArray();
        }
        else
        {
            var next = 0;
            foreach (var (first, end) in blocks)
            {
                int? before = null;
                int? after = null;
                var beforeCount = 0;
                var afterCount = 0;
                for (var i = first; i < end; i++)
                {
                    before ??= lines[i].Before;
                    after ??= lines[i].After;
                    if (lines[i].Before.HasValue) beforeCount++;
                    if (lines[i].After.HasValue) afterCount++;
                }
                visible.Add((-1, $"@@ -{before ?? 0},{beforeCount} +{after ?? 0},{afterCount} @@"));
                for (var i = first; i < end; i++)
                {
                    if (next < starts.Count && starts[next] == i)
                    {
                        starts[next++] = visible.Count;
                    }
                    visible.Add((i, null));
                }
            }
            changeStarts = starts.ToArray();
        }
        grid.Columns[3].Width = Math.Clamp(visible.Where(row => row.Index >= 0)
            .Select(row => lines[row.Index].Text.Length).DefaultIfEmpty(0).Max() * 8 + 16, 700, 12000);
        grid.RowCount = visible.Count;
    }

    private void FillLine(object? sender, DataGridViewCellValueEventArgs e)
    {
        if (diff is null || e.RowIndex < 0 || e.RowIndex >= visible.Count) return;
        var row = visible[e.RowIndex];
        if (row.Heading is string heading)
        {
            if (e.ColumnIndex == 3) e.Value = heading;
            return;
        }
        var line = diff.Lines[row.Index];
        e.Value = e.ColumnIndex switch
        {
            0 => line.Before,
            1 => line.After,
            2 => line.Kind == XmlLineKind.Added ? "+" : line.Kind == XmlLineKind.Removed ? "−" : "",
            _ => line.Text
        };
    }

    private void ColorLine(object? sender, DataGridViewCellFormattingEventArgs e)
    {
        if (diff is null || e.RowIndex < 0 || e.RowIndex >= visible.Count || e.CellStyle is not { } style) return;
        var row = visible[e.RowIndex];
        style.BackColor = row.Heading is not null ? Color.FromArgb(221, 244, 255) :
            diff.Lines[row.Index].Kind switch
            {
                XmlLineKind.Added => Added,
                XmlLineKind.Removed => Removed,
                _ => Color.White
            };
    }

    private void MoveChange(int direction)
    {
        if (changeStarts.Length == 0) return;
        var current = grid.CurrentCell?.RowIndex ?? -1;
        var found = Array.BinarySearch(changeStarts, current);
        var target = found >= 0 ? found + direction : direction > 0 ? ~found : ~found - 1;
        grid.CurrentCell = grid[3, changeStarts[Math.Clamp(target, 0, changeStarts.Length - 1)]];
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing) pending?.Cancel();
        base.Dispose(disposing);
        if (disposing) codeFont.Dispose();
    }
}
