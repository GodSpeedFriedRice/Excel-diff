using ExcelDiff.Core.Comparison;

namespace ExcelDiff.Desktop.Views;

internal sealed class ThreeWayXmlView : UserControl
{
    private readonly DataGridView grid = Appearance.Grid();
    private readonly Label status = Appearance.Label("");
    private readonly CheckBox context = new() { Text = "変更のない行も表示", AutoSize = true };
    private readonly HScrollBar horizontal = new() { Dock = DockStyle.Top, SmallChange = 24 };
    private readonly Font font = new("Consolas", 9);
    private CancellationTokenSource? pending;
    private IReadOnlyList<ThreeWayXmlLine> lines = [];
    private int[] visible = [];

    internal ThreeWayXmlView()
    {
        Dock = DockStyle.Fill;
        var layout = Appearance.Stack();
        var tools = Appearance.Toolbar();
        tools.Controls.AddRange([status, context]);
        Appearance.Row(layout, tools);
        grid.VirtualMode = true;
        grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
        grid.DefaultCellStyle.Font = font;
        grid.DefaultCellStyle.WrapMode = DataGridViewTriState.False;
        grid.ScrollBars = ScrollBars.Vertical;
        grid.AccessibleName = "3方向のXML差分";
        foreach (var title in new[] { "共通ベース", "本番版", "取り込み元" })
        {
            grid.Columns.Add(new DataGridViewTextBoxColumn { Width = 46, SortMode = DataGridViewColumnSortMode.NotSortable });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = title,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, SortMode = DataGridViewColumnSortMode.NotSortable });
        }
        grid.CellValueNeeded += (_, e) =>
        {
            if (LineAt(e.RowIndex, e.ColumnIndex) is not { } line) return;
            e.Value = e.ColumnIndex % 2 != 0 ? line.Text : (e.ColumnIndex < 2 ? line.Before : line.After)?.ToString();
        };
        grid.CellToolTipTextNeeded += (_, e) =>
        {
            if (e.ColumnIndex >= 0 && e.ColumnIndex % 2 != 0) e.ToolTipText = LineAt(e.RowIndex, e.ColumnIndex)?.Text;
        };
        grid.CellFormatting += (_, e) =>
        {
            if (e.CellStyle is not { } style) return;
            style.BackColor = LineAt(e.RowIndex, e.ColumnIndex)?.Kind switch
            {
                XmlLineKind.Removed => Appearance.Removed,
                XmlLineKind.Added => Appearance.Added,
                XmlLineKind.Context => Color.White,
                _ => Appearance.Panel
            };
        };
        grid.CellPainting += PaintCode;
        horizontal.ValueChanged += (_, _) => grid.Invalidate();
        context.CheckedChanged += (_, _) => Render();
        Appearance.Row(layout, grid, fill: true);
        Appearance.Row(layout, horizontal);
        Controls.Add(layout);
    }

    internal async void ShowPart(string[] paths, PackagePartMergeDifference part)
    {
        Clear();
        status.Text = "XMLを比較しています…";
        using var request = new CancellationTokenSource();
        pending = request;
        try
        {
            var result = await Task.Run(() =>
            {
                var token = request.Token;
                var original = PackageComparer.ReadXml(paths[0], part.Base, token);
                var production = PackageComparer.ReadXml(paths[1], part.Production, token);
                var incoming = PackageComparer.ReadXml(paths[2], part.Incoming, token);
                var ours = XmlLineDiff.Compare(original, production, token);
                var theirs = XmlLineDiff.Compare(original, incoming, token);
                return (Lines: ThreeWayXmlDiff.Align(ours, theirs, token), Unaligned: ours.HasUnalignedBlocks || theirs.HasUnalignedBlocks);
            }, request.Token);
            if (IsDisposed || pending != request) return;
            lines = result.Lines;
            var changes = lines.Count(line => line.Changed);
            status.Text = (changes == 0 ? "整形後のXMLは同じです。" : $"変更のある行 {changes:N0}") +
                (result.Unaligned ? "　大きな変更は範囲単位で表示" : "");
            Render();
        }
        catch (OperationCanceledException) when (request.IsCancellationRequested) { }
        catch (Exception error) when (FolderComparer.IsReadError(error))
        {
            if (!IsDisposed && pending == request) status.Text = $"XML差分を表示できません。{error.Message}";
        }
        finally { if (pending == request) pending = null; }
    }

    internal void Clear()
    {
        pending?.Cancel();
        pending = null;
        lines = [];
        status.Text = "";
        horizontal.Value = 0;
        Render();
    }

    private XmlDiffLine? LineAt(int row, int column)
    {
        if (row < 0 || row >= visible.Length || column < 0) return null;
        var item = lines[visible[row]];
        return (column / 2) switch { 0 => item.Base, 1 => item.Production, _ => item.Incoming };
    }

    private void Render()
    {
        var indices = new List<int>();
        var lastChange = -4;
        var longest = 0;
        for (var i = 0; i < Math.Min(3, lines.Count); i++)
            if (lines[i].Changed) lastChange = i;
        for (var i = 0; i < lines.Count; i++)
        {
            var row = lines[i];
            if (i + 3 < lines.Count && lines[i + 3].Changed) lastChange = i + 3;
            if (!context.Checked && i - lastChange > 3) continue;
            indices.Add(i);
            longest = Math.Max(longest, Math.Max(row.Base?.Text.Length ?? 0,
                Math.Max(row.Production?.Text.Length ?? 0, row.Incoming?.Text.Length ?? 0)));
        }
        visible = indices.ToArray();
        grid.RowCount = visible.Length;
        horizontal.Maximum = Math.Min(longest, 100_000) * Math.Max(8, font.Height);
        horizontal.Visible = horizontal.Maximum > 140;
    }

    private void PaintCode(object? sender, DataGridViewCellPaintingEventArgs e)
    {
        if (e.RowIndex < 0 || e.ColumnIndex % 2 == 0 || e.Graphics is not { } canvas ||
            LineAt(e.RowIndex, e.ColumnIndex) is not { } line) return;
        e.Paint(e.ClipBounds, DataGridViewPaintParts.All & ~DataGridViewPaintParts.ContentForeground);
        var region = canvas.Save();
        canvas.IntersectClip(e.CellBounds);
        var selected = (e.State & DataGridViewElementStates.Selected) != 0;
        TextRenderer.DrawText(canvas, line.Text, font,
            new Point(e.CellBounds.X + 3 - horizontal.Value, e.CellBounds.Y + (e.CellBounds.Height - font.Height) / 2),
            selected ? grid.DefaultCellStyle.SelectionForeColor : Appearance.Ink,
            TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine | TextFormatFlags.PreserveGraphicsClipping);
        canvas.Restore(region);
        e.Handled = true;
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            pending?.Cancel();
            font.Dispose();
        }
        base.Dispose(disposing);
    }
}
