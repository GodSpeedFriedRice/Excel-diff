using ExcelDiff.Core.Comparison;
using ExcelDiff.Core.Workbooks;
using System.Text.RegularExpressions;

namespace ExcelDiff.Desktop.Views;

internal sealed class XmlDiffView : UserControl
{
    private const int ColumnsPerSide = 3;
    private const int CodeColumn = 2;
    private static readonly Color Added = Color.FromArgb(230, 255, 236);
    private static readonly Color Removed = Color.FromArgb(255, 235, 233);
    private static readonly Color AddedNumber = Color.FromArgb(204, 255, 216);
    private static readonly Color RemovedNumber = Color.FromArgb(255, 215, 213);
    private static readonly Color AddedWord = Color.FromArgb(171, 242, 188);
    private static readonly Color RemovedWord = Color.FromArgb(255, 171, 168);
    private static readonly Color Empty = Color.FromArgb(246, 248, 250);
    private static readonly Color Hunk = Color.FromArgb(221, 244, 255);
    private static readonly Color HunkGutter = Color.FromArgb(182, 227, 255);
    private static readonly Color Divider = Color.FromArgb(208, 215, 222);
    private static readonly Color Muted = Color.FromArgb(87, 96, 106);
    private static readonly Color Ink = Color.FromArgb(31, 35, 40);
    private static readonly Color TagColor = Color.FromArgb(17, 99, 41);
    private static readonly Color Attribute = Color.FromArgb(5, 80, 174);
    private static readonly Color Literal = Color.FromArgb(10, 48, 105);
    private static readonly Regex XmlTokens = new("""<!--.*?(?:-->|$)|</?[\w.:-]+|"[^"]*"|'[^']*'|[\w.:-]+(?=\s*=)|[<>/=?!]""",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private const TextFormatFlags CodeFlags = TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix |
        TextFormatFlags.SingleLine | TextFormatFlags.PreserveGraphicsClipping;
    private readonly DataGridView grid = Appearance.Grid();
    private readonly HScrollBar horizontal = new() { Dock = DockStyle.Fill, SmallChange = 28, Visible = false };
    private readonly CheckBox showAll = new() { Text = "変更のない行も表示", AutoSize = true };
    private readonly Label summary = Appearance.Label("");
    private readonly Label fileHeader = new() { Dock = DockStyle.Fill, AutoSize = false, Height = 28,
        Padding = new Padding(9, 0, 8, 0), Margin = Padding.Empty, TextAlign = ContentAlignment.MiddleLeft,
        AutoEllipsis = true, BackColor = Empty, ForeColor = Ink };
    private readonly Font codeFont = new("Consolas", 9);
    private CancellationTokenSource? pending;
    private Sheet? before;
    private Sheet? after;
    private XmlLineDiff? diff;
    private readonly List<SplitRow> visible = [];
    private readonly List<int> changeStarts = [];
    private readonly HashSet<int> expandedHunks = [];
    private int longestLine;

    internal XmlDiffView()
    {
        Dock = DockStyle.Fill;
        var layout = Appearance.Stack();
        var toolbar = Appearance.Toolbar();
        toolbar.BackColor = Color.White;
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
        grid.CellBorderStyle = DataGridViewCellBorderStyle.None;
        grid.ColumnHeadersVisible = false;
        grid.AllowUserToResizeColumns = false;
        grid.ScrollBars = ScrollBars.Vertical;
        grid.RowTemplate.Height = Math.Max(20, codeFont.Height + 3);
        AddSide("変更前のXML");
        AddSide("変更後のXML");
        grid.CellValueNeeded += FillLine;
        grid.CellFormatting += ColorLine;
        grid.CellPainting += PaintDiffCell;
        grid.CellClick += (_, e) => { if (e.RowIndex >= 0) ExpandHunk(e.RowIndex); };
        grid.KeyDown += (_, e) =>
        {
            if (e.KeyCode != Keys.Enter || grid.CurrentCell is not { } cell) return;
            e.Handled = e.SuppressKeyPress = ExpandHunk(cell.RowIndex);
        };
        grid.SizeChanged += (_, _) => UpdateHorizontalScroll();
        horizontal.ValueChanged += (_, _) => grid.Invalidate();
        showAll.CheckedChanged += (_, _) => Render();
        var file = Appearance.Stack();
        file.BackColor = Divider;
        file.Padding = new Padding(1);
        file.Margin = new Padding(8, 0, 8, 8);
        Appearance.Row(file, fileHeader);
        Appearance.Row(file, grid, fill: true);
        Appearance.Row(file, horizontal);
        Appearance.Row(layout, file, fill: true);
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
            Width = 20, SortMode = DataGridViewColumnSortMode.NotSortable,
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
        var oldPath = left?.XmlPartPath;
        var newPath = right?.XmlPartPath;
        fileHeader.Text = oldPath is not null && newPath is not null && oldPath != newPath
            ? $"{oldPath} → {newPath}" : oldPath ?? newPath ?? left?.Name ?? right?.Name ?? "";
        expandedHunks.Clear();
        horizontal.Value = 0;
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
        longestLine = 0;
        UpdateHorizontalScroll();
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
        else if (blocks.Count > 0)
        {
            var precedingEnd = 0;
            var oldLine = 0;
            var newLine = 0;
            foreach (var (first, end) in blocks)
            {
                for (var i = precedingEnd; i < first; i++)
                {
                    if (lines[i].Before is int oldNumber) oldLine = oldNumber;
                    if (lines[i].After is int newNumber) newLine = newNumber;
                }
                var oldCount = 0;
                var newCount = 0;
                for (var i = first; i < end; i++)
                {
                    if (lines[i].Before.HasValue) oldCount++;
                    if (lines[i].After.HasValue) newCount++;
                }
                var oldStart = oldCount == 0 ? oldLine : oldLine + 1;
                var newStart = newCount == 0 ? newLine : newLine + 1;
                var hidden = first - precedingEnd;
                var expanded = expandedHunks.Contains(first);
                visible.Add(new SplitRow(null, null, $"@@ -{oldStart},{oldCount} +{newStart},{newCount} @@", first, hidden, expanded));
                if (expanded && hidden > 0) AppendRows(lines, precedingEnd, first);
                AppendRows(lines, first, end);
                oldLine += oldCount;
                newLine += newCount;
                precedingEnd = end;
            }
            if (precedingEnd < lines.Count)
            {
                var expanded = expandedHunks.Contains(lines.Count);
                visible.Add(new SplitRow(null, null, "残りの行", lines.Count, lines.Count - precedingEnd, expanded));
                if (expanded) AppendRows(lines, precedingEnd, lines.Count);
            }
        }
        foreach (var row in visible)
        {
            longestLine = Math.Max(longestLine, DisplayLength(row.Before?.Text));
            longestLine = Math.Max(longestLine, DisplayLength(row.After?.Text));
        }
        grid.RowCount = visible.Count;
        grid.ClearSelection();
        UpdateHorizontalScroll();
    }

    private bool ExpandHunk(int index)
    {
        if (index >= visible.Count || visible[index] is not { Hidden: > 0 } row) return false;
        var first = grid.FirstDisplayedScrollingRowIndex;
        var offset = horizontal.Value;
        if (!expandedHunks.Add(row.Anchor)) expandedHunks.Remove(row.Anchor);
        Render();
        var heading = visible.FindIndex(candidate => candidate.Anchor == row.Anchor);
        if (heading >= 0 && first >= 0)
            grid.FirstDisplayedScrollingRowIndex = Math.Clamp(heading - (index - first), 0, grid.RowCount - 1);
        horizontal.Value = Math.Clamp(offset, 0, Math.Max(0, horizontal.Maximum - horizontal.LargeChange + 1));
        return true;
    }

    private static int DisplayLength(string? text) => text is null ? 0 : text.Length + text.Count(c => c > 127);

    private void UpdateHorizontalScroll()
    {
        if (grid.Columns.Count < 2 * ColumnsPerSide) return;
        var available = Math.Max(1, Math.Min(grid.Columns[CodeColumn].Width, grid.Columns[ColumnsPerSide + CodeColumn].Width) - 16);
        var glyph = TextRenderer.MeasureText("0", codeFont, Size.Empty, CodeFlags).Width;
        var width = (int)Math.Min(10_000_000L, (long)longestLine * Math.Max(1, glyph));
        var needed = width > available;
        horizontal.LargeChange = available;
        horizontal.Maximum = Math.Max(0, width - 1);
        horizontal.Value = Math.Min(horizontal.Value, Math.Max(0, horizontal.Maximum - horizontal.LargeChange + 1));
        horizontal.Visible = needed;
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
            {
                var oldLine = offset < removedCount ? lines[removedStart + offset] : null;
                var newLine = offset < addedCount ? lines[addedStart + offset] : null;
                var (oldSpan, newSpan) = oldLine is not null && newLine is not null
                    ? ChangedText(oldLine.Text, newLine.Text) : default;
                visible.Add(new SplitRow(oldLine, newLine, OldSpan: oldSpan, NewSpan: newSpan));
            }
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
            1 => line.Kind == XmlLineKind.Added ? "+" : line.Kind == XmlLineKind.Removed ? "-" : "",
            _ => line.Text
        };
    }

    private void ColorLine(object? sender, DataGridViewCellFormattingEventArgs e)
    {
        if (e.RowIndex < 0 || e.RowIndex >= visible.Count || e.CellStyle is not { } style) return;
        var row = visible[e.RowIndex];
        var line = e.ColumnIndex < ColumnsPerSide ? row.Before : row.After;
        var column = e.ColumnIndex % ColumnsPerSide;
        style.BackColor = row.Heading is not null ? column == 0 ? HunkGutter : Hunk : line?.Kind switch
        {
            XmlLineKind.Added => column == 0 ? AddedNumber : Added,
            XmlLineKind.Removed => column == 0 ? RemovedNumber : Removed,
            XmlLineKind.Context => Color.White,
            _ => Empty
        };
        style.ForeColor = column == 0 || row.Heading is not null ? Muted : Ink;
        style.SelectionBackColor = style.BackColor;
        style.SelectionForeColor = style.ForeColor;
    }

    private void PaintDiffCell(object? sender, DataGridViewCellPaintingEventArgs e)
    {
        if (e.RowIndex < 0 || e.RowIndex >= visible.Count || e.Graphics is not { } graphics || e.CellStyle is not { } style) return;
        var row = visible[e.RowIndex];
        var column = e.ColumnIndex % ColumnsPerSide;
        var bounds = e.CellBounds;
        if (column != CodeColumn && row.Heading is null) e.Paint(e.ClipBounds, e.PaintParts);
        else
        {
            using var brush = new SolidBrush(style.BackColor);
            graphics.FillRectangle(brush, bounds);
            if (row.Heading is not null)
            {
                if (e.ColumnIndex == CodeColumn)
                    TextRenderer.DrawText(graphics, row.Heading, codeFont, bounds, Muted, CodeFlags | TextFormatFlags.VerticalCenter);
                if (e.ColumnIndex == 0 && row.Hidden > 0)
                    TextRenderer.DrawText(graphics, row.Expanded ? "▴" : "⋯", Font, bounds, Muted,
                        TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                if (e.ColumnIndex == ColumnsPerSide + CodeColumn && row.Hidden > 0)
                    TextRenderer.DrawText(graphics, row.Expanded ? "行を折りたたむ" : $"{row.Hidden} 行を展開", Font, bounds, Muted,
                        TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
            }
            else
            {
                var left = e.ColumnIndex < ColumnsPerSide;
                var line = left ? row.Before : row.After;
                if (line is not null)
                    DrawCode(graphics, bounds, line, left ? row.OldSpan : row.NewSpan);
            }
        }
        if (e.ColumnIndex is 0 or CodeColumn or ColumnsPerSide)
        {
            using var pen = new Pen(Divider);
            var edge = bounds.Right - 1;
            graphics.DrawLine(pen, edge, bounds.Top, edge, bounds.Bottom - 1);
        }
        if (e.ColumnIndex == CodeColumn || e.ColumnIndex == ColumnsPerSide + CodeColumn)
            if ((e.State & DataGridViewElementStates.Selected) != 0 && grid.Focused)
                ControlPaint.DrawFocusRectangle(graphics, Rectangle.Inflate(bounds, -1, -1));
        e.Handled = true;
    }

    private void DrawCode(Graphics graphics, Rectangle bounds, XmlDiffLine line, TextSpan change)
    {
        var state = graphics.Save();
        graphics.IntersectClip(Rectangle.Inflate(bounds, -3, 0));
        var x = bounds.Left + 5 - horizontal.Value;
        var y = bounds.Top + (bounds.Height - codeFont.Height) / 2;
        if (line.Text.Length > 2048)
            DrawPart(0, line.Text.Length, Ink);
        else
        {
            var start = 0;
            foreach (Match token in XmlTokens.Matches(line.Text))
            {
                DrawPart(start, token.Index, Ink);
                var color = token.Value[0] switch
                {
                    '"' or '\'' => Literal,
                    '<' when token.Value.StartsWith("<!--", StringComparison.Ordinal) => Muted,
                    '<' or '>' or '/' => TagColor,
                    _ when char.IsLetter(token.Value[0]) => Attribute,
                    _ => Ink
                };
                DrawPart(token.Index, token.Index + token.Length, color);
                start = token.Index + token.Length;
            }
            DrawPart(start, line.Text.Length, Ink);
        }
        graphics.Restore(state);

        void DrawPart(int start, int end, Color color)
        {
            if (start >= end || x >= bounds.Right) return;
            var first = Math.Max(start, change.Start);
            var last = Math.Min(end, change.End);
            if (first >= last) DrawText(start, end, color, null);
            else
            {
                DrawText(start, first, color, null);
                DrawText(first, last, color, line.Kind == XmlLineKind.Added ? AddedWord : RemovedWord);
                DrawText(last, end, color, null);
            }
        }

        void DrawText(int start, int end, Color color, Color? background)
        {
            if (end <= start || x >= bounds.Right) return;
            var part = line.Text[start..end];
            var width = TextRenderer.MeasureText(graphics, part, codeFont, Size.Empty, CodeFlags).Width;
            if (x + width > bounds.Left)
            {
                if (background is Color fill)
                {
                    using var brush = new SolidBrush(fill);
                    graphics.FillRectangle(brush, x, y, width, codeFont.Height);
                }
                TextRenderer.DrawText(graphics, part, codeFont, new Point(x, y), color, CodeFlags);
            }
            x += width;
        }
    }

    private static (TextSpan Before, TextSpan After) ChangedText(string before, string after)
    {
        var common = Math.Min(before.Length, after.Length);
        var prefix = 0;
        while (prefix < common && before[prefix] == after[prefix]) prefix++;
        var suffix = 0;
        while (suffix < common - prefix && before[^(1 + suffix)] == after[^(1 + suffix)]) suffix++;
        if (prefix == 0 && suffix == 0) return default;
        return (new TextSpan(prefix, before.Length - suffix), new TextSpan(prefix, after.Length - suffix));
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

    private readonly record struct TextSpan(int Start, int End);
    private readonly record struct SplitRow(XmlDiffLine? Before, XmlDiffLine? After, string? Heading = null,
        int Anchor = -1, int Hidden = 0, bool Expanded = false, TextSpan OldSpan = default, TextSpan NewSpan = default);
}
