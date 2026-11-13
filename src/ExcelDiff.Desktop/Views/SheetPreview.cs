using System.Drawing.Drawing2D;
using System.Globalization;
using ExcelDiff.Core.Workbooks;

namespace ExcelDiff.Desktop.Views;

internal sealed class SheetPreview : Control
{
    private const int HeaderWidth = 50;
    private const int HeaderHeight = 24;
    private readonly Dictionary<(string Name, float Size, FontStyle Style), Font> fonts = [];
    private (string Name, Rectangle Bounds)[] columns = [];
    private (string Name, Rectangle Bounds)[] rows = [];
    private List<PreviewCell> cells = [];
    private Sheet? sheet;
    private double zoom;
    private int focusRow;
    private int focusColumn;
    private bool highlight;
    private bool date1904;

    internal SheetPreview()
    {
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint, true);
        BackColor = Color.White;
        AccessibleName = "選択した差分の周辺";
    }

    internal void ShowArea(Sheet? source, int row, int column, MergeRange? selected, bool mark, bool savedZoom, bool use1904, double viewScale = 1)
    {
        var scale = (savedZoom ? Math.Clamp(Number(source, "表示倍率", 100), 25, 400) / 100 : 1) * viewScale;
        if (ReferenceEquals(sheet, source) && focusRow == row && focusColumn == column && highlight == mark &&
            zoom == scale && date1904 == use1904 && cells.Count > 0) return;

        if (!ReferenceEquals(sheet, source) || zoom != scale) ClearFonts();
        sheet = source;
        zoom = scale;
        focusRow = row;
        focusColumn = column;
        highlight = mark;
        date1904 = use1904;
        if (source is null)
        {
            columns = [];
            rows = [];
            cells = [];
            Size = new Size(240, 40);
            Invalidate();
            return;
        }

        // Limit work to the context of one selection. Empty rows and columns retain their widths and heights.
        var firstRow = Math.Max(1, row - 1);
        var lastRow = Math.Min(1048576, Math.Min(firstRow + 11, Math.Max(row + 5, selected?.LastRow ?? row)));
        var firstColumn = Math.Max(1, column - 1);
        var lastColumn = Math.Min(16384, Math.Min(firstColumn + 21, Math.Max(column + 12, selected?.LastColumn ?? column)));
        var x = new int[lastColumn - firstColumn + 2];
        var y = new int[lastRow - firstRow + 2];
        x[0] = HeaderWidth;
        y[0] = HeaderHeight;
        var defaultWidth = Number(source, "既定レイアウト / defaultColWidth", 8.43);
        var defaultHeight = Number(source, "既定レイアウト / defaultRowHeight", 15);
        for (var c = firstColumn; c <= lastColumn; c++)
        {
            var name = $"列 {Address.Column(c)}";
            var width = IsHidden(source, name) ? 0 : Pixels((Number(source, name + " / width", defaultWidth) * 7 + 5) * zoom);
            x[c - firstColumn + 1] = x[c - firstColumn] + width;
        }
        for (var r = firstRow; r <= lastRow; r++)
        {
            var name = $"行 {r}";
            var height = IsHidden(source, name) ? 0 : Pixels(Number(source, name + " / ht", defaultHeight) * 96 / 72 * zoom);
            y[r - firstRow + 1] = y[r - firstRow] + height;
        }
        Size = new Size(x[^1] + 1, y[^1] + 1);
        columns = Enumerable.Range(firstColumn, lastColumn - firstColumn + 1).Select(c =>
            (Address.Column(c), Rectangle.FromLTRB(x[c - firstColumn], 0, x[c - firstColumn + 1], HeaderHeight))).ToArray();
        rows = Enumerable.Range(firstRow, lastRow - firstRow + 1).Select(r =>
            (r.ToString(), Rectangle.FromLTRB(0, y[r - firstRow], HeaderWidth, y[r - firstRow + 1]))).ToArray();

        var merged = new MergeRange?[rows.Length, columns.Length];
        var visibleMerges = source.Merges.InRows(firstRow, lastRow);
        foreach (var range in visibleMerges)
        {
            for (var r = Math.Max(firstRow, range.FirstRow); r <= Math.Min(lastRow, range.LastRow); r++)
                for (var c = Math.Max(firstColumn, range.FirstColumn); c <= Math.Min(lastColumn, range.LastColumn); c++)
                    merged[r - firstRow, c - firstColumn] = range;
        }

        cells = [];
        for (var r = firstRow; r <= lastRow; r++)
        {
            var ri = r - firstRow;
            var rowCells = source.Rows.GetValueOrDefault(r);
            bool IsOccupied(int c) => merged[ri, c - firstColumn] is not null ||
                rowCells?.GetValueOrDefault(c) is Cell neighbor && (neighbor.Value.Length != 0 || neighbor.Formula is not null);

            for (var c = firstColumn; c <= lastColumn; c++)
            {
                var ci = c - firstColumn;
                if (merged[ri, ci] is not null || x[ci] == x[ci + 1] || y[ri] == y[ri + 1]) continue;
                var bounds = Rectangle.FromLTRB(x[ci], y[ri], x[ci + 1], y[ri + 1]);
                var cell = rowCells?.GetValueOrDefault(c);
                var spill = bounds;
                if (cell is { Kind: "text", Value.Length: > 0 } &&
                    cell.Format.GetValueOrDefault("配置/wrapText") != "1" &&
                    cell.Format.GetValueOrDefault("配置/shrinkToFit") != "1" &&
                    cell.Format.GetValueOrDefault("配置/textRotation") is null or "0")
                {
                    var align = cell.Format.GetValueOrDefault("配置/horizontal");
                    if (align is null or "general" or "left" or "center")
                    {
                        var end = c;
                        while (end < lastColumn && !IsOccupied(end + 1)) end++;
                        spill.Width = x[end - firstColumn + 1] - spill.Left;
                    }
                    if (align is "center" or "right")
                    {
                        var start = c;
                        while (start > firstColumn && !IsOccupied(start - 1)) start--;
                        spill = Rectangle.FromLTRB(x[start - firstColumn], spill.Top, spill.Right, spill.Bottom);
                    }
                }
                cells.Add(CreateCell(cell, bounds, spill, mark && row == r && column == c));
            }
        }
        foreach (var range in visibleMerges)
        {
            if (range.FirstColumn > lastColumn || range.LastColumn < firstColumn) continue;
            var bounds = Rectangle.FromLTRB(x[Math.Max(0, range.FirstColumn - firstColumn)],
                y[Math.Max(0, range.FirstRow - firstRow)], x[Math.Min(columns.Length, range.LastColumn - firstColumn + 1)],
                y[Math.Min(rows.Length, range.LastRow - firstRow + 1)]);
            if (bounds.Width == 0 || bounds.Height == 0) continue;
            // A merged cell uses the combined width of its columns and clips at the range boundary.
            var anchor = source.At(range.FirstRow, range.FirstColumn);
            var marked = mark && row >= range.FirstRow && row <= range.LastRow && range.ContainsColumn(column);
            cells.Add(CreateCell(anchor, bounds, bounds, marked));
        }
        if (Parent is ScrollableControl viewport)
        {
            var targetX = x[column - firstColumn];
            var targetY = y[row - firstRow];
            viewport.AutoScrollPosition = new Point(
                targetX + 24 > viewport.ClientSize.Width ? targetX - 4 : 0,
                targetY + 16 > viewport.ClientSize.Height ? targetY - 4 : 0);
        }
        Invalidate();
    }

    private PreviewCell CreateCell(Cell? cell, Rectangle bounds, Rectangle spill, bool selected)
    {
        var format = cell?.Format;
        var alignment = format?.GetValueOrDefault("配置/horizontal") switch
        {
            "right" => StringAlignment.Far,
            "center" or "centerContinuous" => StringAlignment.Center,
            null or "general" when cell?.Kind == "n" => StringAlignment.Far,
            _ => StringAlignment.Near
        };
        var vertical = format?.GetValueOrDefault("配置/vertical") switch
        {
            "top" => StringAlignment.Near,
            "center" => StringAlignment.Center,
            _ => StringAlignment.Far
        };
        var background = format?.GetValueOrDefault("背景/patternFill/patternType") == "solid"
            ? ReadColor(format, "背景/patternFill/fgColor/rgb", Color.White) : Color.White;
        return new PreviewCell(bounds, spill, Display(cell, date1904),
            cell is null ? Font : CellFont(cell), ReadColor(format, "フォント/color/rgb", Appearance.Ink),
            selected ? Appearance.Changed : background, alignment, vertical,
            format?.GetValueOrDefault("配置/wrapText") == "1", format?.GetValueOrDefault("配置/shrinkToFit") == "1");
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        if (sheet is null)
        {
            TextRenderer.DrawText(e.Graphics, "対応するシートなし", Font, new Point(8, 8), Appearance.Muted);
            return;
        }

        using var fill = new SolidBrush(Appearance.Panel);
        using var line = new Pen(Appearance.Line);
        var graphics = e.Graphics;
        graphics.FillRectangle(fill, 0, 0, Width, HeaderHeight);
        graphics.FillRectangle(fill, 0, HeaderHeight, HeaderWidth, Height - HeaderHeight);
        foreach (var header in columns.Concat(rows))
        {
            if (header.Bounds.Width == 0 || header.Bounds.Height == 0 || !header.Bounds.IntersectsWith(e.ClipRectangle)) continue;
            TextRenderer.DrawText(graphics, header.Name, Font, header.Bounds, Appearance.Muted,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }

        foreach (var cell in cells)
        {
            if (!cell.Bounds.IntersectsWith(e.ClipRectangle)) continue;
            if (cell.Background != Color.White)
            {
                fill.Color = cell.Background;
                graphics.FillRectangle(fill, cell.Bounds);
            }
            graphics.DrawRectangle(line, cell.Bounds.X, cell.Bounds.Y, cell.Bounds.Width - 1, cell.Bounds.Height - 1);
        }
        foreach (var cell in cells)
            if (cell.Text.Length != 0 && cell.Spill.IntersectsWith(e.ClipRectangle))
                PaintText(graphics, cell, fill);
    }

    internal static void PaintText(Graphics graphics, PreviewCell cell, SolidBrush brush)
    {
        if (cell.Bounds.Width <= 0 || cell.Bounds.Height <= 0) return;
        // At low saved zoom even a visible cell may be only one pixel wide.
        var paddingX = Math.Min(2, (cell.Bounds.Width - 1) / 2);
        var paddingY = Math.Min(1, (cell.Bounds.Height - 1) / 2);
        var inset = Rectangle.Inflate(cell.Bounds, -paddingX, -paddingY);

        var state = graphics.Save();
        Font? smaller = null;
        try
        {
            var clip = Rectangle.Inflate(cell.Spill, -paddingX, -paddingY);
            graphics.SetClip(clip, CombineMode.Intersect);
            using var format = new StringFormat(StringFormat.GenericTypographic)
            {
                Alignment = cell.Alignment, LineAlignment = cell.Vertical, Trimming = StringTrimming.None
            };
            var font = cell.Font;
            var layout = (RectangleF)inset;
            if (!cell.Wrap)
            {
                // The graphics clip already limits the text to the row and the available columns.
                format.FormatFlags = (format.FormatFlags | StringFormatFlags.NoWrap | StringFormatFlags.NoClip)
                    & ~StringFormatFlags.LineLimit;
                var width = graphics.MeasureString(cell.Text, font, PointF.Empty, StringFormat.GenericTypographic).Width + 1;
                if (cell.Shrink && width > inset.Width)
                {
                    smaller = new Font(font.FontFamily, Math.Max(1, font.Size * inset.Width / width), font.Style);
                    font = smaller;
                }
                else if (width > inset.Width)
                {
                    // Measure text from its original cell and only enlarge the clipping area into available neighbors.
                    layout.Width = width;
                    layout.X -= cell.Alignment switch
                    {
                        StringAlignment.Far => width - inset.Width,
                        StringAlignment.Center => (width - inset.Width) / 2,
                        _ => 0
                    };
                }
            }
            brush.Color = cell.Foreground;
            graphics.DrawString(cell.Text, font, brush, layout, format);
        }
        finally
        {
            smaller?.Dispose();
            graphics.Restore(state);
        }
    }

    private Font CellFont(Cell cell)
    {
        var format = cell.Format;
        var name = format.GetValueOrDefault("フォント/name/val") ?? Font.Name;
        var size = float.TryParse(format.GetValueOrDefault("フォント/sz/val"), CultureInfo.InvariantCulture, out var points) &&
            float.IsFinite(points) ? points : Font.Size;
        var style = FontStyle.Regular;
        if (FontFlag(format, "b")) style |= FontStyle.Bold;
        if (FontFlag(format, "i")) style |= FontStyle.Italic;
        if (FontFlag(format, "u")) style |= FontStyle.Underline;
        if (FontFlag(format, "strike")) style |= FontStyle.Strikeout;
        var key = (Name: name, Size: (float)Math.Clamp(size * zoom, 2, 120), Style: style);
        if (!fonts.TryGetValue(key, out var font))
        {
            try { font = new Font(key.Name, key.Size, key.Style); }
            catch (ArgumentException) { font = new Font(Font.FontFamily, key.Size); }
            fonts.Add(key, font);
        }
        return font;
    }

    internal static bool FontFlag(IReadOnlyDictionary<string, string> format, string key) =>
        format.GetValueOrDefault($"フォント/{key}") == "1" ||
        format.GetValueOrDefault($"フォント/{key}/val") is "1" or "single" or "double";

    private static double Number(Sheet? source, string key, double fallback) =>
        double.TryParse(source?.Settings.GetValueOrDefault(key), CultureInfo.InvariantCulture, out var value) &&
        double.IsFinite(value) ? value : fallback;

    private static bool IsHidden(Sheet sheet, string name) => sheet.Settings.GetValueOrDefault(name + " / hidden") == "1";
    private static int Pixels(double value) => (int)Math.Round(Math.Clamp(value, 0, 4096));

    internal static Color ReadColor(IReadOnlyDictionary<string, string>? values, string key, Color fallback)
    {
        var hex = values?.GetValueOrDefault(key);
        return hex is { Length: >= 6 } &&
            int.TryParse(hex.AsSpan(hex.Length - 6), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var rgb)
            ? Color.FromArgb(unchecked((int)0xFF000000) | rgb) : fallback;
    }

    internal static string Display(Cell? cell, bool date1904)
    {
        if (cell is null) return "";
        if (cell.Kind != "n" || !double.TryParse(cell.Value, CultureInfo.InvariantCulture, out var number)) return cell.Display;
        var format = cell.Format.GetValueOrDefault("表示形式");
        if (format is "組み込み 14" or "組み込み 15" or "組み込み 16" or "組み込み 17" or "組み込み 22")
        {
            try
            {
                if (!date1904 && number >= 60 && number < 61) return "1900/02/29";
                var serial = date1904 ? number + 1462 : number < 60 ? number + 1 : number;
                return DateTime.FromOADate(serial).ToString(format == "組み込み 22" ? "yyyy/MM/dd HH:mm" : "yyyy/MM/dd");
            }
            catch (ArgumentException) { return cell.Display; }
        }
        return format switch
        {
            "組み込み 1" => number.ToString("0"),
            "組み込み 2" => number.ToString("0.00"),
            "組み込み 3" => number.ToString("#,##0"),
            "組み込み 4" => number.ToString("#,##0.00"),
            "組み込み 9" => number.ToString("0%"),
            "組み込み 10" => number.ToString("0.00%"),
            _ => cell.Display
        };
    }

    private void ClearFonts()
    {
        foreach (var font in fonts.Values) font.Dispose();
        fonts.Clear();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing) ClearFonts();
        base.Dispose(disposing);
    }

    internal sealed record PreviewCell(Rectangle Bounds, Rectangle Spill, string Text, Font Font,
        Color Foreground, Color Background, StringAlignment Alignment, StringAlignment Vertical, bool Wrap, bool Shrink);
}
