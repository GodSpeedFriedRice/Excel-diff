using System.Drawing.Imaging;
using System.Globalization;
using ExcelDiff.Core.Workbooks;
using PreviewCell = ExcelDiff.Desktop.Views.SheetPreview.PreviewCell;

namespace ExcelDiff.Desktop.Views;

internal readonly record struct PreviewTile(int Column, int Row);

internal sealed class SheetRaster : IDisposable
{
    internal const int TileSize = 512;
    private readonly Sheet sheet;
    private readonly Font defaultFont;
    private readonly double zoom;
    private readonly bool date1904;
    private readonly Dictionary<(string Name, float Size, FontStyle Style), Font> fonts = [];
    private readonly Dictionary<int, int[]> occupiedRows = [];

    internal SheetRaster(Sheet sheet, Font defaultFont, bool savedZoom, bool date1904)
    {
        this.sheet = sheet;
        this.defaultFont = (Font)defaultFont.Clone();
        this.date1904 = date1904;
        zoom = savedZoom && double.TryParse(sheet.Settings.GetValueOrDefault("表示倍率"),
            CultureInfo.InvariantCulture, out var percent) && double.IsFinite(percent) ? Math.Clamp(percent, 25, 400) / 100 : 1;
        Rows = new PreviewAxis(sheet, rows: true, zoom);
        Columns = new PreviewAxis(sheet, rows: false, zoom);
        var (lastRow, lastColumn) = sheet.UsedCellExtent ?? (sheet.Rows.Keys.DefaultIfEmpty().Max(),
            sheet.Rows.Values.SelectMany(cells => cells.Keys).DefaultIfEmpty().Max());
        var mergedRow = sheet.Merges.Ranges.Select(range => range.LastRow).DefaultIfEmpty().Max();
        var mergedColumn = sheet.Merges.Ranges.Select(range => range.LastColumn).DefaultIfEmpty().Max();
        LastRow = Math.Min(Rows.Limit, Math.Max(30, Math.Max(Rows.LastSized, Math.Max(lastRow, mergedRow)) + 16));
        LastColumn = Math.Min(Columns.Limit, Math.Max(18, Math.Max(Columns.LastSized, Math.Max(lastColumn, mergedColumn)) + 8));
    }

    internal PreviewAxis Rows { get; }
    internal PreviewAxis Columns { get; }
    internal int LastRow { get; }
    // Keep horizontal navigation near the used range, while allowing overflow into visible blank columns.
    internal int LastColumn { get; }
    internal int TileColumns => TileCount(Columns.Position(Columns.Limit + 1));
    internal int UsedTileColumns => TileCount(Columns.Position(LastColumn + 1));
    internal int TileRows => TileCount(Rows.Position(LastRow + 1));

    private static int TileCount(long end) => (int)((end + TileSize - 1) / TileSize);

    internal Bitmap Render(PreviewTile tile)
    {
        var bitmap = new Bitmap(TileSize, TileSize, PixelFormat.Format24bppRgb);
        try
        {
            using var graphics = Graphics.FromImage(bitmap);
            graphics.Clear(Color.White);
            using var pen = new Pen(Appearance.Line);
            using var brush = new SolidBrush(Color.White);
            var originX = (long)tile.Column * TileSize;
            var originY = (long)tile.Row * TileSize;
            var firstColumn = Columns.AtPixel(originX);
            var lastColumn = Columns.AtPixel(originX + TileSize - 1);
            var firstRow = Rows.AtPixel(originY);
            var lastRow = Math.Min(LastRow, Rows.AtPixel(originY + TileSize - 1));
            var columns = Columns.Visible(firstColumn, lastColumn, int.MaxValue);
            var rows = Rows.Visible(firstRow, lastRow, int.MaxValue);
            var renderedMerges = new HashSet<MergeRange>();
            var text = new List<PreviewCell>();

            int X(int column) => (int)Math.Clamp(Columns.Position(column) - originX, -1000000, 1000000);
            int Y(int row) => (int)Math.Clamp(Rows.Position(row) - originY, -1000000, 1000000);
            void Paint(Cell? cell, int row, int column, Rectangle bounds, Rectangle spill)
            {
                if (bounds.Width <= 0 || bounds.Height <= 0) return;
                var format = cell?.Format ?? sheet.InheritedFormat(row, column);
                var fill = format.GetValueOrDefault("背景/patternFill/patternType") == "solid"
                    ? SheetPreview.ReadColor(format, "背景/patternFill/fgColor/rgb", Color.White) : Color.White;
                if (bounds.IntersectsWith(new Rectangle(0, 0, TileSize, TileSize)))
                {
                    if (fill != Color.White)
                    {
                        brush.Color = fill;
                        graphics.FillRectangle(brush, bounds);
                    }
                    graphics.DrawRectangle(pen, bounds.X, bounds.Y, bounds.Width - 1, bounds.Height - 1);
                }
                var value = SheetPreview.Display(cell, date1904);
                if (value.Length == 0) return;
                var horizontal = format.GetValueOrDefault("配置/horizontal") switch
                {
                    "right" => StringAlignment.Far,
                    "center" or "centerContinuous" => StringAlignment.Center,
                    null or "general" when cell?.Kind == "n" => StringAlignment.Far,
                    _ => StringAlignment.Near
                };
                var vertical = format.GetValueOrDefault("配置/vertical") switch
                {
                    "top" => StringAlignment.Near, "center" => StringAlignment.Center, _ => StringAlignment.Far
                };
                text.Add(new PreviewCell(bounds, spill, value, GetFont(format),
                    SheetPreview.ReadColor(format, "フォント/color/rgb", Appearance.Ink), fill, horizontal, vertical,
                    format.GetValueOrDefault("配置/wrapText") == "1", format.GetValueOrDefault("配置/shrinkToFit") == "1"));
            }

            foreach (var row in rows)
            {
                var rowCells = sheet.Rows.GetValueOrDefault(row);
                var occupied = Occupied(row, rowCells);
                var merges = sheet.Merges.InRows(row, row);
                if (merges.Count > 1) merges.Sort((a, b) => a.FirstColumn.CompareTo(b.FirstColumn));
                foreach (var column in columns)
                {
                    var mergeIndex = MergePosition(merges, column);
                    if (mergeIndex >= 0)
                    {
                        var range = merges[mergeIndex];
                        if (renderedMerges.Add(range))
                        {
                            var bounds = Rectangle.FromLTRB(X(range.FirstColumn), Y(range.FirstRow),
                                X(range.LastColumn + 1), Y(range.LastRow + 1));
                            Paint(sheet.At(range.FirstRow, range.FirstColumn), range.FirstRow, range.FirstColumn, bounds, bounds);
                        }
                        continue;
                    }
                    DrawNormal(column, ~mergeIndex);
                }

                // Text may start on either side of a tile boundary and continue into the visible tile.
                var preceding = OccupiedNeighbor(occupied, firstColumn, forward: false);
                var following = OccupiedNeighbor(occupied, lastColumn, forward: true);
                if (preceding is int previous && MergePosition(merges, previous) is < 0 and var before)
                    DrawNormal(previous, ~before, outside: true);
                if (following is int next && MergePosition(merges, next) is < 0 and var after)
                    DrawNormal(next, ~after, outside: true);

                void DrawNormal(int column, int mergeInsertion, bool outside = false)
                {
                    var cell = rowCells?.GetValueOrDefault(column);
                    var bounds = Rectangle.FromLTRB(X(column), Y(row), X(column + 1), Y(row + 1));
                    if (bounds.Width <= 0) return;
                    var spill = bounds;
                    if (cell is { Kind: "text", Value.Length: > 0 } &&
                        cell.Format.GetValueOrDefault("配置/wrapText") != "1" &&
                        cell.Format.GetValueOrDefault("配置/shrinkToFit") != "1" &&
                        cell.Format.GetValueOrDefault("配置/textRotation") is null or "0")
                    {
                        var align = cell.Format.GetValueOrDefault("配置/horizontal");
                        if (align is null or "general" or "left" or "center")
                        {
                            var end = OccupiedNeighbor(occupied, column, forward: true) ?? Columns.Limit + 1;
                            if (mergeInsertion < merges.Count) end = Math.Min(end, merges[mergeInsertion].FirstColumn);
                            spill.Width = X(end) - spill.Left;
                        }
                        if (align is "center" or "right")
                        {
                            var start = (OccupiedNeighbor(occupied, column, forward: false) ?? 0) + 1;
                            if (mergeInsertion > 0) start = Math.Max(start, merges[mergeInsertion - 1].LastColumn + 1);
                            spill = Rectangle.FromLTRB(X(start), spill.Top, spill.Right, spill.Bottom);
                        }
                    }
                    if (!outside || spill.Right > 0 && spill.Left < TileSize)
                        Paint(cell, row, column, bounds, spill);
                }
            }
            foreach (var cell in text)
                SheetPreview.PaintText(graphics, cell, brush);
            return bitmap;
        }
        catch
        {
            bitmap.Dispose();
            throw;
        }
    }

    private int[] Occupied(int row, IReadOnlyDictionary<int, Cell>? cells)
    {
        if (cells is null) return [];
        if (occupiedRows.TryGetValue(row, out var columns)) return columns;
        if (occupiedRows.Count >= 256) occupiedRows.Clear();
        return occupiedRows[row] = cells.Where(pair => pair.Value.Value.Length > 0 || pair.Value.Formula is not null)
            .Select(pair => pair.Key).Order().ToArray();
    }

    private static int? OccupiedNeighbor(int[] columns, int column, bool forward)
    {
        var index = Array.BinarySearch(columns, column);
        index = index < 0 ? ~index : index + (forward ? 1 : 0);
        if (!forward) index--;
        return index >= 0 && index < columns.Length ? columns[index] : null;
    }

    private static int MergePosition(List<MergeRange> ranges, int column)
    {
        var low = 0;
        var high = ranges.Count;
        while (low < high)
        {
            var middle = (low + high) / 2;
            var range = ranges[middle];
            if (column < range.FirstColumn) high = middle;
            else if (column > range.LastColumn) low = middle + 1;
            else return middle;
        }
        return ~low;
    }

    private Font GetFont(IReadOnlyDictionary<string, string> format)
    {
        var name = format.GetValueOrDefault("フォント/name/val") ?? defaultFont.Name;
        var size = float.TryParse(format.GetValueOrDefault("フォント/sz/val"), CultureInfo.InvariantCulture, out var points) &&
            float.IsFinite(points) ? points : defaultFont.Size;
        var style = FontStyle.Regular;
        if (SheetPreview.FontFlag(format, "b")) style |= FontStyle.Bold;
        if (SheetPreview.FontFlag(format, "i")) style |= FontStyle.Italic;
        if (SheetPreview.FontFlag(format, "u")) style |= FontStyle.Underline;
        if (SheetPreview.FontFlag(format, "strike")) style |= FontStyle.Strikeout;
        var key = (name, (float)Math.Clamp(size * zoom, 2, 120), style);
        if (fonts.TryGetValue(key, out var font)) return font;
        try { font = new Font(key.name, key.Item2, key.style); }
        catch (ArgumentException) { font = new Font(defaultFont.FontFamily, key.Item2); }
        fonts.Add(key, font);
        return font;
    }

    public void Dispose()
    {
        foreach (var font in fonts.Values) font.Dispose();
        defaultFont.Dispose();
    }
}
