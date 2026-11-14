using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using ExcelDiff.Core.Comparison;
using ExcelDiff.Core.Workbooks;

namespace ExcelDiff.Desktop.Views;

internal sealed class FullSheetPane : UserControl
{
    private readonly Surface surface;
    private IReadOnlyList<Difference>? indexedDifferences;
    private PreviewChanges? marks;
    private readonly VScrollBar vertical = new() { Dock = DockStyle.Fill, Minimum = 1, LargeChange = 18, Enabled = false };
    private readonly HScrollBar horizontal = new() { Dock = DockStyle.Fill, Minimum = 1, LargeChange = 8, Enabled = false };
    private TileSession? session;
    private Sheet? source;
    private bool savedZoom;
    private bool showGridlines = true;
    private double scale = 1;
    private bool date1904;
    private bool moving;
    private int row = 1;
    private int column = 1;
    private int wheelRemainder;
    private int highlightedRow;
    private int highlightedColumn;
    private bool markRow;
    private bool markColumn;
    private int layoutCount;
    private MergeRange? highlightedMerge;

    internal FullSheetPane()
    {
        Dock = DockStyle.Fill;
        surface = new Surface(this) { Dock = DockStyle.Fill, AccessibleName = "全体プレビュー" };
        var layout = new TableLayoutPanel { ColumnCount = 2, RowCount = 2, Dock = DockStyle.Fill, Margin = Padding.Empty };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, SystemInformation.VerticalScrollBarWidth));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, SystemInformation.HorizontalScrollBarHeight));
        foreach (var control in new Control[] { surface, vertical, horizontal }) control.Margin = Padding.Empty;
        layout.Controls.Add(surface, 0, 0);
        layout.Controls.Add(vertical, 1, 0);
        layout.Controls.Add(horizontal, 0, 1);
        Controls.Add(layout);
        vertical.ValueChanged += (_, _) => ScrollbarChanged();
        horizontal.ValueChanged += (_, _) => ScrollbarChanged();
        surface.Resize += (_, _) => RefreshVisible();
        surface.HandleCreated += (_, _) => RefreshVisible();
        surface.VisibleChanged += (_, _) => RefreshVisible();
    }

    internal event Action<int, int>? Scrolled;
    internal event Action<int>? ZoomRequested;

    internal void SetZoom(double zoom)
    {
        if (scale == zoom) return;
        scale = zoom;
        RefreshVisible();
    }

    internal void SetGridlines(bool visible)
    {
        if (showGridlines == visible) return;
        showGridlines = visible;
        if (source is null) return;
        ResetSession();
        RefreshVisible();
    }
    internal event Action<CellPick>? Picked;

    internal void SetChanges(IReadOnlyList<Difference>? differences, bool before)
    {
        if (ReferenceEquals(indexedDifferences, differences)) return;
        indexedDifferences = differences;
        marks = differences is { Count: > 0 } ? new PreviewChanges(differences, before) : null;
        surface.Invalidate();
    }

    internal void ShowSheet(Sheet? sheet, int targetRow, int targetColumn, MergeRange? merge,
        bool rowLayout, bool columnLayout, bool useSavedZoom, bool use1904, bool jump, bool mark, int count)
    {
        var sheetChanged = !ReferenceEquals(source, sheet);
        if (sheetChanged || savedZoom != useSavedZoom || date1904 != use1904)
        {
            source = sheet;
            savedZoom = useSavedZoom;
            date1904 = use1904;
            ResetSession();
            moving = true;
            vertical.Maximum = (session?.Raster.LastRow ?? 1) + vertical.LargeChange - 1;
            horizontal.Maximum = (session?.Raster.LastColumn ?? 1) + horizontal.LargeChange - 1;
            vertical.Enabled = horizontal.Enabled = session is not null;
            moving = false;
        }

        highlightedRow = mark ? targetRow : 0;
        highlightedColumn = mark ? targetColumn : 0;
        layoutCount = count;
        markRow = rowLayout;
        markColumn = columnLayout;
        highlightedMerge = mark ? merge : null;
        if (sheetChanged) SetPosition(1, 1);
        if (jump && session is { } active)
        {
            var top = merge?.FirstRow ?? targetRow;
            var left = merge?.FirstColumn ?? targetColumn;
            SetPosition(
                columnLayout ? row : Reveal(active.Raster.Rows, row, top, SourcePixels(surface.Height - HeaderHeight)),
                rowLayout ? column : Reveal(active.Raster.Columns, column, left, SourcePixels(surface.Width - HeaderWidth)));
        }
        else if (!sheetChanged) SetPosition(row, column);
    }

    private static int Reveal(PreviewAxis axis, int first, int target, int available)
    {
        if (target >= first && (target == first || axis.Position(target + 1) - axis.Position(first) <= available))
            return first;
        return Math.Max(1, target - 1);
    }

    internal void FocusSurface() => surface.Focus();

    internal void ClearSelection()
    {
        highlightedRow = 0;
        highlightedColumn = 0;
        highlightedMerge = null;
        surface.Invalidate();
    }

    internal void ScrollBy(int rows, int columns) => SetPosition(row + rows, column + columns);

    private void SetPosition(int top, int left)
    {
        moving = true;
        row = vertical.Value = Math.Clamp(top, 1, vertical.Maximum - vertical.LargeChange + 1);
        column = horizontal.Value = Math.Clamp(left, 1, horizontal.Maximum - horizontal.LargeChange + 1);
        moving = false;
        RefreshVisible();
    }

    private void ScrollbarChanged()
    {
        if (moving) return;
        var rows = vertical.Value - row;
        var columns = horizontal.Value - column;
        row = vertical.Value;
        column = horizontal.Value;
        RefreshVisible();
        if (rows != 0 || columns != 0) Scrolled?.Invoke(rows, columns);
    }

    private void Wheel(MouseEventArgs e)
    {
        if ((ModifierKeys & Keys.Control) != 0)
        {
            if (e is HandledMouseEventArgs handled) handled.Handled = true;
            ZoomRequested?.Invoke(e.Delta);
            return;
        }

        wheelRemainder += e.Delta;
        var ticks = wheelRemainder / SystemInformation.MouseWheelScrollDelta;
        if (ticks == 0) return;
        wheelRemainder %= SystemInformation.MouseWheelScrollDelta;
        var step = SystemInformation.MouseWheelScrollLines;
        if (step == 0) return;
        var previousRow = row;
        var previousColumn = column;
        if ((ModifierKeys & Keys.Shift) != 0)
            ScrollBy(0, -ticks * (step < 0 ? horizontal.LargeChange : step));
        else
            ScrollBy(-ticks * (step < 0 ? vertical.LargeChange : step), 0);
        Scrolled?.Invoke(row - previousRow, column - previousColumn);
    }

    private const int HeaderWidth = 52;
    private const int HeaderHeight = 24;

    private void RefreshVisible()
    {
        if (surface.IsDisposed) return;
        surface.Invalidate();
        if (!surface.IsHandleCreated || !surface.Visible || session is not TileSession active || surface.Width <= HeaderWidth || surface.Height <= HeaderHeight) return;
        var x = active.Raster.Columns.Position(column);
        var y = active.Raster.Rows.Position(row);
        var firstX = (int)(x / SheetRaster.TileSize);
        var firstY = (int)(y / SheetRaster.TileSize);
        var lastX = Math.Min(active.Raster.TileColumns - 1, (int)((x + SourcePixels(surface.Width - HeaderWidth) - 1) / SheetRaster.TileSize));
        var lastY = Math.Min(active.Raster.TileRows - 1, (int)((y + SourcePixels(surface.Height - HeaderHeight) - 1) / SheetRaster.TileSize));
        var tiles = new List<PreviewTile>();
        for (var tileRow = firstY; tileRow <= lastY; tileRow++)
            for (var tileColumn = firstX; tileColumn <= lastX; tileColumn++)
                tiles.Add(new PreviewTile(tileColumn, tileRow));
        active.Visible = tiles.ToArray();
        var nextTiles = new List<PreviewTile>();
        if (lastY + 1 < active.Raster.TileRows)
            for (var tileColumn = firstX; tileColumn <= lastX; tileColumn++) nextTiles.Add(new PreviewTile(tileColumn, lastY + 1));
        if (lastX + 1 < active.Raster.TileColumns)
            for (var tileRow = firstY; tileRow <= lastY; tileRow++) nextTiles.Add(new PreviewTile(lastX + 1, tileRow));
        active.Prefetch = nextTiles.ToArray();
        StartWorker(active);
    }

    private void StartWorker(TileSession active)
    {
        if (active.Working || active.Failed) return;
        var next = active.Next();
        if (next is PreviewTile tile) _ = Prepare(active, tile);
    }

    private async Task Prepare(TileSession active, PreviewTile tile)
    {
        active.Working = true;
        Bitmap? image = null;
        try
        {
            image = await Task.Run(() => active.Load(tile));
            if (!ReferenceEquals(session, active)) return;
            active.Add(tile, image);
            image = null;
            if (active.Visible.Contains(tile)) surface.Invalidate();
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ExternalException or ArgumentException)
        {
            active.Failed = true;
            if (ReferenceEquals(session, active)) surface.Invalidate();
        }
        finally
        {
            image?.Dispose();
            active.Working = false;
            if (ReferenceEquals(session, active)) StartWorker(active);
            else active.Dispose();
        }
    }

    private void SelectAt(Point point)
    {
        if (session is not TileSession active || source is null || point.X < 0 || point.Y < 0) return;
        var inRowHeader = point.X < HeaderWidth;
        var inColumnHeader = point.Y < HeaderHeight;
        if (inRowHeader && inColumnHeader) return;

        var raster = active.Raster;
        var x = raster.Columns.Position(column) + SourceOffset(point.X - HeaderWidth);
        var y = raster.Rows.Position(row) + SourceOffset(point.Y - HeaderHeight);
        if ((!inRowHeader && x >= raster.Columns.Position(raster.Columns.Limit + 1)) ||
            (!inColumnHeader && y >= raster.Rows.Position(raster.LastRow + 1))) return;
        int? pickedRow = inColumnHeader ? null : raster.Rows.AtPixel(y);
        int? pickedColumn = inRowHeader ? null : raster.Columns.AtPixel(x);
        var mark = marks?.Pick(pickedRow, pickedColumn);

        MergeRange? merge = null;
        Cell? cell = null;
        if (pickedRow is int r && pickedColumn is int c)
        {
            merge = source.Merges.InRows(r, r).FirstOrDefault(range => range.ContainsColumn(c));
            cell = source.At(merge?.FirstRow ?? r, merge?.FirstColumn ?? c);
        }

        highlightedRow = pickedRow ?? 0;
        highlightedColumn = pickedColumn ?? 0;
        highlightedMerge = merge;
        markRow = inRowHeader;
        markColumn = inColumnHeader;
        layoutCount = 1;
        surface.Invalidate();

        var address = merge?.ToString() ?? (pickedColumn is int onlyColumn && pickedRow is null
            ? $"{Address.Column(onlyColumn)} 列" : Address.Format(pickedRow, pickedColumn));
        Picked?.Invoke(new CellPick(address, SheetPreview.Display(cell, date1904), cell?.Formula, mark?.Indices ?? []));
    }

    internal sealed record CellPick(string Address, string Value, string? Formula, IReadOnlyList<int> DifferenceIndices);

    private void Draw(Graphics graphics)
    {
        graphics.Clear(Color.White);
        if (session is not TileSession active)
        {
            TextRenderer.DrawText(graphics, "対応するシートなし", Font, new Point(8, 8), Appearance.Muted);
            return;
        }
        var body = Rectangle.FromLTRB(HeaderWidth, HeaderHeight, Math.Max(HeaderWidth, surface.Width), Math.Max(HeaderHeight, surface.Height));
        var x = active.Raster.Columns.Position(column);
        var y = active.Raster.Rows.Position(row);
        var visibleRows = active.Raster.Rows.Visible(row, active.Raster.LastRow, SourcePixels(body.Height));
        var visibleColumns = active.Raster.Columns.Visible(column, active.Raster.Columns.Limit, SourcePixels(body.Width));
        using var tint = new SolidBrush(Color.Transparent);
        using var outline = new Pen(Color.Transparent);
        var state = graphics.Save();
        graphics.IntersectClip(body);
        if (scale != 1) graphics.InterpolationMode = InterpolationMode.Bilinear;
        foreach (var tile in active.Visible)
        {
            if (!active.Images.TryGetValue(tile, out var image)) continue;
            var tileX = (long)tile.Column * SheetRaster.TileSize - x;
            var tileY = (long)tile.Row * SheetRaster.TileSize - y;
            var left = HeaderWidth + ScreenPixels(tileX);
            var top = HeaderHeight + ScreenPixels(tileY);
            if (scale == 1) graphics.DrawImageUnscaled(image, left, top);
            else
                graphics.DrawImage(image, Rectangle.FromLTRB(left, top,
                    HeaderWidth + ScreenPixels(tileX + SheetRaster.TileSize),
                    HeaderHeight + ScreenPixels(tileY + SheetRaster.TileSize)));
        }
        if (active.Failed)
            TextRenderer.DrawText(graphics, "プレビュー画像を読み込めません", Font, new Point(HeaderWidth + 8, HeaderHeight + 8), Appearance.Muted);
        DrawBodyChanges(graphics, active.Raster, body, visibleRows, visibleColumns, outline);
        DrawSelection(graphics, active.Raster, body);
        graphics.Restore(state);

        using var headers = new SolidBrush(Appearance.Panel);
        graphics.FillRectangle(headers, 0, 0, surface.Width, HeaderHeight);
        graphics.FillRectangle(headers, 0, HeaderHeight, HeaderWidth, surface.Height);
        DrawHeaderChanges(graphics, active.Raster, body, visibleRows, visibleColumns, tint, outline);
        foreach (var number in visibleColumns)
        {
            var left = ColumnOffset(active.Raster, number);
            var right = ColumnOffset(active.Raster, number + 1);
            TextRenderer.DrawText(graphics, Address.Column(number), Font, Rectangle.FromLTRB(left, 0, right, HeaderHeight),
                Appearance.Muted, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }
        foreach (var number in visibleRows)
        {
            var top = RowOffset(active.Raster, number);
            var bottom = RowOffset(active.Raster, number + 1);
            TextRenderer.DrawText(graphics, number.ToString(), Font, Rectangle.FromLTRB(0, top, HeaderWidth, bottom),
                Appearance.Muted, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }
    }

    private void DrawBodyChanges(Graphics graphics, SheetRaster raster, Rectangle body, int[] visibleRows,
        int[] visibleColumns, Pen outline)
    {
        if (marks is null || visibleRows.Length == 0 || visibleColumns.Length == 0) return;
        foreach (var (range, mark) in marks.Rows(visibleRows[0], visibleRows[^1]))
        {
            if (!mark.FillRow) continue;
            var band = Rectangle.FromLTRB(body.Left, RowOffset(raster, range.FirstRow),
                body.Right, RowOffset(raster, range.LastRow + 1));
            DrawMark(graphics, body, band, mark.Color, outline);
        }
        foreach (var (range, mark) in marks.Merges(visibleRows[0], visibleRows[^1]))
        {
            if (range.FirstColumn > visibleColumns[^1] || range.LastColumn < visibleColumns[0]) continue;
            var bounds = CellBounds(raster, range.FirstRow, range.FirstColumn, range.LastRow, range.LastColumn);
            DrawMark(graphics, body, bounds, mark.Color, outline);
        }
        foreach (var r in visibleRows)
        {
            if (marks.CellsInRow(r) is not { } cells) continue;
            foreach (var c in visibleColumns)
                if (cells.TryGetValue(c, out var mark))
                    DrawMark(graphics, body, CellBounds(raster, r, c, r, c), mark.Color, outline);
        }
    }

    private void DrawHeaderChanges(Graphics graphics, SheetRaster raster, Rectangle body, int[] visibleRows,
        int[] visibleColumns, SolidBrush tint, Pen outline)
    {
        if (marks is null) return;
        if (visibleRows.Length > 0)
        {
            var clip = new Rectangle(0, HeaderHeight, HeaderWidth, body.Height);
            foreach (var (range, mark) in marks.Rows(visibleRows[0], visibleRows[^1]))
            {
                var band = Rectangle.FromLTRB(0, RowOffset(raster, range.FirstRow),
                    HeaderWidth, RowOffset(raster, range.LastRow + 1));
                DrawMark(graphics, clip, band, mark.Color, outline, tint);
            }
        }
        if (visibleColumns.Length > 0)
        {
            var clip = new Rectangle(HeaderWidth, 0, body.Width, HeaderHeight);
            foreach (var (range, mark) in marks.Columns(visibleColumns[0], visibleColumns[^1]))
            {
                var band = Rectangle.FromLTRB(ColumnOffset(raster, range.FirstRow),
                    0, ColumnOffset(raster, range.LastRow + 1), HeaderHeight);
                DrawMark(graphics, clip, band, mark.Color, outline, tint);
            }
        }
    }

    private int SourcePixels(int pixels) => (int)Math.Ceiling(pixels / scale);
    private long SourceOffset(int pixel) => (long)Math.Ceiling((pixel + 0.5) / scale) - 1;
    private int ScreenPixels(long pixels) =>
        (int)Math.Clamp(Math.Round(pixels * scale, MidpointRounding.AwayFromZero), -100000, 100000);
    private int ColumnOffset(SheetRaster raster, int number) =>
        HeaderWidth + ScreenPixels(raster.Columns.Position(number) - raster.Columns.Position(column));
    private int RowOffset(SheetRaster raster, int number) =>
        HeaderHeight + ScreenPixels(raster.Rows.Position(number) - raster.Rows.Position(row));

    private Rectangle CellBounds(SheetRaster raster, int firstRow, int firstColumn, int lastRow, int lastColumn) =>
        Rectangle.FromLTRB(ColumnOffset(raster, firstColumn), RowOffset(raster, firstRow),
            ColumnOffset(raster, lastColumn + 1), RowOffset(raster, lastRow + 1));

    private static void DrawMark(Graphics graphics, Rectangle clip, Rectangle bounds, Color color, Pen outline, SolidBrush? tint = null)
    {
        var visible = Rectangle.Intersect(clip, bounds);
        if (visible.Width <= 0 || visible.Height <= 0) return;
        var border = color == Appearance.Added ? Color.FromArgb(33, 122, 76) :
            color == Appearance.Removed ? Color.FromArgb(176, 49, 71) : Color.FromArgb(180, 107, 0);
        if (tint is not null)
        {
            tint.Color = Color.FromArgb(48, border);
            graphics.FillRectangle(tint, visible);
        }
        outline.Color = border;
        graphics.DrawRectangle(outline, bounds.X, bounds.Y, bounds.Width - 1, bounds.Height - 1);
    }

    private void DrawSelection(Graphics graphics, SheetRaster raster, Rectangle body)
    {
        if (highlightedRow <= 0 && highlightedColumn <= 0) return;
        var firstRow = highlightedMerge?.FirstRow ?? Math.Max(1, highlightedRow);
        var lastRow = highlightedMerge?.LastRow ?? firstRow + (markRow ? layoutCount - 1 : 0);
        var firstColumn = highlightedMerge?.FirstColumn ?? Math.Max(1, highlightedColumn);
        var lastColumn = highlightedMerge?.LastColumn ?? firstColumn + (markColumn ? layoutCount - 1 : 0);
        var left = markRow ? body.Left : ColumnOffset(raster, firstColumn);
        var right = markRow ? body.Right : ColumnOffset(raster, lastColumn + 1);
        var top = markColumn ? body.Top : RowOffset(raster, firstRow);
        var bottom = markColumn ? body.Bottom : RowOffset(raster, lastRow + 1);
        if (right <= left || bottom <= top) return;
        using var pen = new Pen(Color.FromArgb(20, 111, 220), 2) { Alignment = PenAlignment.Inset };
        graphics.DrawRectangle(pen, left, top, right - left - 1, bottom - top - 1);
    }

    private void ResetSession()
    {
        ReleaseSession();
        if (source is not null)
            session = new TileSession(new SheetRaster(source, Font, savedZoom, date1904, showGridlines));
    }

    private void ReleaseSession()
    {
        var old = session;
        session = null;
        if (old is not null && !old.Working) old.Dispose();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing) ReleaseSession();
        base.Dispose(disposing);
    }

    private sealed class Surface : Control
    {
        private readonly FullSheetPane owner;

        internal Surface(FullSheetPane owner)
        {
            this.owner = owner;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.Selectable, true);
        }

        protected override void OnPaint(PaintEventArgs e) => owner.Draw(e.Graphics);
        protected override void OnMouseWheel(MouseEventArgs e) => owner.Wheel(e);
        protected override void OnMouseDown(MouseEventArgs e)
        {
            Focus();
            if (e.Button == MouseButtons.Left) owner.SelectAt(e.Location);
            base.OnMouseDown(e);
        }
    }

    private sealed class TileSession : IDisposable
    {
        private const int ResidentLimit = 24;
        private readonly Queue<PreviewTile> eager = [];
        private readonly bool keepOnDisk;
        private DirectoryInfo? directory;
        private readonly HashSet<PreviewTile> prepared = [];

        internal TileSession(SheetRaster raster)
        {
            Raster = raster;
            keepOnDisk = (long)raster.TileRows * raster.UsedTileColumns > ResidentLimit;
            if (!keepOnDisk)
                for (var tileRow = 0; tileRow < raster.TileRows; tileRow++)
                    for (var tileColumn = 0; tileColumn < raster.UsedTileColumns; tileColumn++)
                        eager.Enqueue(new PreviewTile(tileColumn, tileRow));
        }

        internal SheetRaster Raster { get; }
        internal PreviewTile[] Visible { get; set; } = [];
        internal PreviewTile[] Prefetch { get; set; } = [];
        internal Dictionary<PreviewTile, Bitmap> Images { get; } = [];
        internal bool Working { get; set; }
        internal bool Failed { get; set; }

        internal PreviewTile? Next()
        {
            foreach (var tile in Visible)
                if (!Images.ContainsKey(tile)) return tile;
            while (eager.TryDequeue(out var tile))
                if (!prepared.Contains(tile)) return tile;
            foreach (var tile in Prefetch)
                if (!prepared.Contains(tile)) return tile;
            return null;
        }

        internal Bitmap Load(PreviewTile tile)
        {
            if (!keepOnDisk && tile.Column < Raster.UsedTileColumns) return Raster.Render(tile);
            directory ??= Directory.CreateTempSubdirectory("excel-diff-preview-");
            var path = Path.Combine(directory.FullName, $"{tile.Column}-{tile.Row}.png");
            if (File.Exists(path))
            {
                using var saved = new Bitmap(path);
                return new Bitmap(saved);
            }
            var image = Raster.Render(tile);
            try
            {
                image.Save(path, ImageFormat.Png);
                return image;
            }
            catch
            {
                image.Dispose();
                throw;
            }
        }

        internal void Add(PreviewTile tile, Bitmap bitmap)
        {
            Images.Add(tile, bitmap);
            prepared.Add(tile);
            if (Images.Count <= Math.Max(ResidentLimit, Visible.Length + 6)) return;
            foreach (var candidate in Images.Keys)
            {
                if ((!keepOnDisk && candidate.Column < Raster.UsedTileColumns) ||
                    candidate == tile || Visible.Contains(candidate)) continue;
                Images[candidate].Dispose();
                Images.Remove(candidate);
                break;
            }
        }

        public void Dispose()
        {
            foreach (var image in Images.Values) image.Dispose();
            Raster.Dispose();
            if (directory is null) return;
            try { directory.Delete(recursive: true); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }
}
