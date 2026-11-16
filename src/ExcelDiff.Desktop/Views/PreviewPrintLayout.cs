using System.Drawing.Drawing2D;
using System.Globalization;
using ExcelDiff.Core.Workbooks;

namespace ExcelDiff.Desktop.Views;

internal sealed class PreviewPrintLayout
{
    private static readonly Color BoundaryColor = Color.FromArgb(26, 90, 195);
    private readonly PrintArea[] areas;
    private readonly bool acrossFirst;
    private readonly bool unresolved;

    internal PreviewPrintLayout(Sheet sheet, PreviewAxis rows, PreviewAxis columns, int usedRow, int usedColumn)
    {
        var print = sheet.PrintSettings;
        unresolved = print.HasUnresolvedArea;
        var ranges = print.Areas.Select(range => new MergeRange(range.FirstRow ?? 1, range.FirstColumn ?? 1,
            range.LastRow ?? Math.Max(1, usedRow), range.LastColumn ?? Math.Max(1, usedColumn))).Distinct().ToArray();
        if (ranges.Length == 0 && usedRow > 0 && usedColumn > 0)
            ranges = [new MergeRange(1, 1, usedRow, usedColumn)];

        var settings = sheet.Settings;
        var (width, height) = PaperSize(settings);
        if (Get("orientation") == "landscape" && width < height) (width, height) = (height, width);
        var headings = Get("headings") == "1";
        var pageWidth = Math.Max(1, (width - Margin("left", .7) - Margin("right", .7)) * 96 - (headings ? 36 : 0));
        var pageHeight = Math.Max(1, (height - Margin("top", .75) - Margin("bottom", .75)) * 96 - (headings ? 20 : 0));
        var rowTitles = print.Titles.FirstOrDefault(range => range.FirstColumn is null);
        var columnTitles = print.Titles.FirstOrDefault(range => range.FirstRow is null);
        var fit = Get("fitToPage") == "1";
        acrossFirst = Get("pageOrder") == "overThenDown";
        var numbering = Get("useFirstPageNumber") == "1" ? ReadInt("firstPageNumber", 1) : 1;
        double scale = Math.Clamp(ReadInt("scale", 100), 10, 400) / 100d;

        if (fit)
        {
            var fitWidth = ReadInt("fitToWidth", 1);
            var fitHeight = ReadInt("fitToHeight", 1);
            double lower = .1, upper = 1;
            if (Fits(1)) lower = 1;
            else
                for (var iteration = 0; iteration < 14; iteration++)
                {
                    var middle = (lower + upper) / 2;
                    if (Fits(middle)) lower = middle;
                    else upper = middle;
                }
            scale = lower;

            bool Fits(double candidate) => ranges.All(range =>
                FitsAxis(columns, range.FirstColumn, range.LastColumn, columnTitles.FirstColumn, columnTitles.LastColumn,
                    pageWidth / candidate, fitWidth) &&
                FitsAxis(rows, range.FirstRow, range.LastRow, rowTitles.FirstRow, rowTitles.LastRow,
                    pageHeight / candidate, fitHeight));
        }

        var results = new List<PrintArea>(ranges.Length);
        long page = numbering;
        foreach (var range in ranges)
        {
            var rowPages = BuildAxis(rows, range.FirstRow, range.LastRow, rowTitles.FirstRow, rowTitles.LastRow,
                pageHeight / scale, print.RowBreaks, fit);
            var columnPages = BuildAxis(columns, range.FirstColumn, range.LastColumn, columnTitles.FirstColumn, columnTitles.LastColumn,
                pageWidth / scale, print.ColumnBreaks, fit);
            results.Add(new(range, rowPages, columnPages, page));
            page += (long)rowPages.Count * columnPages.Count;
        }
        areas = results.ToArray();
        PageCount = page - numbering;
        LastRow = areas.Select(area => area.Cells.LastRow).DefaultIfEmpty().Max();
        LastColumn = areas.Select(area => area.Cells.LastColumn).DefaultIfEmpty().Max();

        string? Get(string key) => settings.GetValueOrDefault("印刷 / " + key);
        int ReadInt(string key, int fallback) =>
            int.TryParse(Get(key), NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) ? value : fallback;
        double Margin(string key, double fallback) => TryPositive(settings.GetValueOrDefault("印刷余白 / " + key), out var value) ? value : fallback;
    }

    internal long PageCount { get; }
    internal int LastRow { get; }
    internal int LastColumn { get; }
    internal string Summary => unresolved ? $"印刷範囲を推定　{PageCount:N0} ページ（目安）" :
        PageCount == 0 ? "印刷する範囲がありません" : $"印刷範囲　{PageCount:N0} ページ（目安）";

    private static (double Width, double Height) PaperSize(IReadOnlyDictionary<string, string> settings)
    {
        string? Get(string key) => settings.GetValueOrDefault("印刷 / " + key);
        if (TryPositive(Get("paperWidth"), out var width) && width > 0 && TryPositive(Get("paperHeight"), out var height) && height > 0)
        {
            var perInch = Get("paperUnits") switch { "in" => 1, "cm" => 2.54, "pt" => 72, "pc" or "pi" => 6, _ => 25.4 };
            return (width / perInch, height / perInch);
        }

        return Get("paperSize") switch
        {
            "1" or "2" => (8.5, 11),
            "3" or "4" => (11, 17),
            "5" => (8.5, 14),
            "7" => (7.25, 10.5),
            "8" => (297 / 25.4, 420 / 25.4),
            "11" => (148 / 25.4, 210 / 25.4),
            "12" => (250 / 25.4, 354 / 25.4),
            "13" => (182 / 25.4, 257 / 25.4),
            "66" => (420 / 25.4, 594 / 25.4),
            _ => (210 / 25.4, 297 / 25.4)
        };
    }

    private static bool TryPositive(string? input, out double value) =>
        double.TryParse(input, NumberStyles.Float, CultureInfo.InvariantCulture, out value) && double.IsFinite(value) && value >= 0;

    private static bool FitsAxis(PreviewAxis axis, int first, int last, int? titleStart, int? titleEnd,
        double pixels, int pages)
    {
        if (pages <= 0) return true;
        var current = first;
        for (var page = 0; page < pages && current <= last; page++)
            current = Next(axis, current, last + 1, pixels - TitleCost(axis, current, titleStart, titleEnd));
        return current > last;
    }

    private static PageAxis BuildAxis(PreviewAxis axis, int first, int last, int? titleStart, int? titleEnd,
        double pixels, IReadOnlyList<PrintBreak> saved, bool fit)
    {
        var breaks = saved.Where(point => point.Before > first && point.Before <= last && (!fit || !point.Manual)).ToArray();
        var manual = breaks.Where(point => point.Manual).Select(point => point.Before).ToHashSet();
        if (breaks.Any(point => !point.Manual))
            return new([first, .. breaks.Select(point => point.Before), last + 1], manual);

        var edges = new List<int> { first };
        foreach (var next in breaks.Select(point => point.Before).Append(last + 1))
        {
            var current = edges[^1];
            while (current < next)
            {
                var room = pixels - TitleCost(axis, current, titleStart, titleEnd);
                current = Next(axis, current, next, room);
                edges.Add(current);
            }
        }
        return new(edges.ToArray(), manual);
    }

    private static int Next(PreviewAxis axis, int start, int limit, double pixels)
    {
        var end = axis.Position(start) + (long)Math.Clamp(Math.Floor(pixels), 1, int.MaxValue);
        return axis.Position(limit) <= end ? limit : Math.Clamp(axis.AtPixel(end), start + 1, limit);
    }

    private static long TitleCost(PreviewAxis axis, int start, int? first, int? last) =>
        first is int low && last is int high && low < start ? axis.Position(Math.Min(start, high + 1)) - axis.Position(low) : 0;

    internal void Draw(Graphics graphics, Rectangle body, Font font, Func<int, int> columnX, Func<int, int> rowY,
        int firstRow, int lastRow, int firstColumn, int lastColumn)
    {
        Rectangle Bounds(MergeRange range) => Rectangle.FromLTRB(columnX(range.FirstColumn), rowY(range.FirstRow),
            columnX(range.LastColumn + 1), rowY(range.LastRow + 1));
        using var outside = new Region(body);
        foreach (var area in areas)
            if (Bounds(area.Cells).IntersectsWith(body)) outside.Exclude(Bounds(area.Cells));
        using var shade = new SolidBrush(Color.FromArgb(162, 156, 156, 156));
        graphics.FillRegion(shade, outside);
        using var outer = new Pen(BoundaryColor, 2);
        using var inner = new Pen(BoundaryColor) { DashStyle = DashStyle.Dash };
        using var manual = new Pen(BoundaryColor);
        using var pageFont = new Font(font.FontFamily, Math.Max(11, font.Size * 1.65f), FontStyle.Bold);
        using var pageInk = new SolidBrush(Color.FromArgb(132, 75, 90, 113));
        using var centered = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
        foreach (var area in areas)
        {
            var bounds = Bounds(area.Cells);
            if (bounds.Width < 1 || bounds.Height < 1 || !bounds.IntersectsWith(body)) continue;
            graphics.DrawRectangle(outer, bounds.Left, bounds.Top, bounds.Width - 1, bounds.Height - 1);
            var (r0, r1) = area.Rows.Visible(firstRow, lastRow);
            var (c0, c1) = area.Columns.Visible(firstColumn, lastColumn);
            var previous = int.MinValue;
            for (var index = Math.Max(1, r0); index <= r1; index++)
            {
                var position = rowY(area.Rows.Edges[index]);
                if (position == previous) continue;
                previous = position;
                graphics.DrawLine(area.Rows.IsManual(index) ? manual : inner, bounds.Left, position, bounds.Right - 1, position);
            }
            previous = int.MinValue;
            for (var index = Math.Max(1, c0); index <= c1; index++)
            {
                var position = columnX(area.Columns.Edges[index]);
                if (position == previous) continue;
                previous = position;
                graphics.DrawLine(area.Columns.IsManual(index) ? manual : inner, position, bounds.Top, position, bounds.Bottom - 1);
            }
            var numberedColumns = Enumerable.Range(c0, c1 - c0 + 1).Where(index =>
                columnX(area.Columns.Edges[index + 1]) - columnX(area.Columns.Edges[index]) >= 65).ToArray();
            if (numberedColumns.Length == 0) continue;
            for (var r = r0; r <= r1; r++)
            {
                var top = rowY(area.Rows.Edges[r]);
                var bottom = rowY(area.Rows.Edges[r + 1]);
                if (bottom - top < pageFont.Height + 6) continue;
                foreach (var c in numberedColumns)
                {
                    var page = area.FirstNumber + (acrossFirst ? (long)r * area.Columns.Count + c : (long)c * area.Rows.Count + r);
                    var rect = Rectangle.FromLTRB(columnX(area.Columns.Edges[c]), top, columnX(area.Columns.Edges[c + 1]), bottom);
                    var visible = Rectangle.Intersect(rect, body);
                    if (visible.Width >= 65 && visible.Height >= pageFont.Height + 6)
                        graphics.DrawString($"{page} ページ", pageFont, pageInk, visible, centered);
                }
            }
        }
    }

    private sealed record PrintArea(MergeRange Cells, PageAxis Rows, PageAxis Columns, long FirstNumber);

    private sealed record PageAxis(int[] Edges, HashSet<int> Manual)
    {
        internal int Count => Edges.Length - 1;
        internal bool IsManual(int index) => Manual.Contains(Edges[index]);
        internal (int First, int Last) Visible(int first, int last) => (PageAt(first), PageAt(last));

        private int PageAt(int position)
        {
            var index = Array.BinarySearch(Edges, position);
            return Math.Clamp(index >= 0 ? index : ~index - 1, 0, Count - 1);
        }
    }
}
