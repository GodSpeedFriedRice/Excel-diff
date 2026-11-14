using System.Drawing.Drawing2D;
using System.Globalization;
using ExcelDiff.Core.Workbooks;

namespace ExcelDiff.Desktop.Views;

internal static class PreviewText
{
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

    internal static bool FontFlag(IReadOnlyDictionary<string, string> format, string key) =>
        format.GetValueOrDefault($"フォント/{key}") == "1" ||
        format.GetValueOrDefault($"フォント/{key}/val") is "1" or "single" or "double";

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

    internal sealed record PreviewCell(Rectangle Bounds, Rectangle Spill, string Text, Font Font,
        Color Foreground, StringAlignment Alignment, StringAlignment Vertical, bool Wrap, bool Shrink);
}
