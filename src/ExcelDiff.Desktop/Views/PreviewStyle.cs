using System.Drawing.Drawing2D;
using System.Globalization;
using ExcelDiff.Core.Workbooks;

namespace ExcelDiff.Desktop.Views;

[Flags]
internal enum BorderSides { None = 0, Left = 1, Top = 2, Right = 4, Bottom = 8, Diagonal = 16 }

internal sealed class PreviewStyles
{
    private readonly Dictionary<IReadOnlyDictionary<string, string>, PreviewStyle> styles = new(ReferenceEqualityComparer.Instance);

    internal PreviewStyle Get(IReadOnlyDictionary<string, string> format)
    {
        if (!styles.TryGetValue(format, out var style))
            styles.Add(format, style = new PreviewStyle(format));
        return style;
    }

    internal void Clear() => styles.Clear();
}

internal sealed class PreviewStyle
{
    private readonly Stroke left;
    private readonly Stroke top;
    private readonly Stroke right;
    private readonly Stroke bottom;
    private readonly Stroke diagonal;
    private readonly bool diagonalUp;
    private readonly bool diagonalDown;

    internal PreviewStyle(IReadOnlyDictionary<string, string> format)
    {
        HasFill = format.GetValueOrDefault("背景/patternFill/patternType") == "solid";
        Background = HasFill ? ReadColor(format, "背景/patternFill/fgColor", Color.White) : Color.White;
        var automatic = Background.GetBrightness() < 0.45 ? Color.White : Appearance.Ink;
        Foreground = ReadColor(format, "フォント/color", automatic);
        left = ReadStroke(format, "罫線/left");
        if (left == default) left = ReadStroke(format, "罫線/start");
        top = ReadStroke(format, "罫線/top");
        right = ReadStroke(format, "罫線/right");
        if (right == default) right = ReadStroke(format, "罫線/end");
        bottom = ReadStroke(format, "罫線/bottom");
        diagonal = ReadStroke(format, "罫線/diagonal");
        diagonalUp = format.GetValueOrDefault("罫線/diagonalUp") == "1";
        diagonalDown = format.GetValueOrDefault("罫線/diagonalDown") == "1";
        Sides = (left != default ? BorderSides.Left : 0) | (top != default ? BorderSides.Top : 0) |
            (right != default ? BorderSides.Right : 0) | (bottom != default ? BorderSides.Bottom : 0) |
            (diagonal != default && (diagonalUp || diagonalDown) ? BorderSides.Diagonal : 0);
    }

    internal Color Background { get; }
    internal bool HasFill { get; }
    internal Color Foreground { get; }
    internal BorderSides Sides { get; }

    internal static BorderSides MergeEdges(MergeRange range, int row, int column) =>
        (column == range.FirstColumn ? BorderSides.Left : 0) | (row == range.FirstRow ? BorderSides.Top : 0) |
        (column == range.LastColumn ? BorderSides.Right : 0) | (row == range.LastRow ? BorderSides.Bottom : 0);

    internal void DrawBorders(Graphics graphics, Pen pen, Rectangle bounds, BorderSides sides, PreviewStyle? anchor)
    {
        var x = bounds.Left;
        var y = bounds.Top;
        var lastX = bounds.Right - 1;
        var lastY = bounds.Bottom - 1;
        if ((sides & BorderSides.Left) != 0)
            Draw(WithFallback(left, anchor?.left), x, y, x, lastY, 1, 0);
        if ((sides & BorderSides.Top) != 0)
            Draw(WithFallback(top, anchor?.top), x, y, lastX, y, 0, 1);
        if ((sides & BorderSides.Right) != 0)
            Draw(WithFallback(right, anchor?.right), lastX, y, lastX, lastY, -1, 0);
        if ((sides & BorderSides.Bottom) != 0)
            Draw(WithFallback(bottom, anchor?.bottom), x, lastY, lastX, lastY, 0, -1);
        if ((sides & BorderSides.Diagonal) != 0)
        {
            if (diagonalDown) Draw(diagonal, x, y, lastX, lastY, 0, 0);
            if (diagonalUp) Draw(diagonal, x, lastY, lastX, y, 0, 0);
        }

        void Draw(Stroke stroke, int x1, int y1, int x2, int y2, int dx, int dy)
        {
            if (stroke == default) return;
            pen.Color = stroke.Color;
            pen.Width = 1;
            pen.DashStyle = stroke.Dash;
            var room = dx != 0 ? bounds.Width : bounds.Height;
            for (var offset = 0; offset < Math.Min(stroke.Width, room); offset++)
            {
                if (stroke.Double && offset == 1) continue;
                graphics.DrawLine(pen, x1 + dx * offset, y1 + dy * offset, x2 + dx * offset, y2 + dy * offset);
            }
        }
    }

    private static Stroke WithFallback(Stroke edge, Stroke? fallback) => edge == default ? fallback.GetValueOrDefault() : edge;

    private static Stroke ReadStroke(IReadOnlyDictionary<string, string> format, string key)
    {
        var style = format.GetValueOrDefault(key + "/style");
        if (style is null or "none") return default;
        var dash = style switch
        {
            "dotted" or "hair" => DashStyle.Dot,
            "dashed" or "mediumDashed" => DashStyle.Dash,
            "dashDot" or "mediumDashDot" or "slantDashDot" => DashStyle.DashDot,
            "dashDotDot" or "mediumDashDotDot" => DashStyle.DashDotDot,
            _ => DashStyle.Solid
        };
        var width = style switch
        {
            "thick" or "double" => 3,
            "medium" or "mediumDashed" or "mediumDashDot" or "mediumDashDotDot" => 2,
            _ => 1
        };
        return new Stroke(ReadColor(format, key + "/color", Color.Black), dash, width, style == "double");
    }

    private static Color ReadColor(IReadOnlyDictionary<string, string> format, string key, Color fallback)
    {
        var hex = format.GetValueOrDefault(key + "/rgb");
        if (hex is not { Length: >= 6 } ||
            !int.TryParse(hex.AsSpan(hex.Length - 6), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var rgb))
            return fallback;
        var color = Color.FromArgb(unchecked((int)0xFF000000) | rgb);
        if (!double.TryParse(format.GetValueOrDefault(key + "/tint"), CultureInfo.InvariantCulture, out var tint) ||
            !double.IsFinite(tint) || tint == 0) return color;

        tint = Math.Clamp(tint, -1, 1);
        var lightness = (Math.Max(color.R, Math.Max(color.G, color.B)) + Math.Min(color.R, Math.Min(color.G, color.B))) / 510d;
        var next = tint < 0 ? lightness * (1 + tint) : lightness + (1 - lightness) * tint;
        var range = 1 - Math.Abs(2 * lightness - 1);
        var scale = range == 0 ? 0 : (1 - Math.Abs(2 * next - 1)) / range;
        int Channel(int value) => Math.Clamp((int)Math.Round((next + (value / 255d - lightness) * scale) * 255), 0, 255);
        return Color.FromArgb(Channel(color.R), Channel(color.G), Channel(color.B));
    }

    private readonly record struct Stroke(Color Color, DashStyle Dash, int Width, bool Double);
}

internal readonly record struct PreviewBorder(Rectangle Bounds, PreviewStyle Style, BorderSides Sides, PreviewStyle? Anchor = null)
{
    internal void Draw(Graphics graphics, Pen pen) => Style.DrawBorders(graphics, pen, Bounds, Sides, Anchor);
}
