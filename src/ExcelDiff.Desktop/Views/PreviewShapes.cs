using System.Drawing.Drawing2D;
using ExcelDiff.Core.Workbooks;

namespace ExcelDiff.Desktop.Views;

internal static class PreviewShapes
{
    internal static void Draw(Graphics graphics, RectangleF bounds, SheetShape shape, double zoom)
    {
        var state = graphics.Save();
        try
        {
            graphics.SmoothingMode = SmoothingMode.AntiAlias;
            graphics.TranslateTransform(bounds.X + bounds.Width / 2, bounds.Y + bounds.Height / 2);
            if (shape.Rotation != 0) graphics.RotateTransform(shape.Rotation);
            var body = new RectangleF(-bounds.Width / 2, -bounds.Height / 2, bounds.Width, bounds.Height);
            using var geometry = Geometry(body, shape.Geometry, shape.FlipX, shape.FlipY);
            var supported = geometry is not null;
            using var placeholder = supported ? null : Geometry(body, "rect", false, false);
            var path = geometry ?? placeholder!;
            if (supported && shape.Fill is uint fill)
            {
                using var brush = new SolidBrush(ToColor(fill));
                graphics.FillPath(brush, path);
            }
            var outline = shape.Outline;
            if (outline.Color is not null || !supported)
            {
                using var pen = new Pen(outline.Color is uint stroke && supported ? ToColor(stroke) : Appearance.Muted,
                    (float)Math.Clamp(outline.Width / 9525d * zoom, .6, 64));
                pen.DashStyle = supported ? Dash(outline.Dash) : DashStyle.Dash;
                if (outline.StartArrow) pen.StartCap = LineCap.ArrowAnchor;
                if (outline.EndArrow) pen.EndCap = LineCap.ArrowAnchor;
                graphics.DrawPath(pen, path);
            }
            var text = shape.Label.Text.Length > 0 ? shape.Label.Text : supported ? "" : shape.Name ?? "図形";
            if (text.Length > 0) DrawLabel(graphics, body, shape.Label, text, zoom);
        }
        finally { graphics.Restore(state); }
    }

    private static DashStyle Dash(string? dash) => dash switch
    {
        "dot" or "sysDot" => DashStyle.Dot,
        "dash" or "lgDash" or "sysDash" => DashStyle.Dash,
        "dashDot" or "lgDashDot" or "sysDashDot" => DashStyle.DashDot,
        "lgDashDotDot" or "sysDashDotDot" => DashStyle.DashDotDot,
        _ => DashStyle.Solid
    };

    private static GraphicsPath? Geometry(RectangleF box, string type, bool flipX, bool flipY)
    {
        var path = new GraphicsPath();
        PointF Point(float x, float y) => new(box.Left + box.Width * (flipX ? 1 - x : x),
            box.Top + box.Height * (flipY ? 1 - y : y));
        PointF[] Arrow(string direction)
        {
            (float X, float Y)[] points = [(0, .26f), (.53f, .26f), (.53f, 0), (1, .5f), (.53f, 1), (.53f, .74f), (0, .74f)];
            return points.Select(p => direction switch
            {
                "leftArrow" => Point(1 - p.X, p.Y),
                "upArrow" => Point(p.Y, 1 - p.X),
                "downArrow" => Point(p.Y, p.X),
                _ => Point(p.X, p.Y)
            }).ToArray();
        }
        switch (type)
        {
            case "rect":
                path.AddRectangle(box);
                break;
            case "roundRect":
                var diameter = Math.Min(box.Width, box.Height) * .32f;
                path.AddArc(box.Left, box.Top, diameter, diameter, 180, 90);
                path.AddArc(box.Right - diameter, box.Top, diameter, diameter, 270, 90);
                path.AddArc(box.Right - diameter, box.Bottom - diameter, diameter, diameter, 0, 90);
                path.AddArc(box.Left, box.Bottom - diameter, diameter, diameter, 90, 90);
                path.CloseFigure();
                break;
            case "ellipse":
                path.AddEllipse(box);
                break;
            case "triangle":
                path.AddPolygon([Point(.5f, 0), Point(1, 1), Point(0, 1)]);
                break;
            case "rtTriangle":
                path.AddPolygon([Point(0, 0), Point(1, 1), Point(0, 1)]);
                break;
            case "diamond":
                path.AddPolygon([Point(.5f, 0), Point(1, .5f), Point(.5f, 1), Point(0, .5f)]);
                break;
            case "rightArrow" or "leftArrow" or "upArrow" or "downArrow":
                path.AddPolygon(Arrow(type));
                break;
            case "line" or "straightConnector1":
                path.AddLine(Point(0, 0), Point(1, 1));
                break;
            case "lineInv":
                path.AddLine(Point(0, 1), Point(1, 0));
                break;
            case "bentConnector2":
                path.AddLines([Point(0, 0), Point(1, 0), Point(1, 1)]);
                break;
            case "bentConnector3":
                path.AddLines([Point(0, 0), Point(.5f, 0), Point(.5f, 1), Point(1, 1)]);
                break;
            default:
                path.Dispose();
                return null;
        }
        return path;
    }

    private static void DrawLabel(Graphics graphics, RectangleF body, ShapeLabel label, string text, double zoom)
    {
        var inset = (float)(3 * zoom);
        body.Inflate(-inset, -inset);
        if (body.Width <= 1 || body.Height <= 1) return;
        var style = (label.Bold ? FontStyle.Bold : FontStyle.Regular) | (label.Italic ? FontStyle.Italic : FontStyle.Regular);
        var points = (float)Math.Clamp(label.Points * zoom, 2, 120);
        using var font = CreateFont(label.Font, points, style);
        using var brush = new SolidBrush(ToColor(label.Color));
        using var format = new StringFormat
        {
            Alignment = label.Horizontal switch { "ctr" => StringAlignment.Center, "r" => StringAlignment.Far, _ => StringAlignment.Near },
            LineAlignment = label.Vertical switch { "ctr" => StringAlignment.Center, "b" => StringAlignment.Far, _ => StringAlignment.Near },
            Trimming = StringTrimming.EllipsisCharacter
        };
        graphics.DrawString(text, font, brush, body, format);
    }

    private static Font CreateFont(string? name, float size, FontStyle style)
    {
        var fallback = SystemFonts.MessageBoxFont!;
        try { return new Font(name is { Length: > 0 } && !name.StartsWith('+') ? name : fallback.Name, size, style); }
        catch (ArgumentException) { return new Font(fallback.FontFamily, size); }
    }

    private static Color ToColor(uint value) => Color.FromArgb(unchecked((int)value));
}
