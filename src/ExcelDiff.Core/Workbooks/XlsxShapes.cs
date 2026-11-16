using System.Globalization;
using System.Xml.Linq;

namespace ExcelDiff.Core.Workbooks;

public sealed record SheetShape(PictureMarker From, PictureMarker? To, long Width, long Height, int Order,
    string Geometry, string? Name, uint? Fill, ShapeOutline Outline, ShapeLabel Label, float Rotation, bool FlipX, bool FlipY);

public sealed record ShapeOutline(uint? Color, long Width, string? Dash, bool StartArrow, bool EndArrow);
public sealed record ShapeLabel(string Text, uint Color, float Points, string? Font, bool Bold, bool Italic,
    string Horizontal, string Vertical);

public static partial class XlsxReader
{
    private static SheetShape ReadShape(XElement shape, PictureMarker from, PictureMarker? to, long width, long height,
        int order, string? name, DrawingTheme theme)
    {
        var ns = shape.Name.Namespace;
        var properties = shape.Element(ns + "spPr");
        var style = shape.Element(ns + "style");
        var transform = DrawingTheme.Child(properties, "xfrm");
        var preset = (string?)DrawingTheme.Child(properties, "prstGeom")?.Attribute("prst");
        var geometry = preset ?? (shape.Name.LocalName == "cxnSp" ? "line" :
            shape.Name.LocalName == "sp" && DrawingTheme.Child(properties, "custGeom") is null &&
            True((string?)shape.Element(ns + "nvSpPr")?.Element(ns + "cNvSpPr")?.Attribute("txBox")) ? "rect" : "unsupported");
        var fill = theme.Paint(properties, DrawingTheme.Child(style, "fillRef"), line: false);
        var line = DrawingTheme.Child(properties, "ln");
        var lineRef = DrawingTheme.Child(style, "lnRef");
        var lineStyle = theme.LineStyle(lineRef);
        var stroke = theme.Paint(line, lineRef, line: true);
        var outline = new ShapeOutline(stroke, Number((string?)line?.Attribute("w")) ??
            Number((string?)lineStyle?.Attribute("w")) ?? 9525,
            (string?)(DrawingTheme.Child(line, "prstDash") ?? DrawingTheme.Child(lineStyle, "prstDash"))?.Attribute("val"),
            HasArrow(line, "headEnd"), HasArrow(line, "tailEnd"));
        var body = shape.Element(ns + "txBody");
        var paragraphs = body?.Elements().Where(element => element.Name.LocalName == "p").ToArray() ?? [];
        var text = string.Join("\n", paragraphs.Select(paragraph => string.Concat(paragraph.Elements().Select(run =>
            run.Name.LocalName == "br" ? "\n" : (string?)DrawingTheme.Child(run, "t")))));
        var paragraphFormat = DrawingTheme.Child(paragraphs.FirstOrDefault(), "pPr");
        var runFormat = paragraphs.SelectMany(paragraph => paragraph.Elements()).Select(run => DrawingTheme.Child(run, "rPr"))
            .FirstOrDefault(format => format is not null);
        var defaults = DrawingTheme.Child(paragraphFormat, "defRPr");
        var fontRef = DrawingTheme.Child(style, "fontRef");
        var bodyFormat = DrawingTheme.Child(body, "bodyPr");
        var label = new ShapeLabel(text.Length <= 4096 ? text : text[..4096],
            theme.Paint(runFormat, null, line: false) ?? theme.Paint(defaults, null, line: false) ?? theme.Color(fontRef) ?? 0xFF000000,
            (float)Math.Clamp((Number((string?)runFormat?.Attribute("sz")) ?? Number((string?)defaults?.Attribute("sz")) ?? 1100) / 100d, 2, 120),
            (string?)(DrawingTheme.Child(runFormat, "latin") ?? DrawingTheme.Child(defaults, "latin"))?.Attribute("typeface"),
            True((string?)runFormat?.Attribute("b") ?? (string?)defaults?.Attribute("b")),
            True((string?)runFormat?.Attribute("i") ?? (string?)defaults?.Attribute("i")),
            (string?)paragraphFormat?.Attribute("algn") ?? "l", (string?)bodyFormat?.Attribute("anchor") ?? "t");
        return new(from, to, width, height, order, geometry, name, fill, outline, label,
            (Number((string?)transform?.Attribute("rot")) ?? 0) % 21600000 / 60000f,
            True((string?)transform?.Attribute("flipH")), True((string?)transform?.Attribute("flipV")));

        bool HasArrow(XElement? format, string tag) => (string?)(DrawingTheme.Child(format, tag) ??
            DrawingTheme.Child(lineStyle, tag))?.Attribute("type") is not (null or "none");
    }
}

internal sealed class DrawingTheme(XDocument? theme)
{
    private readonly XElement? palette = theme?.Descendants().FirstOrDefault(node => node.Name.LocalName == "clrScheme");
    private readonly XElement? formats = theme?.Descendants().FirstOrDefault(node => node.Name.LocalName == "fmtScheme");

    internal static XElement? Child(XElement? parent, string name) =>
        parent?.Elements().FirstOrDefault(node => node.Name.LocalName == name);

    internal XElement? LineStyle(XElement? reference) => Style(reference, "lnStyleLst");

    private XElement? Style(XElement? reference, string list)
    {
        if (!int.TryParse((string?)reference?.Attribute("idx"), NumberStyles.None, CultureInfo.InvariantCulture, out var index) || index <= 0)
            return null;
        if (list == "fillStyleLst" && index > 1000)
        {
            list = "bgFillStyleLst";
            index -= 1000;
        }
        return Child(formats, list)?.Elements().ElementAtOrDefault(index - 1);
    }

    internal uint? Paint(XElement? properties, XElement? reference, bool line)
    {
        var fill = Fill(properties);
        if (fill?.Name.LocalName == "noFill") return null;
        var referenceColor = Color(reference);
        if (fill is not null) return Color(fill, referenceColor);
        var style = Style(reference, line ? "lnStyleLst" : "fillStyleLst");
        if (line) style = Fill(style);
        return style?.Name.LocalName == "noFill" || (string?)reference?.Attribute("idx") is "0" or "1000" ? null :
            Color(style, referenceColor) ?? referenceColor;
    }

    private static XElement? Fill(XElement? properties) => properties?.Elements().FirstOrDefault(node =>
        node.Name.LocalName is "noFill" or "solidFill" or "gradFill" or "pattFill");

    internal uint? Color(XElement? container, uint? placeholder = null)
    {
        var node = container?.DescendantsAndSelf().FirstOrDefault(item =>
            item.Name.LocalName is "srgbClr" or "schemeClr" or "sysClr" or "prstClr");
        if (node is null) return null;
        var value = (string?)node.Attribute("val");
        if (node.Name.LocalName == "schemeClr")
        {
            if (value == "phClr") return placeholder is uint color ? Modify(color, node) : null;
            var key = value switch { "bg1" => "lt1", "tx1" => "dk1", "bg2" => "lt2", "tx2" => "dk2", _ => value };
            var entry = key is null ? null : Child(palette, key)?.Elements().FirstOrDefault();
            value = (string?)entry?.Attribute("lastClr") ?? (string?)entry?.Attribute("val") ?? key switch
            {
                "lt1" => "FFFFFF", "dk1" => "000000", "lt2" => "E7E6E6", "dk2" => "44546A",
                "accent1" => "4472C4", "accent2" => "ED7D31", "accent3" => "A5A5A5", "accent4" => "FFC000",
                "accent5" => "5B9BD5", "accent6" => "70AD47", "hlink" => "0563C1", "folHlink" => "954F72", _ => null
            };
        }
        else if (node.Name.LocalName == "sysClr") value = (string?)node.Attribute("lastClr") ?? value;
        else if (node.Name.LocalName == "prstClr")
            value = value switch { "black" => "000000", "white" => "FFFFFF", "red" => "FF0000", "blue" => "0000FF",
                "green" => "008000", "yellow" => "FFFF00", "gray" => "808080", _ => null };

        return value is { Length: 6 } && uint.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var rgb)
            ? Modify(0xFF000000 | rgb, node) : null;
    }

    private static uint Modify(uint color, XElement node)
    {
        double Ratio(string tag, double fallback) =>
            double.TryParse((string?)Child(node, tag)?.Attribute("val"), CultureInfo.InvariantCulture, out var value) &&
            double.IsFinite(value) ? Math.Clamp(value / 100000, 0, 1) : fallback;
        var shade = Ratio("shade", 1) * Ratio("lumMod", 1);
        var tint = Ratio("tint", 1);
        var lift = Ratio("lumOff", 0) * 255;
        uint Channel(int shift)
        {
            var source = (color >> shift) & 255;
            return (uint)Math.Clamp(Math.Round(source * shade * tint + (1 - tint) * 255 + lift), 0, 255);
        }
        var alpha = (uint)Math.Round((color >> 24) * Ratio("alpha", 1));
        return (alpha << 24) | (Channel(16) << 16) | (Channel(8) << 8) | Channel(0);
    }
}
