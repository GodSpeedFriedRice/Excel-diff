using System.Globalization;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace ExcelDiff.Core.Workbooks;

public sealed record PictureAsset(string WorkbookPath, string PartPath, long WorkbookLength, DateTime LastWriteTimeUtc);
public readonly record struct PictureMarker(int Row, int Column, long OffsetX, long OffsetY);
public sealed record SheetPicture(PictureMarker From, PictureMarker? To, long Width, long Height, PictureAsset? Asset)
{
    public int Order { get; init; }
}

public static partial class XlsxReader
{
    private static (IReadOnlyList<SheetPicture>, IReadOnlyList<SheetShape>) ReadDrawings(ZipArchive zip, XElement sheet, string sheetPath,
        string workbookPath, long workbookLength, DateTime modifiedAt, DrawingTheme theme, CancellationToken token)
    {
        var drawingId = RelationshipId(sheet.Element(sheet.Name.Namespace + "drawing"), "id");
        if (drawingId is null) return ([], []);

        var sheetRelationships = ReadRelationships(zip, sheetPath);
        var drawingPath = ResolvePart(sheetPath, sheetRelationships, drawingId, "drawing");
        if (drawingPath is null || ReadOptionalPart(zip, drawingPath)?.Root is not { } drawing) return ([], []);
        var imageRelationships = ReadRelationships(zip, drawingPath);
        var pictures = new List<SheetPicture>();
        var shapes = new List<SheetShape>();
        var order = 0;
        var ns = drawing.Name.Namespace;
        foreach (var anchor in drawing.Elements())
        {
            token.ThrowIfCancellationRequested();
            order++;
            var child = anchor.Elements().FirstOrDefault(node => node.Name.LocalName is "pic" or "sp" or "cxnSp" or "grpSp" or "graphicFrame");
            if (child is null) continue;
            var properties = child.Elements().FirstOrDefault()?.Element(ns + "cNvPr");
            if (True((string?)properties?.Attribute("hidden"))) continue;
            var from = ReadMarker(anchor.Element(ns + "from"));
            var to = ReadMarker(anchor.Element(ns + "to"));
            var extent = anchor.Element(ns + "ext");
            if (anchor.Name.LocalName == "absoluteAnchor" && anchor.Element(ns + "pos") is { } origin &&
                Number(origin.Attribute("x")?.Value) is long x && Number(origin.Attribute("y")?.Value) is long y)
                from = new PictureMarker(1, 1, x, y);

            if (from is not PictureMarker start) continue;
            var width = Number(extent?.Attribute("cx")?.Value) ?? 0;
            var height = Number(extent?.Attribute("cy")?.Value) ?? 0;
            if (to is null && (width < 0 || height < 0 || (width == 0 && height == 0))) continue;

            if (child.Name != ns + "pic")
            {
                shapes.Add(ReadShape(child, start, to, width, height, order, (string?)properties?.Attribute("name"), theme));
                continue;
            }

            var blip = child.Element(ns + "blipFill")?.Elements().FirstOrDefault(e => e.Name.LocalName == "blip");
            var embedded = RelationshipId(blip, "embed");
            var path = embedded is null ? null : ResolvePart(drawingPath, imageRelationships, embedded, "image");
            var asset = path is null ? null : new PictureAsset(workbookPath, path, workbookLength, modifiedAt);
            pictures.Add(new SheetPicture(start, to, width, height, asset) { Order = order });
        }
        return (pictures, shapes);
    }

    private static IEnumerable<XElement> ReadRelationships(ZipArchive zip, string part)
    {
        return ReadOptionalPart(zip, RelationshipsPath(part))?.Root?.Elements() ?? [];
    }

    private static XDocument? ReadOptionalPart(ZipArchive zip, string path)
    {
        if (zip.GetEntry(path) is not { Length: <= 8 * 1024 * 1024 }) return null;
        try { return Part(zip, path); }
        catch (Exception e) when (e is InvalidDataException or XmlException or IOException) { return null; }
    }

    private static string? ResolvePart(string parent, IEnumerable<XElement> relationships, string id, string type)
    {
        var relation = relationships.FirstOrDefault(e => (string?)e.Attribute("Id") == id && IsType(e, type));
        if (relation is null) return null;
        try { return Resolve(parent, relation); }
        catch (InvalidDataException) { return null; }
    }

    private static string? RelationshipId(XElement? element, string name) =>
        element?.Attributes().FirstOrDefault(a => a.Name.LocalName == name && a.Name.NamespaceName.Length > 0)?.Value;

    private static long? Number(string? value) =>
        long.TryParse(value, CultureInfo.InvariantCulture, out var number) ? number : null;

    private static PictureMarker? ReadMarker(XElement? marker)
    {
        if (marker is null) return null;
        var ns = marker.Name.Namespace;
        var row = Number(marker.Element(ns + "row")?.Value);
        var column = Number(marker.Element(ns + "col")?.Value);
        if (row is not (>= 0 and < 1048576) || column is not (>= 0 and < 16384)) return null;
        return new PictureMarker((int)row.Value + 1, (int)column.Value + 1,
            Number(marker.Element(ns + "colOff")?.Value) ?? 0, Number(marker.Element(ns + "rowOff")?.Value) ?? 0);
    }
}
