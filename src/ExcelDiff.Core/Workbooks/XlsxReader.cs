using System.Globalization;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace ExcelDiff.Core.Workbooks;

public static class XlsxReader
{
    public static Workbook Read(string path, CancellationToken token = default)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read);
        var office = Part(zip, "_rels/.rels").Root!.Elements().FirstOrDefault(e => IsType(e, "officeDocument"))
            ?? throw new InvalidDataException("ブックが見つかりません。");
        var workbookPath = Resolve("", office);
        var book = Part(zip, workbookPath).Root!;
        var ns = book.Name.Namespace;
        var slash = workbookPath.LastIndexOf('/') + 1;
        var relations = Part(zip, workbookPath[..slash] + "_rels/" + workbookPath[slash..] + ".rels")
            .Root!.Elements().ToDictionary(e => Required(e, "Id"));
        XDocument? Related(string type) => relations.Values.FirstOrDefault(e => IsType(e, type)) is XElement relation
            ? Part(zip, Resolve(workbookPath, relation)) : null;

        var shared = Related("sharedStrings")?.Root?.Elements().Select(ReadText).ToArray() ?? [];
        var formats = new StyleTable(Related("styles"), Related("theme"));
        var sheets = new List<Sheet>();
        foreach (var entry in book.Element(ns + "sheets")?.Elements(ns + "sheet") ?? [])
        {
            token.ThrowIfCancellationRequested();
            var id = entry.Attributes().FirstOrDefault(a => a.Name.LocalName == "id")?.Value;
            if (id is null || !relations.TryGetValue(id, out var relation))
                throw new InvalidDataException("シートへの参照が無効です。");
            if (!IsType(relation, "worksheet")) continue;
            var name = Required(entry, "name");

            var area = book.Element(ns + "definedNames")?.Elements(ns + "definedName").FirstOrDefault(e =>
                (string?)e.Attribute("name") == "_xlnm.Print_Area" &&
                (int?)e.Attribute("localSheetId") == entry.ElementsBeforeSelf().Count());
            sheets.Add(ReadSheet(name, (string?)entry.Attribute("state") ?? "visible",
                Part(zip, Resolve(workbookPath, relation)), shared, formats, area?.Value, token));
        }
        return new Workbook(sheets, True((string?)book.Element(ns + "workbookPr")?.Attribute("date1904")));
    }

    private static Sheet ReadSheet(string name, string visibility, XDocument document, string[] shared,
        StyleTable styles, string? area, CancellationToken token)
    {
        var sheet = document.Root!;
        var ns = sheet.Name.Namespace;
        var rows = new SortedDictionary<int, IReadOnlyDictionary<int, Cell>>();
        var settings = new SortedDictionary<string, string>(StringComparer.Ordinal);
        var view = sheet.Element(ns + "sheetViews")?.Element(ns + "sheetView");
        settings["表示倍率"] = (string?)view?.Attribute("zoomScale") ?? "100";
        Add(settings, "ウィンドウ枠", view?.Element(ns + "pane"), "xSplit", "ySplit", "topLeftCell", "state");
        Add(settings, "印刷", sheet.Element(ns + "pageSetup"), "scale", "orientation", "paperSize", "fitToWidth", "fitToHeight");
        Add(settings, "印刷", sheet.Element(ns + "sheetPr")?.Element(ns + "pageSetUpPr"), "fitToPage");
        Add(settings, "印刷余白", sheet.Element(ns + "pageMargins"), "left", "right", "top", "bottom", "header", "footer");
        Add(settings, "既定レイアウト", sheet.Element(ns + "sheetFormatPr"), "defaultRowHeight", "defaultColWidth");
        if (area is not null) settings["印刷範囲"] = area;

        var columnStyles = new Dictionary<int, IReadOnlyDictionary<string, string>>();
        var rowStyles = new Dictionary<int, IReadOnlyDictionary<string, string>>();
        foreach (var column in sheet.Element(ns + "cols")?.Elements(ns + "col") ?? [])
        {
            var first = (int?)column.Attribute("min") ?? 0;
            var last = (int?)column.Attribute("max") ?? 0;
            if (first < 1 || last < first || last > 16384) throw new InvalidDataException("列の範囲が無効です。");
            for (var index = first; index <= last; index++)
            {
                var key = $"列 {Address.Column(index)}";
                Add(settings, key, column, "width", "hidden");
                if ((int?)column.Attribute("style") is int style)
                {
                    columnStyles[index] = styles[style];
                    AddFormat(settings, key, styles[style], styles[0]);
                }
            }
        }
        var merges = new MergeIndex((sheet.Element(ns + "mergeCells")?.Elements(ns + "mergeCell") ?? [])
            .Select(merge => MergeRange.Parse(Required(merge, "ref"))));

        var rowElements = sheet.Element(ns + "sheetData")?.Elements(ns + "row") ?? [];
        var sharedFormulas = rowElements.SelectMany(r => r.Elements(ns + "c"))
            .Where(c => c.Element(ns + "f") is XElement f && (string?)f.Attribute("t") == "shared" && f.Value.Length != 0)
            .ToDictionary(c => Required(c.Element(ns + "f")!, "si"), c => c);
        var rowIndex = 0;
        var lastUsedRow = 0;
        var lastUsedColumn = 0;
        foreach (var row in rowElements)
        {
            token.ThrowIfCancellationRequested();
            rowIndex = (int?)row.Attribute("r") ?? rowIndex + 1;
            if (rowIndex is < 1 or > 1048576) throw new InvalidDataException("行番号が無効です。");
            Add(settings, $"行 {rowIndex}", row, "ht", "hidden");
            IReadOnlyDictionary<string, string>? rowFormat = null;
            if ((int?)row.Attribute("s") is int style)
            {
                rowFormat = styles[style];
                rowStyles[rowIndex] = rowFormat;
                AddFormat(settings, $"行 {rowIndex}", rowFormat, styles[0]);
            }
            var cells = new SortedDictionary<int, Cell>();
            var column = 0;
            foreach (var cell in row.Elements(ns + "c"))
            {
                column = (string?)cell.Attribute("r") is string address ? Address.Parse(address).Column : column + 1;
                if (column > 16384) throw new InvalidDataException("列番号が無効です。");
                var kind = (string?)cell.Attribute("t") ?? "n";
                var value = (string?)cell.Element(ns + "v") ?? "";
                if (kind == "s")
                {
                    if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var index) || index >= shared.Length)
                        throw new InvalidDataException("共有文字列への参照が無効です。");
                    value = shared[index];
                }
                else if (kind == "inlineStr" && cell.Element(ns + "is") is XElement inline)
                    value = ReadText(inline);
                if (kind is "s" or "inlineStr" or "str") kind = "text";
                if (kind == "n" && double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
                    value = number == 0 ? "0" : number.ToString("G", CultureInfo.InvariantCulture);

                var node = cell.Element(ns + "f");
                var formula = node?.Value;
                if (node is not null && (string?)node.Attribute("t") == "shared" && formula?.Length == 0)
                {
                    if (!sharedFormulas.TryGetValue(Required(node, "si"), out var source))
                        throw new InvalidDataException("共有数式の参照先が見つかりません。");
                    var origin = Address.Parse(Required(source, "r"));
                    formula = $"{source.Element(ns + "f")!.Value} [共有数式: 行 {rowIndex - origin.Row}, 列 {column - origin.Column}]";
                }
                var inherited = rowFormat ?? columnStyles.GetValueOrDefault(column) ?? styles[0];
                var format = (int?)cell.Attribute("s") is int cellStyle ? styles[cellStyle] : inherited;
                if (value.Length != 0 || formula is not null || !ReferenceEquals(format, inherited))
                {
                    cells.Add(column, new Cell(kind, value, formula, format));
                    lastUsedColumn = Math.Max(lastUsedColumn, column);
                }
            }
            if (cells.Count > 0)
            {
                rows.Add(rowIndex, cells);
                lastUsedRow = Math.Max(lastUsedRow, rowIndex);
            }
        }
        return new Sheet(name, visibility, rows, settings, styles[0], merges, document)
        {
            RowFormats = rowStyles,
            ColumnFormats = columnStyles,
            UsedCellExtent = (lastUsedRow, lastUsedColumn)
        };
    }

    private static XDocument Part(ZipArchive zip, string path)
    {
        var entry = zip.GetEntry(path) ?? throw new InvalidDataException($"必要なデータがありません: {path}");
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            MaxCharactersInDocument = 100_000_000
        });
        return XDocument.Load(reader);
    }

    private static string Resolve(string parent, XElement relation)
    {
        var baseUri = new Uri(new Uri("https://xlsx.invalid/"), parent);
        if ((string?)relation.Attribute("TargetMode") == "External" ||
            !Uri.TryCreate(baseUri, Required(relation, "Target"), out var uri) ||
            uri.Host != "xlsx.invalid" || uri.Scheme != "https" || uri.Query.Length != 0 || uri.Fragment.Length != 0)
            throw new InvalidDataException("外部または無効な参照です。");
        return Uri.UnescapeDataString(uri.AbsolutePath.TrimStart('/'));
    }

    private static string Required(XElement element, string name) => (string?)element.Attribute(name)
        ?? throw new InvalidDataException($"必要な属性がありません: {name}");

    private static bool IsType(XElement relation, string type) =>
        ((string?)relation.Attribute("Type"))?.EndsWith("/" + type, StringComparison.Ordinal) == true;

    private static string ReadText(XElement element)
    {
        var ns = element.Name.Namespace;
        return string.Concat(element.Elements().Where(e => e.Name == ns + "t" || e.Name == ns + "r")
            .Select(e => e.Name == ns + "t" ? e.Value : (string?)e.Element(ns + "t")));
    }

    internal static bool True(string? value) => value is "1" or "true";

    private static void Add(IDictionary<string, string> target, string group, XElement? element, params string[] names)
    {
        foreach (var name in names)
            if ((string?)element?.Attribute(name) is string value)
                target[$"{group} / {name}"] = value is "true" or "false" ? (True(value) ? "1" : "0") : value;
    }

    private static void AddFormat(IDictionary<string, string> target, string group,
        IReadOnlyDictionary<string, string> format, IReadOnlyDictionary<string, string> defaults)
    {
        foreach (var (key, value) in format)
            if (defaults.GetValueOrDefault(key) != value) target[$"{group} / {key}"] = value;
    }
}
