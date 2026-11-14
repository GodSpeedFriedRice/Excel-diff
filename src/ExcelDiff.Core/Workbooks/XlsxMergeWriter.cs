using System.IO.Compression;
using System.Xml.Linq;

namespace ExcelDiff.Core.Workbooks;

internal static class XlsxMergeWriter
{
    internal static void Apply(ZipArchive zip, IReadOnlyList<WorkbookMergePlan.Edit> writes, CancellationToken token)
    {
        var replacedSharedString = false;
        foreach (var sheet in writes.GroupBy(edit => edit.Part))
        {
            token.ThrowIfCancellationRequested();
            var part = XlsxReader.Part(zip, sheet.Key);
            var root = part.Root!;
            var ns = root.Name.Namespace;
            var data = root.Element(ns + "sheetData") ?? throw new InvalidDataException("シートのセル領域が見つかりません。");
            var rows = MergeWorksheet.RowNodes(root).ToDictionary(item => item.Index, item => item.Node);
            foreach (var rowChanges in sheet.GroupBy(edit => edit.Change.Cell.Row))
            {
                token.ThrowIfCancellationRequested();
                if (!rows.TryGetValue(rowChanges.Key, out var row))
                {
                    row = new XElement(ns + "row", new XAttribute("r", rowChanges.Key));
                    rows.Add(rowChanges.Key, row);
                }

                var cells = MergeWorksheet.CellNodes(row).ToDictionary(item => item.Index, item => item.Node);
                foreach (var edit in rowChanges)
                {
                    var address = edit.Change.Cell;
                    if (!cells.TryGetValue(address.Column, out var cell))
                    {
                        cell = new XElement(ns + "c", new XAttribute("r", address.Address));
                        cells.Add(address.Column, cell);
                        ExpandDimension(root, address);
                    }
                    replacedSharedString |= (string?)cell.Attribute("t") == "s";
                    SetContent(cell, edit.Incoming);
                }
                if (cells.Values.Any(cell => cell.Parent is null))
                {
                    foreach (var (index, cell) in cells)
                        if (cell.Attribute("r") is null) cell.SetAttributeValue("r", Address.Format(rowChanges.Key, index));
                    row.Elements(ns + "c").Remove();
                    row.AddFirst(cells.OrderBy(item => item.Key).Select(item => item.Value));
                    row.Attribute("spans")?.Remove();
                }
            }
            if (rows.Values.Any(row => row.Parent is null))
            {
                foreach (var (index, row) in rows)
                    if (row.Attribute("r") is null) row.SetAttributeValue("r", index);
                data.Elements(ns + "row").Remove();
                data.Add(rows.OrderBy(item => item.Key).Select(item => item.Value));
            }
            Replace(zip, sheet.Key, part);
        }
        UpdateBook(zip, replacedSharedString);
    }

    private static void SetContent(XElement cell, MergeContent content)
    {
        var ns = cell.Name.Namespace;
        cell.Elements().Where(node => node.Name == ns + "f" || node.Name == ns + "v" || node.Name == ns + "is").Remove();
        cell.SetAttributeValue("t", content.StorageType);
        if (content.Key.Kind == "formula")
            cell.AddFirst(new XElement(ns + "f", content.RawValue));
        else if (content.StorageType == "inlineStr")
            cell.AddFirst(new XElement(ns + "is", new XElement(ns + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), content.RawValue)));
        else if (content.RawValue is not null)
            cell.AddFirst(new XElement(ns + "v", content.RawValue));
    }

    private static void ExpandDimension(XElement sheet, MergeCellAddress address)
    {
        var dimension = sheet.Element(sheet.Name.Namespace + "dimension")?.Attribute("ref");
        if (dimension is null) return;
        var previous = MergeRange.Parse(dimension.Value);
        var first = Address.Format(Math.Min(address.Row, previous.FirstRow), Math.Min(address.Column, previous.FirstColumn));
        var last = Address.Format(Math.Max(address.Row, previous.LastRow), Math.Max(address.Column, previous.LastColumn));
        dimension.Value = first == last ? first : $"{first}:{last}";
    }

    private static void UpdateBook(ZipArchive zip, bool replacedSharedString)
    {
        var workbookPath = XlsxReader.WorkbookPartPath(zip);
        var workbook = XlsxReader.Part(zip, workbookPath);
        var root = workbook.Root!;
        var ns = root.Name.Namespace;
        var calc = root.Element(ns + "calcPr");
        if (calc is null)
        {
            calc = new XElement(ns + "calcPr", new XAttribute("calcId", 0));
            var next = root.Elements().FirstOrDefault(node => node.Name.LocalName is "oleSize" or "customWorkbookViews" or
                "pivotCaches" or "smartTagPr" or "smartTagTypes" or "webPublishing" or "fileRecoveryPr" or "webPublishObjects" or "extLst");
            if (next is null) root.Add(calc);
            else next.AddBeforeSelf(calc);
        }
        calc.SetAttributeValue("fullCalcOnLoad", "1");
        Replace(zip, workbookPath, workbook);

        var relsPath = XlsxReader.RelationshipsPath(workbookPath);
        var rels = XlsxReader.Part(zip, relsPath);
        var relations = rels.Root!.Elements();
        if (replacedSharedString && relations.FirstOrDefault(node => XlsxReader.IsType(node, "sharedStrings")) is XElement shared)
        {
            var sharedPath = XlsxReader.Resolve(workbookPath, shared);
            var strings = XlsxReader.Part(zip, sharedPath);
            if (strings.Root!.Attribute("count") is XAttribute count)
            {
                count.Remove();
                Replace(zip, sharedPath, strings);
            }
        }

        if (relations.FirstOrDefault(node => XlsxReader.IsType(node, "calcChain")) is not XElement chain) return;
        var chainPath = XlsxReader.Resolve(workbookPath, chain);
        zip.GetEntry(chainPath)?.Delete();
        chain.Remove();
        Replace(zip, relsPath, rels);

        const string typesPath = "[Content_Types].xml";
        var types = XlsxReader.Part(zip, typesPath);
        types.Root!.Elements().Where(node => node.Name.LocalName == "Override" &&
            (string?)node.Attribute("ContentType") == "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml")
            .Remove();
        Replace(zip, typesPath, types);
    }

    private static void Replace(ZipArchive zip, string path, XDocument part)
    {
        zip.GetEntry(path)!.Delete();
        using var output = zip.CreateEntry(path, CompressionLevel.Optimal).Open();
        part.Save(output, SaveOptions.DisableFormatting);
    }
}
