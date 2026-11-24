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
            var contents = new Dictionary<(int Row, int Column), MergeContent>();
            foreach (var edit in sheet)
            {
                token.ThrowIfCancellationRequested();
                if (edit.Layout is not { } layout)
                {
                    contents[(edit.Change.Cell.Row, edit.Change.Cell.Column)] = edit.Incoming;
                    continue;
                }
                foreach (var (rowNumber, row) in rows)
                {
                    token.ThrowIfCancellationRequested();
                    var ranges = layout.Scope.InRows(rowNumber, rowNumber);
                    if (ranges.Count == 0) continue;
                    foreach (var (column, _) in MergeWorksheet.CellNodes(row))
                        if (ranges.Any(range => range.ContainsColumn(column)))
                            contents[(rowNumber, column)] = MergeContent.Blank;
                }
                foreach (var cell in layout.IncomingCells) contents[cell.Key] = cell.Value;
                SetMerges(root, layout);
            }
            foreach (var rowChanges in contents.GroupBy(cell => cell.Key.Row))
            {
                token.ThrowIfCancellationRequested();
                if (!rows.TryGetValue(rowChanges.Key, out var row))
                {
                    row = new XElement(ns + "row", new XAttribute("r", rowChanges.Key));
                    rows.Add(rowChanges.Key, row);
                }

                var cells = MergeWorksheet.CellNodes(row).ToDictionary(item => item.Index, item => item.Node);
                foreach (var (address, content) in rowChanges)
                {
                    if (!cells.TryGetValue(address.Column, out var cell))
                    {
                        cell = new XElement(ns + "c", new XAttribute("r", Address.Format(address.Row, address.Column)));
                        cells.Add(address.Column, cell);
                        ExpandDimension(root, new("", address.Row, address.Column));
                    }
                    replacedSharedString |= (string?)cell.Attribute("t") == "s";
                    SetContent(cell, content);
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

    private static void SetMerges(XElement root, WorkbookMergePlan.LayoutEdit layout)
    {
        var ns = root.Name.Namespace;
        var merges = root.Element(ns + "mergeCells");
        if (merges is null)
        {
            if (layout.IncomingMerges.Count == 0) return;
            merges = new XElement(ns + "mergeCells");
            var previous = root.Elements().Last(element => element.Name.Namespace == ns && element.Name.LocalName is
                "sheetPr" or "dimension" or "sheetViews" or "sheetFormatPr" or "cols" or "sheetData" or
                "sheetCalcPr" or "sheetProtection" or "protectedRanges" or "scenarios" or "autoFilter" or
                "sortState" or "dataConsolidate" or "customSheetViews");
            previous.AddAfterSelf(merges);
        }
        merges.Elements(ns + "mergeCell").Where(element =>
        {
            var range = MergeRange.Parse((string)element.Attribute("ref")!);
            return layout.Scope.InRows(range.FirstRow, range.LastRow).Any(other =>
                range.FirstColumn <= other.LastColumn && other.FirstColumn <= range.LastColumn);
        }).Remove();
        foreach (var range in layout.IncomingMerges)
        {
            merges.Add(new XElement(ns + "mergeCell", new XAttribute("ref", range.ToString())));
            ExpandDimension(root, new("", range.LastRow, range.LastColumn));
        }
        var count = merges.Elements(ns + "mergeCell").Count();
        if (count == 0) merges.Remove();
        else merges.SetAttributeValue("count", count);
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
