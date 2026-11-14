using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using ExcelDiff.Core.Comparison;
using ExcelDiff.Core.Workbooks;

internal static class PackageChecks
{
    internal static void Run(string directory, string template, Action<bool, string> check)
    {
        var original = Path.Combine(directory, "parts-base.xlsx");
        var production = Path.Combine(directory, "parts-production.xlsx");
        var incoming = Path.Combine(directory, "parts-incoming.xlsx");
        var recompressed = Path.Combine(directory, "parts-repacked.xlsx");
        File.Copy(template, original);
        const string image = "xl/media/image.png";
        const string originalOnly = "xl/embeddings/old.bin";
        const string newPart = "xl/custom-data/record.custom";
        const string deletedFromBoth = "xl/embeddings/removed.bin";
        const string sameAddition = "xl/embeddings/shared.bin";
        const string theme = "xl/theme/theme1.xml";
        const string relationships = "xl/worksheets/_rels/sheet1.xml.rels";
        const string sheet = "xl/worksheets/sheet1.xml";

        Set(original, image, [0, 1, 2]);
        Set(original, originalOnly, [7]);
        Set(original, deletedFromBoth, [8]);
        Set(original, relationships, Encoding.UTF8.GetBytes("<rels><item>old</item></rels>"));
        File.Copy(original, production);
        File.Copy(original, incoming);
        Set(production, image, [0, 1, 3]);
        Set(incoming, image, [0, 2, 2]);
        Set(incoming, originalOnly, null);
        Set(production, deletedFromBoth, null);
        Set(incoming, deletedFromBoth, null);
        Set(production, sameAddition, [11, 12]);
        Set(incoming, sameAddition, [11, 12]);
        Set(production, theme, Encoding.UTF8.GetBytes("<theme />"));
        Set(incoming, newPart, Encoding.UTF8.GetBytes("<content>新規</content>"));
        Set(incoming, relationships, Encoding.UTF8.GetBytes("<rels><item>new</item></rels>"));

        var contentTypes = Read(incoming, "[Content_Types].xml");
        contentTypes.Root!.Add(new XElement(contentTypes.Root.Name.Namespace + "Override",
            new XAttribute("PartName", "/" + newPart), new XAttribute("ContentType", "application/example+xml")));
        Set(incoming, "[Content_Types].xml", Encoding.UTF8.GetBytes(contentTypes.ToString()));

        var semantic = WorkbookComparer.Compare(XlsxReader.Read(original), XlsxReader.Read(incoming)).Single();
        check(semantic.Differences.Count == 0, "画像や埋め込み部品だけの変更では従来のシート差分が空になる");
        var pair = PackageComparer.Compare(original, incoming).ToDictionary(part => part.Path);
        check(pair.Count == 7 && !pair.ContainsKey(sheet) &&
            pair[image] is { Kind: PackagePartChange.Modified, Before.IsXml: false, After.IsXml: false } &&
            pair[image].Before!.Length == pair[image].After!.Length,
            "同じサイズでもバイナリの内容変更を検出する");
        check(pair[originalOnly] is { Kind: PackagePartChange.Removed, After: null } &&
            pair[newPart] is { Kind: PackagePartChange.Added, After.IsXml: true } &&
            pair[relationships] is { Kind: PackagePartChange.Modified, After.IsXml: true }, "追加・削除・参照と拡張子以外のXML分類");
        var oldXml = PackageComparer.ReadXml(original, pair[relationships].Before);
        var newXml = PackageComparer.ReadXml(incoming, pair[relationships].After);
        check(XmlLineDiff.Compare(oldXml, newXml).Lines.Any(line => line.Kind == XmlLineKind.Added && line.Text.Contains("new")),
            "選択したXML部品の行差分を取得する");
        check(PackageComparer.ReadXml(original, pair[newPart].Before) is null &&
            PackageComparer.ReadXml(incoming, pair[newPart].After)!.Root!.Value == "新規", "片側にだけあるXMLを表示する");

        using (var source = ZipFile.OpenRead(original))
        using (var output = ZipFile.Open(recompressed, ZipArchiveMode.Create))
            foreach (var item in source.Entries.Reverse())
            {
                using var from = item.Open();
                using var into = output.CreateEntry(item.FullName, CompressionLevel.NoCompression).Open();
                from.CopyTo(into);
            }
        check(PackageComparer.Compare(original, recompressed).Count == 0, "ZIPの圧縮方式とエントリ順の違いを無視する");

        var plan = WorkbookMerger.Analyze(original, production, incoming);
        var triple = plan.PartChanges.ToDictionary(part => part.Path);
        check(plan.Changes.Count == 0 && triple[image].Status == PackagePartMergeStatus.BothChanged &&
            triple[theme].Status == PackagePartMergeStatus.ProductionOnly &&
            triple[relationships].Status == PackagePartMergeStatus.IncomingOnly, "共通ベースから両方の変更方向を判定する");
        check(triple[sameAddition].Status == PackagePartMergeStatus.SameChange &&
            triple[deletedFromBoth] is { Status: PackagePartMergeStatus.SameChange, ProductionChange: PackagePartChange.Removed },
            "両側で同じ部品を追加または削除した場合を判定する");

        var workbook = Read(incoming, "xl/workbook.xml");
        workbook.Root!.AddFirst(new XElement(workbook.Root.Name.Namespace + "workbookPr", new XAttribute("date1904", "1")));
        Set(incoming, "xl/workbook.xml", Encoding.UTF8.GetBytes(workbook.ToString()));
        var incompatible = WorkbookMerger.Analyze(original, production, incoming);
        check(incompatible.Changes.Count == 0 && incompatible.Notices.Single().Reason.Contains("日付システム") &&
            incompatible.PartChanges.Any(part => part.Path == "xl/workbook.xml"), "日付システムが違っても比較結果を表示する");

        Set(incoming, relationships, Encoding.UTF8.GetBytes("<rels><item>changed again</item></rels>"));
        try
        {
            PackageComparer.ReadXml(incoming, pair[relationships].After);
            check(false, "比較後の部品変更を検出する");
        }
        catch (InvalidDataException) { }
    }

    private static XDocument Read(string book, string part)
    {
        using var zip = ZipFile.OpenRead(book);
        using var input = zip.GetEntry(part)!.Open();
        return XDocument.Load(input);
    }

    private static void Set(string book, string part, byte[]? content)
    {
        using var zip = ZipFile.Open(book, ZipArchiveMode.Update);
        zip.GetEntry(part)?.Delete();
        if (content is null) return;
        using var output = zip.CreateEntry(part).Open();
        output.Write(content);
    }
}
