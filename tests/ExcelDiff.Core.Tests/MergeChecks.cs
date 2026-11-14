using System.IO.Compression;
using System.Xml.Linq;
using ExcelDiff.Core.Workbooks;

internal static class MergeChecks
{
    internal static void Run(string directory, string template, Action<bool, string> check)
    {
        var original = Path.Combine(directory, "merge-base.xlsx");
        var production = Path.Combine(directory, "merge-production.xlsx");
        var incoming = Path.Combine(directory, "merge-incoming.xlsx");
        var output = Path.Combine(directory, "merge-output.xlsx");
        File.Copy(template, original);
        File.Copy(template, production);
        File.Copy(template, incoming);
        var ns = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main");

        void Change(string book, string address, string fragment)
        {
            EditPart(book, "xl/worksheets/sheet1.xml", root =>
            {
                var (rowIndex, _) = Address.Parse(address);
                var sheet = root.Element(ns + "sheetData")!;
                var row = sheet.Elements(ns + "row").SingleOrDefault(row => (int?)row.Attribute("r") == rowIndex);
                if (row is null)
                {
                    row = new XElement(ns + "row", new XAttribute("r", rowIndex));
                    sheet.Add(row);
                }
                row.Elements(ns + "c").Where(cell => (string?)cell.Attribute("r") == address).Remove();
                if (fragment.Length > 0)
                    row.Add(XElement.Parse($"""<c xmlns="{ns}" r="{address}" {fragment}"""));
            });
        }

        Change(production, "B2", """s="1"><v>125</v></c>""");
        Change(production, "C2", """><f>1+1</f><v>999</v></c>""");
        Change(production, "A3", """t="inlineStr"><is><t>本番の明細</t></is></c>""");
        Change(production, "F4", """><v>7</v></c>""");
        Change(production, "F5", """><v>11</v></c>""");
        Change(original, "A6", """><v>88</v></c>""");
        Change(production, "A6", """s="1"><v>88</v></c>""");
        Change(original, "F5", """><v>10</v></c>""");
        Change(incoming, "B2", """><v>160</v></c>""");
        Change(incoming, "C2", """><f>B2+F4</f><v>500</v></c>""");
        Change(incoming, "F4", """><v>7.0</v></c>""");
        Change(incoming, "F5", """><v>12</v></c>""");
        Change(incoming, "E4", """t="inlineStr"><is><t xml:space="preserve"> 追記 </t></is></c>""");
        Change(incoming, "B7", """t="b"><v>1</v></c>""");
        Change(incoming, "G7", """t="inlineStr"><is><r><rPr><b/></rPr><t>装飾あり</t></r></is></c>""");
        Change(incoming, "H8", """><f t="array" ref="H8:H9">ROW(H8:H9)</f><v>8</v></c>""");
        Change(incoming, "H9", """><v>9</v></c>""");
        Change(original, "I8", """><f>2+2</f><v>4</v></c>""");
        Change(production, "I8", """><f>2+2</f><v>50</v></c>""");
        Change(incoming, "I8", """><f>2+2</f><v>60</v></c>""");
        Change(incoming, "A1", """t="inlineStr"><is><t>更新した見出し</t></is></c>""");
        EditPart(production, "xl/sharedStrings.xml", root => root.SetAttributeValue("count", 1));
        EditPart(production, "xl/_rels/workbook.xml.rels", root =>
            root.Add(new XElement(root.Name.Namespace + "Relationship", new XAttribute("Id", "r4"),
                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"),
                new XAttribute("Target", "calcChain.xml"))));
        EditPart(production, "[Content_Types].xml", root =>
            root.Add(new XElement(root.Name.Namespace + "Override", new XAttribute("PartName", "/xl/calcChain.xml"),
                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"))));
        using (var zip = ZipFile.Open(production, ZipArchiveMode.Update))
        using (var data = zip.CreateEntry("xl/calcChain.xml").Open())
            new XDocument(new XElement(ns + "calcChain", new XElement(ns + "c", new XAttribute("r", "C2"), new XAttribute("i", 1)))).Save(data);

        var plan = WorkbookMerger.Analyze(original, production, incoming);
        var changes = plan.Changes.ToDictionary(change => change.Cell.Address);
        check(plan.Notices.Count == 0, "セル編集はシート構造の変更と区別する");
        check(changes["A3"].Disposition == MergeDisposition.KeepProduction, "本番版だけの変更を保持する");
        check(changes["F4"].Disposition == MergeDisposition.AlreadyMatches, "両側で同じ数値は競合しない");
        check(changes["B2"].Disposition == MergeDisposition.Conflict && changes["F5"].Disposition == MergeDisposition.Conflict,
            "両側の変更が異なる場合は競合させる");
        check(changes["C2"].Disposition == MergeDisposition.TakeIncoming && !changes.ContainsKey("I8"),
            "数式キャッシュを比較から除く");
        check(changes["G7"].Disposition == MergeDisposition.Unsupported && changes["H8"].Disposition == MergeDisposition.Unsupported &&
            changes["H9"].Disposition == MergeDisposition.Unsupported, "セル内装飾と配列数式の範囲を維持する");
        Expect<InvalidOperationException>(() => plan.Save(output, new Dictionary<MergeCellAddress, MergeChoice>()));
        check(!File.Exists(output), "未解決の競合があると保存しない");
        var choices = new Dictionary<MergeCellAddress, MergeChoice>
        {
            [changes["B2"].Cell] = MergeChoice.Production,
            [changes["F5"].Cell] = MergeChoice.Incoming
        };
        var productionBytes = File.ReadAllBytes(production);
        Expect<IOException>(() => plan.Save(production, choices));
        plan.Save(output, choices);
        check(File.ReadAllBytes(production).SequenceEqual(productionBytes), "入力ファイルを上書きしない");
        var sheet = XlsxReader.Read(output).Sheets.Single();
        check(sheet.At(2, 2) is { Value: "125" } formatted && formatted.Format["フォント/sz/val"] == "10" &&
            sheet.At(3, 1)?.Value == "本番の明細" && sheet.At(5, 6)?.Value == "12", "両方の採用先と本番版の書式を反映する");
        check(sheet.At(2, 3) is { Formula: "B2+F4", Value: "" }, "採用した数式を保存し、古い計算値をコピーしない");
        check(sheet.At(6, 1) is { Value: "" } blank && blank.Format["フォント/sz/val"] == "10" &&
            sheet.At(4, 5)?.Value == " 追記 " && sheet.At(7, 2)?.Display == "TRUE" &&
            sheet.At(1, 1)?.Value == "更新した見出し", "セル消去・挿入・文字列・真偽値を書き戻す");
        using (var zip = ZipFile.OpenRead(output))
        {
            using var workbook = zip.GetEntry("xl/workbook.xml")!.Open();
            using var rels = zip.GetEntry("xl/_rels/workbook.xml.rels")!.Open();
            using var types = zip.GetEntry("[Content_Types].xml")!.Open();
            using var shared = zip.GetEntry("xl/sharedStrings.xml")!.Open();
            check((string?)XDocument.Load(workbook).Root!.Element(ns + "calcPr")?.Attribute("fullCalcOnLoad") == "1" &&
                zip.GetEntry("xl/calcChain.xml") is null &&
                !XDocument.Load(rels).Root!.Elements().Any(node => (string?)node.Attribute("Target") == "calcChain.xml") &&
                !XDocument.Load(types).Root!.Elements().Any(node => (string?)node.Attribute("PartName") == "/xl/calcChain.xml") &&
                XDocument.Load(shared).Root!.Attribute("count") is null, "再計算を指定し、古い計算チェーンと共有文字列件数を無効にする");
        }

        Change(incoming, "B2", """><v>170</v></c>""");
        Expect<IOException>(() => plan.Save(Path.Combine(directory, "stale.xlsx"), choices));
        check(!File.Exists(Path.Combine(directory, "stale.xlsx")), "解析後の入力変更を検出する");

        var shifted = Path.Combine(directory, "shifted.xlsx");
        File.Copy(original, shifted);
        Change(shifted, "A3", "");
        Change(shifted, "A4", """t="inlineStr"><is><t>P-102</t></is></c>""");
        Change(shifted, "A5", """t="inlineStr"><is><t>P-104</t></is></c>""");
        var structural = WorkbookMerger.Analyze(original, production, shifted);
        check(structural.Changes.Count == 0 && structural.Notices.Single().Reason.Contains("行・列の移動"),
            "既存行が移動したシートは自動で変更しない");

        var merged = Path.Combine(directory, "changed-merge.xlsx");
        File.Copy(original, merged);
        EditPart(merged, "xl/worksheets/sheet1.xml", root =>
            root.Element(ns + "mergeCells")!.Element(ns + "mergeCell")!.SetAttributeValue("ref", "D1:E1"));
        var geometry = WorkbookMerger.Analyze(original, production, merged);
        check(geometry.Changes.Count == 0 && geometry.Notices.Single().Reason.Contains("結合範囲"), "結合範囲の変更を通知する");
    }

    private static void EditPart(string path, string name, Action<XElement> change)
    {
        using var zip = ZipFile.Open(path, ZipArchiveMode.Update);
        var entry = zip.GetEntry(name)!;
        XDocument document;
        using (var source = entry.Open()) document = XDocument.Load(source);
        change(document.Root!);
        entry.Delete();
        using var output = zip.CreateEntry(name).Open();
        document.Save(output);
    }

    private static void Expect<T>(Action action) where T : Exception
    {
        try { action(); }
        catch (T) { return; }
        throw new InvalidOperationException(typeof(T).Name + " が必要です。");
    }
}
