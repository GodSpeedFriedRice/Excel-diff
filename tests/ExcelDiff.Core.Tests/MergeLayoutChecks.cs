using System.IO.Compression;
using System.Xml.Linq;
using ExcelDiff.Core.Workbooks;

internal static class MergeLayoutChecks
{
    private const string DefaultCells = "<row r='1'><c r='A1'><v>10</v></c><c r='K1'><v>40</v></c></row>";

    internal static void Run(string scratch, string template, Action<bool, string> check)
    {
        var directory = Directory.CreateDirectory(Path.Combine(scratch, "merge-layout")).FullName;
        string Book(string name, string[] ranges, string cells = DefaultCells)
        {
            var path = Path.Combine(directory, name + ".xlsx");
            File.Copy(template, path);
            using var zip = ZipFile.Open(path, ZipArchiveMode.Update);
            var entry = zip.GetEntry("xl/worksheets/sheet1.xml")!;
            XDocument document;
            using (var input = entry.Open()) document = XDocument.Load(input);
            var ns = document.Root!.Name.Namespace;
            document.Root.Element(ns + "sheetData")!.ReplaceWith(
                XElement.Parse($"<sheetData xmlns='{ns}'>{cells}</sheetData>"));
            document.Root.Element(ns + "mergeCells")!.ReplaceWith(new XElement(ns + "mergeCells",
                new XAttribute("count", ranges.Length), ranges.Select(range =>
                    new XElement(ns + "mergeCell", new XAttribute("ref", range)))));
            if (ranges.Length == 0) document.Root.Element(ns + "mergeCells")!.Remove();
            entry.Delete();
            using var output = zip.CreateEntry("xl/worksheets/sheet1.xml").Open();
            document.Save(output);
            return path;
        }
        string Save(WorkbookMergePlan plan, string name, MergeChoice? choice = null)
        {
            var selected = choice is null ? new Dictionary<MergeCellAddress, MergeChoice>() :
                plan.Changes.Where(change => change.CanChoose).ToDictionary(change => change.Cell, _ => choice.Value);
            var path = Path.Combine(directory, name + ".xlsx");
            plan.Save(path, selected);
            return path;
        }
        void Ranges(string path, params string[] expected)
        {
            var sheet = XlsxReader.Read(path).Sheets.Single();
            check(sheet.Merges.Ranges.Select(range => range.ToString()).Order().SequenceEqual(expected.Order()),
                "保存した結合範囲を読み直す " + Path.GetFileName(path));
            using var zip = ZipFile.OpenRead(path);
            var entry = zip.GetEntry("xl/worksheets/sheet1.xml")!;
            using var data = entry.Open();
            var root = XDocument.Load(data).Root!;
            var element = root.Element(root.Name.Namespace + "mergeCells");
            check(expected.Length == 0 ? element is null : (int?)element?.Attribute("count") == expected.Length,
                "結合範囲の件数と空になった要素を正しく保存する");
            var names = root.Elements().Select(node => node.Name.LocalName).ToArray();
            if (expected.Length > 0)
                check(Array.IndexOf(names, "sheetData") < Array.IndexOf(names, "mergeCells") &&
                    Array.IndexOf(names, "mergeCells") < Array.IndexOf(names, "pageSetup"),
                    "結合セルの要素をワークシートの正しい順序に保存する");
        }

        var baseline = Book("base", ["A1:B1"]);
        var production = Book("production", ["A1:B1"], DefaultCells.Replace("40", "41"));
        var incoming = Book("expanded", ["A1:C1"]);
        var productionBytes = File.ReadAllBytes(production);
        var expanded = WorkbookMerger.Analyze(baseline, production, incoming);
        var layout = expanded.Changes.Single(change => change.MergeRanges.Count > 0);
        check(layout.Disposition == MergeDisposition.TakeIncoming && layout.CanChoose && expanded.Notices.Count == 0,
            "片側だけの結合拡張はシート全体を停止せず自動で取り込む");
        var output = Save(expanded, "expanded-result");
        Ranges(output, "A1:C1");
        check(XlsxReader.Read(output).Sheets.Single().At(1, 11)?.Value == "41" &&
            File.ReadAllBytes(production).SequenceEqual(productionBytes), "結合範囲外の編集と入力ファイルを保持する");
        var keep = Path.Combine(directory, "kept-layout.xlsx");
        expanded.Save(keep, new Dictionary<MergeCellAddress, MergeChoice> { [layout.Cell] = MergeChoice.Production });
        Ranges(keep, "A1:B1");

        var merged = Book("fully-merged", ["A1:C1"]);
        var unmerged = Book("unmerged", [],
            "<row r='1'><c r='A1'><v>10</v></c><c r='B1'><v>22</v></c><c r='C1'><v>33</v></c><c r='K1'><v>40</v></c></row>");
        var unmergePlan = WorkbookMerger.Analyze(merged, merged, unmerged);
        var unmergeOutput = Save(unmergePlan, "unmerge-result");
        Ranges(unmergeOutput);
        check(XlsxReader.Read(unmergeOutput).Sheets.Single().At(1, 2)?.Value == "22" &&
            XlsxReader.Read(unmergeOutput).Sheets.Single().At(1, 3)?.Value == "33", "結合解除時に分離したセルの内容を保存する");
        var contraction = WorkbookMerger.Analyze(merged, merged, Book("contracted", ["A1:B1"],
            "<row r='1'><c r='A1'><v>10</v></c><c r='C1'><v>33</v></c><c r='K1'><v>40</v></c></row>"));
        var contracted = Save(contraction, "contraction-result");
        Ranges(contracted, "A1:B1");
        check(XlsxReader.Read(contracted).Sheets.Single().At(1, 3)?.Value == "33", "結合縮小で独立したセルの内容を取り込む");

        var plain = Book("plain", [], "<row r='1'><c r='A1'><v>10</v></c><c r='B1'><v>20</v></c></row>");
        var newMerge = Book("new-merge", ["A1:B1"], "<row r='1'><c r='A1'><v>10</v></c></row>");
        var mergePlan = WorkbookMerger.Analyze(plain, plain, newMerge);
        var mergeOutput = Save(mergePlan, "new-merge-result");
        Ranges(mergeOutput, "A1:B1");
        var removedCell = XlsxReader.Read(mergeOutput).Sheets.Single().At(1, 2);
        check(removedCell?.Value is null or "" && removedCell?.Formula is null,
            "結合して削除されたセルに元の値を残さない");
        var reverted = Save(mergePlan, "new-merge-reverted", MergeChoice.Production);
        Ranges(reverted);
        check(XlsxReader.Read(reverted).Sheets.Single().At(1, 2)?.Value == "20", "採用先を戻すと結合前の内容を維持する");

        var wider = Book("wider", ["A1:D1"]);
        var conflict = WorkbookMerger.Analyze(baseline, incoming, wider);
        check(conflict.Changes.Count == 1 && conflict.Changes[0] is { Disposition: MergeDisposition.Conflict, CanChoose: true } &&
            !conflict.RequiresManualResolution, "重なる結合範囲の変更を選択可能なコンフリクト1件にまとめる");
        var rejected = false;
        try { Save(conflict, "unresolved"); }
        catch (InvalidOperationException) { rejected = true; }
        check(rejected && !File.Exists(Path.Combine(directory, "unresolved.xlsx")), "結合のコンフリクトも未解決なら保存しない");
        Ranges(Save(conflict, "conflict-ours", MergeChoice.Production), "A1:C1");
        Ranges(Save(conflict, "conflict-theirs", MergeChoice.Incoming), "A1:D1");
        var pair = WorkbookMerger.AnalyzePair(incoming, wider);
        check(pair.Changes.Single().Disposition == MergeDisposition.Conflict, "2方向でも結合範囲の採用先を選択する");
        Ranges(Save(pair, "pair-result", MergeChoice.Incoming), "A1:D1");

        var matching = WorkbookMerger.Analyze(baseline, incoming, incoming);
        check(matching.Notices.Count == 0 && matching.Changes.Single().Disposition == MergeDisposition.AlreadyMatches,
            "両側が同じ結合範囲に変更した場合はコンフリクトにしない");
        var matchingWithOtherEdits = WorkbookMerger.Analyze(baseline,
            Book("matching-with-edits", ["A1:C1"], DefaultCells.Replace("40", "41")), incoming);
        check(matchingWithOtherEdits.Notices.Count == 0 && matchingWithOtherEdits.Changes
            .Single(change => change.MergeRanges.Count > 0).Disposition == MergeDisposition.AlreadyMatches,
            "結合範囲が両側で同じなら他のセルが違うブックでも手動確認に回さない");
        Ranges(Save(matchingWithOtherEdits, "matching-with-edits-result"), "A1:C1");
        var edited = Book("edited-anchor", ["A1:B1"], DefaultCells.Replace("10", "11"));
        var contentConflict = WorkbookMerger.Analyze(baseline, edited, incoming);
        check(contentConflict.Changes.Single().Disposition == MergeDisposition.Conflict,
            "片側のセル編集と他方の結合変更が重なる場合は自動で上書きしない");
        var chosen = Save(contentConflict, "content-conflict-ours", MergeChoice.Production);
        Ranges(chosen, "A1:B1");
        check(XlsxReader.Read(chosen).Sheets.Single().At(1, 1)?.Value == "11", "結合範囲とセル内容を同じ版から採用する");
        var sameGeometry = WorkbookMerger.Analyze(baseline, edited,
            Book("anchor-other", ["A1:B1"], DefaultCells.Replace("10", "12")));
        check(sameGeometry.Changes.Single() is { MergeRanges.Count: 0, Disposition: MergeDisposition.Conflict },
            "結合範囲が同じセルの値は通常のセルマージを使う");

        var twoMerges = Book("two-merges", ["A1:B1", "D1:E1"]);
        var separate = WorkbookMerger.Analyze(twoMerges, Book("separate-ours", ["A1:C1", "D1:E1"]),
            Book("separate-theirs", ["A1:B1", "D1:F1"]));
        check(separate.Changes.Count == 2 && separate.Changes.All(change => change.Disposition != MergeDisposition.Conflict),
            "離れた結合範囲の変更は別々に自動採用する");
        Ranges(Save(separate, "separate-result"), "A1:C1", "D1:F1");
        var overlapping = WorkbookMerger.Analyze(twoMerges, Book("overlapping-ours", ["A1:D1"]),
            Book("overlapping-theirs", ["A1:B1", "D1:F1"]));
        check(overlapping.Changes.Count == 1 && overlapping.Changes[0].Disposition == MergeDisposition.Conflict,
            "複数の結合をまたぐ変更をまとめ、重複した結合を生成しない");
        Ranges(Save(overlapping, "overlapping-result", MergeChoice.Incoming), "A1:B1", "D1:F1");

        var corners = WorkbookMerger.Analyze(Book("corners-base", ["C1:D4", "A1:B2"]),
            Book("corners-ours", ["C1:D4", "A1:B2"]), Book("corners-theirs", ["A4:D4", "A1:B3"]));
        check(corners.Changes.Count == 2, "結合範囲群の外接矩形だけが重なる場合も別々に扱う");
        Ranges(Save(corners, "corners-result"), "A4:D4", "A1:B3");

        var formulaBook = Book("formula-merge", ["A1:C1"], "<row r='1'><c r='A1'><f>2+3</f><v>5</v></c></row>");
        var formulaPlan = WorkbookMerger.Analyze(baseline, baseline, formulaBook);
        var formulaOutput = Save(formulaPlan, "formula-result");
        check(XlsxReader.Read(formulaOutput).Sheets.Single().At(1, 1) is { Formula: "2+3", Value: "" },
            "結合セルの数式を取り込み、古い計算結果を除く");
        var arrayBook = Book("array-merge", ["A1:C1"],
            "<row r='1'><c r='A1'><f t='array' ref='A1:C1'>ROW(A1:C1)</f><v>1</v></c></row>");
        var unsupported = WorkbookMerger.Analyze(baseline, baseline, arrayBook);
        check(unsupported.RequiresManualResolution && unsupported.Changes.Single(change => change.MergeRanges.Count > 0)
            .Disposition == MergeDisposition.Unsupported,
            "特殊な数式を含む結合範囲は未対応として保存を停止する");

        var huge = WorkbookMerger.Analyze(newMerge, newMerge,
            Book("huge-merge", ["A1:XFD1048576"], "<row r='1'><c r='A1'><v>10</v></c></row>"));
        Ranges(Save(huge, "huge-result"), "A1:XFD1048576");
    }
}
