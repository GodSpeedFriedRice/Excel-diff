using System.IO.Compression;
using System.Xml.Linq;
using ExcelDiff.Core.Comparison;
using ExcelDiff.Core.Workbooks;

internal static class FolderMergeChecks
{
    internal static void Run(string scratch, string template, Action<bool, string> check)
    {
        var baseFolder = Directory.CreateDirectory(Path.Combine(scratch, "folder-base")).FullName;
        var production = Directory.CreateDirectory(Path.Combine(scratch, "folder-production")).FullName;
        var incoming = Directory.CreateDirectory(Path.Combine(scratch, "folder-incoming")).FullName;
        var nested = Path.Combine("reports", "second.xlsx");
        foreach (var folder in new[] { baseFolder, production, incoming })
        {
            Directory.CreateDirectory(Path.Combine(folder, "reports"));
            File.Copy(template, Path.Combine(folder, "first.xlsx"));
            File.Copy(template, Path.Combine(folder, nested));
        }
        foreach (var file in new[] { "first.xlsx", nested })
        {
            Change(Path.Combine(production, file), "1200");
            Change(Path.Combine(incoming, file), "1300");
        }
        File.Copy(template, Path.Combine(production, "keep.xlsx"));
        File.Copy(template, Path.Combine(incoming, "new.xlsx"));
        File.Copy(template, Path.Combine(baseFolder, "broken.xlsx"));
        File.Copy(template, Path.Combine(production, "broken.xlsx"));
        File.WriteAllText(Path.Combine(incoming, "broken.xlsx"), "broken");
        File.WriteAllText(Path.Combine(incoming, "~$locked.xlsx"), "");

        var progress = new List<FolderComparisonProgress>();
        var plan = FolderMerger.Analyze(baseFolder, production, incoming, reportProgress: progress.Add);
        check(plan.Books.Count == 5 && progress.SequenceEqual(Enumerable.Range(0, 6).Select(i => new FolderComparisonProgress(i, 5))),
            "相対パスとサブフォルダを使いファイルを対応付け、進捗を通知する");
        var first = plan.Books.Single(book => book.RelativePath == "first.xlsx");
        var second = plan.Books.Single(book => book.RelativePath == nested);
        var firstCell = first.Plan!.Changes.Single(change => change.Disposition == MergeDisposition.Conflict).Cell;
        var secondCell = second.Plan!.Changes.Single(change => change.Disposition == MergeDisposition.Conflict).Cell;
        check(firstCell == secondCell && firstCell.Address == "D2", "同一セル番地でも別ブックに独立した競合がある");
        check(plan.Books.Count(book => book.Plan is null && book.Notice is not null) == 3, "追加、欠落、解析不能を通知する");

        var choices = new Dictionary<string, IReadOnlyDictionary<MergeCellAddress, MergeChoice>>
        {
            [first.RelativePath] = new Dictionary<MergeCellAddress, MergeChoice> { [firstCell] = MergeChoice.Production }
        };
        var destination = Path.Combine(scratch, "folder-output");
        try
        {
            plan.Save(destination, choices);
            check(false, "別ブックの未解決競合があると保存を拒否する");
        }
        catch (InvalidOperationException) { }
        check(!Directory.Exists(destination), "保存が失敗しても途中までの出力を残さない");

        choices[second.RelativePath] = new Dictionary<MergeCellAddress, MergeChoice> { [secondCell] = MergeChoice.Incoming };
        plan.Save(destination, choices);
        string? Value(string file) => XlsxReader.Read(file).Sheets.Single().At(2, 4)?.Value;
        check(Value(Path.Combine(destination, "first.xlsx")) == "1200" &&
            Value(Path.Combine(destination, nested)) == "1300" &&
            Value(Path.Combine(production, nested)) == "1200", "それぞれの解決結果を別フォルダに保存する");
        check(File.Exists(Path.Combine(destination, "keep.xlsx")) && File.Exists(Path.Combine(destination, "broken.xlsx")) &&
            !File.Exists(Path.Combine(destination, "new.xlsx")), "本番版のファイル構成を維持する");

        var pair = FolderMerger.AnalyzePair(production, incoming);
        check(pair.Books.Count == 5 && pair.Books.All(book => book.BasePath is null) &&
            pair.Books.Single(book => book.RelativePath == "keep.xlsx").Notice?.Contains("共通ベース") == false,
            "2方向のフォルダ比較で共通ベースの欠落を通知しない");
        var pairOutput = Path.Combine(scratch, "pair-folder-output");
        pair.Save(pairOutput, choices);
        check(Value(Path.Combine(pairOutput, "first.xlsx")) == "1200" &&
            Value(Path.Combine(pairOutput, nested)) == "1300", "2方向のフォルダも相対パスを保って保存する");

        try
        {
            plan.Save(Path.Combine(production, "out"), choices);
            check(false, "入力フォルダへの保存を拒否する");
        }
        catch (IOException) { }
        File.Copy(template, Path.Combine(production, "changed-after-analysis.xlsx"));
        var stale = Path.Combine(scratch, "stale-folder");
        try
        {
            plan.Save(stale, choices);
            check(false, "解析後の対象ファイルの増減を検出する");
        }
        catch (IOException) { }
        check(!Directory.Exists(stale), "再解析が必要なときに出力を残さない");
    }

    private static void Change(string path, string value)
    {
        using var zip = ZipFile.Open(path, ZipArchiveMode.Update);
        const string part = "xl/worksheets/sheet1.xml";
        var entry = zip.GetEntry(part)!;
        XDocument document;
        using (var input = entry.Open()) document = XDocument.Load(input);
        var ns = document.Root!.Name.Namespace;
        document.Descendants(ns + "c").Single(cell => (string?)cell.Attribute("r") == "D2").SetElementValue(ns + "v", value);
        entry.Delete();
        using var output = zip.CreateEntry(part).Open();
        document.Save(output);
    }
}
