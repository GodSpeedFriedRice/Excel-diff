using System.IO.Compression;
using System.Xml.Linq;
using ExcelDiff.Core.Workbooks;

internal static class PrintChecks
{
    internal static void Run(string scratch, string template, Action<bool, string> check)
    {
        var path = Path.Combine(scratch, "print-layout.xlsx");
        File.Copy(template, path);
        const string name = "西,帳票'1";
        const string bookPart = "xl/workbook.xml";
        const string printPart = "xl/worksheets/sheet1.xml";
        Edit(path, bookPart, root =>
        {
            var ns = root.Name.Namespace;
            var sheets = root.Element(ns + "sheets")!;
            sheets.Element(ns + "sheet")!.SetAttributeValue("name", name);
            sheets.AddAfterSelf(new XElement(ns + "definedNames",
                new XElement(ns + "definedName", new XAttribute("name", "_xlnm.Print_Area"), new XAttribute("localSheetId", 0),
                    "'西,帳票''1'!$B$3:$I$43,'西,帳票''1'!$M$51:$T$60"),
                new XElement(ns + "definedName", new XAttribute("name", "_xlnm.Print_Titles"), new XAttribute("localSheetId", 0),
                    "'西,帳票''1'!$A:$A,'西,帳票''1'!$1:$2")));
        });
        Edit(path, printPart, root =>
        {
            var ns = root.Name.Namespace;
            root.Element(ns + "sheetViews")!.Element(ns + "sheetView")!.SetAttributeValue("view", "pageBreakPreview");
            root.Element(ns + "pageSetup")!.SetAttributeValue("pageOrder", "overThenDown");
            root.Add(new XElement(ns + "rowBreaks", new XElement(ns + "brk", new XAttribute("id", 22), new XAttribute("man", "1"))));
            root.Add(new XElement(ns + "colBreaks", new XElement(ns + "brk", new XAttribute("id", 6))));
        });
        var sheet = XlsxReader.Read(path).Sheets.Single();
        var print = sheet.PrintSettings;
        check(sheet.Name == name && print.Areas.SequenceEqual([new(3, 2, 43, 9), new(51, 13, 60, 20)]) &&
            !print.HasUnresolvedArea, "引用符とカンマを含む名前の複数印刷範囲を読む");
        check(print.Titles.SequenceEqual([new(null, 1, null, 1), new(1, null, 2, null)]) &&
            print.RowBreaks.SequenceEqual([new(23, true)]) && print.ColumnBreaks.SequenceEqual([new(7, false)]),
            "印刷タイトルと手動・自動の改ページの番地を読む");
        check(sheet.Settings["表示モード"] == "pageBreakPreview" && sheet.Settings["印刷 / pageOrder"] == "overThenDown",
            "表示モードとページ順を比較項目として保持する");
        Edit(path, bookPart, root =>
            root.Element(root.Name.Namespace + "definedNames")!.Elements().First().Value = "OFFSET('西,帳票''1'!$A$1,0,0,10,5)");
        var dynamicRange = XlsxReader.Read(path).Sheets.Single().PrintSettings;
        check(dynamicRange.HasUnresolvedArea && dynamicRange.Areas.Count == 0, "動的な式の引数を印刷範囲と誤認しない");
    }

    private static void Edit(string book, string name, Action<XElement> edit)
    {
        using var zip = ZipFile.Open(book, ZipArchiveMode.Update);
        var entry = zip.GetEntry(name)!;
        XDocument document;
        using (var input = entry.Open()) document = XDocument.Load(input);
        edit(document.Root!);
        entry.Delete();
        using var output = zip.CreateEntry(name).Open();
        document.Save(output);
    }
}
