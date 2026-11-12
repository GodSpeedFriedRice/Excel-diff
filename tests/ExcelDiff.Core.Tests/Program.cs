using System.IO.Compression;
using ExcelDiff.Core.Comparison;
using ExcelDiff.Core.Workbooks;

var scratch = Directory.CreateTempSubdirectory("excel-diff-");
try
{
    var before = Directory.CreateDirectory(Path.Combine(scratch.FullName, "before", "sub")).FullName;
    var after = Directory.CreateDirectory(Path.Combine(scratch.FullName, "after", "sub")).FullName;
    var a = Path.Combine(before, "sample.xlsx");
    var b = Path.Combine(after, "sample.xlsx");
    MakeBook(a, updated: false);
    MakeBook(b, updated: true);

    var left = XlsxReader.Read(a);
    var right = XlsxReader.Read(b);
    var keyed = WorkbookComparer.Compare(left, right, 1).Single();
    Check(keyed.UsesKey && keyed.KeyWarning is null, "キー列で比較する");
    Check(keyed.Rows.Contains(new AlignedRow(4, 4)), "共通キーの行を対応付ける");
    Check(keyed.Differences.Any(d => d is { Item: "行", Kind: ChangeKind.Removed, LeftRow: 3, RightRow: null }), "削除行を検出する");
    Check(keyed.Differences.Any(d => d is { Item: "行", Kind: ChangeKind.Added, LeftRow: null, RightRow: 3 }), "追加行を検出する");
    Check(keyed.Differences.Any(d => d is { Kind: ChangeKind.Value, Left: "120", Right: "150" }), "値の差分を検出する");
    Check(keyed.Differences.Any(d => d is { Kind: ChangeKind.Formula, Left: "1+1", Right: "2*1" }), "同じ結果の異なる数式を検出する");
    Check(!keyed.Differences.Any(d => d is { Kind: ChangeKind.Value, Column: 3 }), "同じ数式結果を値の差分にしない");
    Check(keyed.Differences.Any(d => d is { Item: "フォントサイズ", Left: "11", Right: "10" }), "フォントサイズを比較する");
    Check(keyed.Differences.Any(d => d is { Item: "折り返して全体を表示", Left: "0", Right: "1" }), "折り返しを比較する");
    Check(keyed.Differences.Any(d => d is { Item: "縮小して全体を表示", Left: "0", Right: "1" }), "縮小表示を比較する");
    Check(keyed.Differences.Any(d => d is { Item: "表示倍率", Left: "100", Right: "125" }), "画面の表示倍率を比較する");
    Check(keyed.Differences.Any(d => d is { Item: "印刷倍率", Left: "95", Right: "80" }), "印刷倍率を比較する");
    Check(keyed.Differences.Any(d => d.Item.StartsWith("結合 ", StringComparison.Ordinal)), "結合範囲を比較する");
    Check(keyed.Differences.Any(d => d is { Item: "行高", Left: "20", Right: "26", LeftRow: 2, RightRow: 2 }), "行高を左右の行に結び付ける");
    Check(WorkbookComparer.Compare(left, right, 6).Single() is { UsesKey: false, KeyWarning: not null }, "空のキー列はセル番地比較に戻す");

    var equivalent = Path.Combine(after, "equivalent.xlsx");
    var equalSource = Path.Combine(before, "equivalent.xlsx");
    MakeBook(equalSource, updated: false);
    MakeBook(equivalent, updated: false, inline: true);
    Check(WorkbookComparer.Compare(XlsxReader.Read(equalSource), XlsxReader.Read(equivalent)).Single().Differences.Count == 0,
        "共有文字列とインライン文字列を同じ内容として扱う");

    File.Copy(a, Path.Combine(before, "left.xlsx"));
    File.Copy(b, Path.Combine(after, "right.xlsx"));
    File.WriteAllText(Path.Combine(before, "broken.xlsx"), "not an xlsx");
    File.Copy(b, Path.Combine(after, "broken.xlsx"));
    var rootLeft = Directory.GetParent(before)!.FullName;
    var rootRight = Directory.GetParent(after)!.FullName;
    var expected = new Dictionary<string, FileStatus>
    {
        ["sample.xlsx"] = FileStatus.Different, ["left.xlsx"] = FileStatus.LeftOnly,
        ["right.xlsx"] = FileStatus.RightOnly, ["broken.xlsx"] = FileStatus.Unreadable,
        ["equivalent.xlsx"] = FileStatus.Same
    };
    var results = FolderComparer.Compare(rootLeft, rootRight, recursive: true);
    Check(results.Count == expected.Count && results.All(r => expected[Path.GetFileName(r.RelativePath)] == r.Status), "フォルダ一覧の全状態を判定する");
    Check(FolderComparer.Compare(rootLeft, rootRight, recursive: false).Count == 0, "サブフォルダの設定を反映する");
    Check(results.Single(r => Path.GetFileName(r.RelativePath) == "equivalent.xlsx").Sheets.Single() is
        { Name: "Sheet1", Status: FileStatus.Same, DifferenceCount: 0 }, "格納形式だけが違うファイルのシートを同一と表示する");
    Check(results.Single(r => Path.GetFileName(r.RelativePath) == "sample.xlsx").Sheets.Single() is
        { Name: "Sheet1", Status: FileStatus.Different, DifferenceCount: > 0 }, "差分シートの件数をフォルダ比較から取得できる");
    Check(results.Single(r => Path.GetFileName(r.RelativePath) == "broken.xlsx").Sheets.Count == 0, "比較できないファイルのシート結果は空にする");

    if (args is [var output])
    {
        Directory.CreateDirectory(Path.Combine(output, "before"));
        Directory.CreateDirectory(Path.Combine(output, "after"));
        foreach (var name in new[] { "sample.xlsx", "equivalent.xlsx", "left.xlsx" })
            File.Copy(Path.Combine(before, name), Path.Combine(output, "before", name), overwrite: true);
        foreach (var name in new[] { "sample.xlsx", "equivalent.xlsx", "right.xlsx" })
            File.Copy(Path.Combine(after, name), Path.Combine(output, "after", name), overwrite: true);
    }
    Console.WriteLine("Excel-diff: 21 checks passed.");
}
finally
{
    scratch.Delete(recursive: true);
}

static void Check(bool result, string name)
{
    if (!result) throw new InvalidOperationException(name);
}

static void MakeBook(string path, bool updated, bool inline = false)
{
    const string ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const string relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
    const string office = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    using var zip = ZipFile.Open(path, ZipArchiveMode.Create);
    void Add(string name, string content)
    {
        using var writer = new StreamWriter(zip.CreateEntry(name).Open());
        writer.Write(content);
    }
    Add("[Content_Types].xml", """
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
          <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
          <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
        </Types>
        """);
    Add("_rels/.rels", $"""<Relationships xmlns="{relationships}"><Relationship Id="r1" Type="{office}/officeDocument" Target="xl/workbook.xml"/></Relationships>""");
    Add("xl/workbook.xml", $"""<workbook xmlns="{ns}" xmlns:r="{office}"><bookViews><workbookView/></bookViews><sheets><sheet name="Sheet1" sheetId="1" r:id="r1"/></sheets></workbook>""");
    Add("xl/_rels/workbook.xml.rels", $"""
        <Relationships xmlns="{relationships}">
          <Relationship Id="r1" Type="{office}/worksheet" Target="worksheets/sheet1.xml"/>
          <Relationship Id="r2" Type="{office}/styles" Target="styles.xml"/>
          <Relationship Id="r3" Type="{office}/sharedStrings" Target="sharedStrings.xml"/>
        </Relationships>
        """);
    Add("xl/sharedStrings.xml", $"""<sst xmlns="{ns}"><si><t>商品コード</t></si></sst>""");
    Add("xl/styles.xml", $"""
        <styleSheet xmlns="{ns}">
          <fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><sz val="10"/><name val="Calibri"/></font></fonts>
          <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
          <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
          <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
          <cellXfs count="3">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
            <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment wrapText="1"/></xf>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment shrinkToFit="1"/></xf>
          </cellXfs>
          <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
        </styleSheet>
        """);
    string title = inline ? """<c r="A1" t="inlineStr"><is><t>商品コード</t></is></c>""" : """<c r="A1" t="s"><v>0</v></c>""";
    Add("xl/worksheets/sheet1.xml", $"""
        <worksheet xmlns="{ns}">
          <sheetViews><sheetView workbookViewId="0" zoomScale="{(updated ? 125 : 100)}"/></sheetViews>
          <sheetData>
            <row r="1">{title}</row>
            <row r="2" ht="{(updated ? 26 : 20)}" customHeight="1">
              <c r="A2" t="inlineStr"><is><t>P-101</t></is></c>
              <c r="B2" s="{(updated ? 1 : 0)}"><v>{(updated ? 150 : 120)}</v></c>
              <c r="C2"><f>{(updated ? "2*1" : "1+1")}</f><v>2</v></c>
              <c r="D2" s="{(updated ? 2 : 0)}"><v>1000</v></c>
            </row>
            <row r="3"><c r="A3" t="inlineStr"><is><t>{(updated ? "P-103" : "P-102")}</t></is></c></row>
            <row r="4"><c r="A4" t="inlineStr"><is><t>P-104</t></is></c></row>
          </sheetData>
          <mergeCells><mergeCell ref="{(updated ? "D1:E1" : "D1:F1")}"/></mergeCells>
          <pageSetup scale="{(updated ? 80 : 95)}"/>
        </worksheet>
        """);
}
