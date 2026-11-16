using System.IO.Compression;
using System.Xml.Linq;
using ExcelDiff.Core.Workbooks;

internal static class DrawingChecks
{
    internal static void Run(string scratch, string template, Action<bool, string> check)
    {
        var path = Path.Combine(scratch, "drawing.xlsx");
        File.Copy(template, path);
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            const string sheetPath = "xl/worksheets/sheet1.xml";
            var sheetEntry = zip.GetEntry(sheetPath)!;
            XDocument sheet;
            using (var input = sheetEntry.Open()) sheet = XDocument.Load(input);
            sheetEntry.Delete();
            XNamespace rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            sheet.Root!.Add(new XElement(sheet.Root.Name.Namespace + "drawing", new XAttribute(rel + "id", "drawing")));
            Add(sheetPath, sheet.ToString());
            Add("xl/worksheets/_rels/sheet1.xml.rels", $"""
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="drawing" Type="{rel}/drawing" Target="../drawings/drawing1.xml"/>
                </Relationships>
                """);
            Add("xl/drawings/drawing1.xml", """
                <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <xdr:oneCellAnchor>
                    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>9525</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
                    <xdr:ext cx="190500" cy="95250"/>
                    <xdr:sp>
                      <xdr:nvSpPr><xdr:cNvPr id="1" name="受付"/><xdr:cNvSpPr/></xdr:nvSpPr>
                      <xdr:spPr>
                        <a:xfrm rot="5400000" flipH="1"/>
                        <a:prstGeom prst="roundRect"/>
                        <a:solidFill><a:srgbClr val="4472C4"><a:alpha val="50000"/></a:srgbClr></a:solidFill>
                        <a:ln><a:noFill/></a:ln>
                      </xdr:spPr>
                      <xdr:txBody><a:bodyPr anchor="ctr"/><a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="1200" b="1"/><a:t>受付</a:t></a:r></a:p></xdr:txBody>
                    </xdr:sp>
                    <xdr:clientData/>
                  </xdr:oneCellAnchor>
                  <xdr:absoluteAnchor>
                    <xdr:pos x="190500" y="190500"/>
                    <xdr:ext cx="285750" cy="0"/>
                    <xdr:cxnSp>
                      <xdr:nvCxnSpPr><xdr:cNvPr id="2" name="矢印"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr>
                      <xdr:spPr><a:prstGeom prst="straightConnector1"/>
                        <a:ln w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:tailEnd type="triangle"/></a:ln>
                      </xdr:spPr>
                    </xdr:cxnSp>
                    <xdr:clientData/>
                  </xdr:absoluteAnchor>
                </xdr:wsDr>
                """);

            void Add(string name, string text)
            {
                using var writer = new StreamWriter(zip.CreateEntry(name).Open());
                writer.Write(text);
            }
        }
        var shapes = XlsxReader.Read(path).Sheets.Single().Shapes;
        check(shapes.Count == 2 &&
            shapes[0] is { Geometry: "roundRect", Order: 1, From: { Row: 3, Column: 2, OffsetX: 9525 },
                Fill: 0x804472C4, Outline.Color: null, Label: { Text: "受付", Bold: true, Points: 12,
                    Horizontal: "ctr", Vertical: "ctr" }, Rotation: 90, FlipX: true },
            "図形の位置・形状・文字・色・回転を読み込む");
        check(shapes[1] is { Geometry: "straightConnector1", Order: 2, Width: 285750, Height: 0,
            Outline: { Color: 0xFFFF0000, Width: 19050, EndArrow: true } }, "高さゼロの直線と矢印を読み込む");
    }
}
