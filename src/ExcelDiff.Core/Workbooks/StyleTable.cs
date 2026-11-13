using System.Xml.Linq;

namespace ExcelDiff.Core.Workbooks;

internal sealed class StyleTable
{
    private static readonly string[] DefaultIndexedColors =
    [
        "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
        "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
        "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
        "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
        "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
        "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
        "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
        "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333"
    ];
    private readonly IReadOnlyDictionary<string, string>[] formats;

    public StyleTable(XDocument? styles, XDocument? theme)
    {
        var root = styles?.Root;
        if (root is null)
        {
            formats = [new Dictionary<string, string>()];
            return;
        }
        var ns = root.Name.Namespace;
        XElement[] Items(string name) => root.Element(ns + name)?.Elements().ToArray() ?? [];
        var fonts = Items("fonts");
        var fills = Items("fills");
        var borders = Items("borders");
        var bases = Items("cellStyleXfs");
        var numbers = Items("numFmts").ToDictionary(e => (int)e.Attribute("numFmtId")!, e => (string)e.Attribute("formatCode")!);
        var palette = theme?.Descendants().FirstOrDefault(e => e.Name.LocalName == "clrScheme");
        var themeOrder = new[] { "lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5",
            "accent6", "hlink", "folHlink" };
        var indexed = root.Element(ns + "colors")?.Element(ns + "indexedColors")?.Elements().ToArray() ?? [];

        string? Color(XElement e)
        {
            if ((string?)e.Attribute("rgb") is string rgb) return rgb.ToUpperInvariant();
            if ((int?)e.Attribute("theme") is int i && i >= 0 && i < themeOrder.Length)
            {
                var entry = palette?.Elements().FirstOrDefault(c => c.Name.LocalName == themeOrder[i])?.Elements().FirstOrDefault();
                var value = (string?)entry?.Attribute("lastClr") ?? (string?)entry?.Attribute("val");
                if (value is not null) return "FF" + value.ToUpperInvariant();
                if (i is 0 or 1) return i == 0 ? "FFFFFFFF" : "FF000000";
            }
            if ((int?)e.Attribute("indexed") is int j && j >= 0)
            {
                if (j is 64 or 65) return j == 64 ? "FF000000" : "FFFFFFFF";
                if (j < indexed.Length && (string?)indexed[j].Attribute("rgb") is string custom)
                    return "FF" + custom[^Math.Min(6, custom.Length)..].ToUpperInvariant();
                if (j < DefaultIndexedColors.Length) return "FF" + DefaultIndexedColors[j];
            }
            return null;
        }

        void Flatten(IDictionary<string, string> result, string key, XElement? e)
        {
            if (e is null) return;
            var rgb = Color(e);
            if (rgb is not null) result[key + "/rgb"] = rgb;
            foreach (var a in e.Attributes().Where(a => !a.IsNamespaceDeclaration))
            {
                var name = a.Name.LocalName;
                if (rgb is not null && name is "theme" or "indexed" or "rgb") continue;
                result[key + "/" + name] = a.Value is "true" or "false" ? (XlsxReader.True(a.Value) ? "1" : "0") : a.Value;
            }
            if (!e.HasAttributes && !e.HasElements) result[key] = "1";
            foreach (var child in e.Elements().GroupBy(c => c.Name.LocalName))
            {
                var occurrences = child.ToArray();
                for (var i = 0; i < occurrences.Length; i++)
                    Flatten(result, key + "/" + child.Key + (occurrences.Length > 1 ? $"/{i + 1}" : ""), occurrences[i]);
            }
        }

        IReadOnlyDictionary<string, string> Make(XElement xf)
        {
            var result = new Dictionary<string, string>(StringComparer.Ordinal);
            var parentIndex = (int?)xf.Attribute("xfId");
            var parent = parentIndex is int i && i >= 0 && i < bases.Length ? bases[i] : null;
            XElement? FormatSource(string apply) => (string?)xf.Attribute(apply) is "0" or "false" ? parent : xf;
            int Id(string key, string apply) => (int?)FormatSource(apply)?.Attribute(key) ?? (int?)parent?.Attribute(key) ?? 0;
            XElement? Item(XElement[] items, int index) => items.Length == 0 && index == 0 ? null :
                index >= 0 && index < items.Length ? items[index] : throw new InvalidDataException("書式への参照が無効です。");

            Flatten(result, "フォント", Item(fonts, Id("fontId", "applyFont")));
            Flatten(result, "背景", Item(fills, Id("fillId", "applyFill")));
            Flatten(result, "罫線", Item(borders, Id("borderId", "applyBorder")));
            Flatten(result, "配置", FormatSource("applyAlignment")?.Element(ns + "alignment") ?? parent?.Element(ns + "alignment"));
            var number = Id("numFmtId", "applyNumberFormat");
            result["表示形式"] = numbers.GetValueOrDefault(number) ?? (number == 0 ? "General" : $"組み込み {number}");
            foreach (var flag in new[] { "wrapText", "shrinkToFit" })
                result.TryAdd("配置/" + flag, "0");
            return result;
        }

        formats = Items("cellXfs").Select(Make).ToArray();
        if (formats.Length == 0) formats = [new Dictionary<string, string>()];
    }

    public IReadOnlyDictionary<string, string> this[int index] => index >= 0 && index < formats.Length
        ? formats[index] : throw new InvalidDataException("セルの書式が無効です。");
}
