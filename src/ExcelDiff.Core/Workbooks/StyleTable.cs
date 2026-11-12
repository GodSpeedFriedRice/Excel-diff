using System.Xml.Linq;

namespace ExcelDiff.Core.Workbooks;

internal sealed class StyleTable
{
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
            }
            if ((int?)e.Attribute("indexed") is int j && j >= 0 && j < indexed.Length)
                return ((string?)indexed[j].Attribute("rgb"))?.ToUpperInvariant();
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
