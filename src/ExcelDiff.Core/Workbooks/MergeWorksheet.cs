using System.Xml.Linq;

namespace ExcelDiff.Core.Workbooks;

internal readonly record struct MergeContent(string? StorageType, string? RawValue,
    (string Kind, string Text, string? Group) Key, bool Special = false)
{
    internal static readonly MergeContent Blank = new(null, null, ("blank", "", null));

    internal string Display => Key.Kind switch
    {
        "blank" => "（空白）",
        "formula" => "=" + Key.Text,
        "b" => Key.Text == "1" ? "TRUE" : "FALSE",
        "text" when Key.Text.Length == 0 => "\"\"",
        _ => Key.Text
    };

    internal static MergeContent From(XElement node, Cell? parsed)
    {
        var ns = node.Name.Namespace;
        var type = (string?)node.Attribute("t");
        var metadata = node.Attribute("cm") is not null || node.Attribute("vm") is not null;
        if (node.Element(ns + "f") is XElement formula)
        {
            var group = (string?)formula.Attribute("t");
            var special = metadata || (group is not null and not "normal") || formula.Attribute("ref") is not null;
            return new(null, formula.Value,
                ("formula", parsed?.Formula ?? formula.Value, special ? $"{group}:{(string?)formula.Attribute("ref")}" : null),
                special);
        }

        var value = (string?)node.Element(ns + "v");
        if (type is "inlineStr" or "s" or "str")
        {
            if (value is null && node.Element(ns + "is") is null) return Blank with { Special = metadata };
            var text = parsed?.Value ?? "";
            return new("inlineStr", text, ("text", text, null), metadata || parsed?.HasRichText == true);
        }
        if (string.IsNullOrEmpty(value)) return Blank with { Special = metadata };
        var kind = type ?? "n";
        return new(type, value, (kind, parsed?.Value ?? value, null), metadata || kind is not ("n" or "b" or "e" or "d"));
    }
}

internal sealed class MergeWorksheet
{
    internal Dictionary<int, Dictionary<int, MergeContent>> Rows { get; } = [];
    internal MergeIndex FormulaRanges { get; }

    internal MergeWorksheet(Sheet sheet, CancellationToken token)
    {
        var groups = new List<MergeRange>();
        foreach (var (rowNumber, row) in RowNodes(sheet.Xml.Root!))
        {
            token.ThrowIfCancellationRequested();
            var contents = new Dictionary<int, MergeContent>();
            foreach (var (column, node) in CellNodes(row))
            {
                var content = MergeContent.From(node, sheet.At(rowNumber, column));
                if (content.Key != MergeContent.Blank.Key || content.Special) contents.Add(column, content);
                if (node.Element(node.Name.Namespace + "f") is XElement formula && (string?)formula.Attribute("ref") is string range)
                    groups.Add(MergeRange.Parse(range));
            }
            if (contents.Count > 0) Rows.Add(rowNumber, contents);
        }
        FormulaRanges = new(groups);
    }

    internal bool HasMovedLines(MergeWorksheet other, CancellationToken token)
    {
        // Stable, unique data lines that reappear at a different address are evidence of a structural edit.
        return HasMoved(UniqueLines(byColumn: false, token), other.UniqueLines(byColumn: false, token)) ||
            HasMoved(UniqueLines(byColumn: true, token), other.UniqueLines(byColumn: true, token));
    }

    private static bool HasMoved(Dictionary<string, int> before, Dictionary<string, int> after) =>
        before.Any(line => line.Value > 0 && after.TryGetValue(line.Key, out var position) && position > 0 && position != line.Value);

    private Dictionary<string, int> UniqueLines(bool byColumn, CancellationToken token)
    {
        var lines = new Dictionary<int, List<(int Position, MergeContent Content)>>();
        foreach (var (row, cells) in Rows)
        {
            token.ThrowIfCancellationRequested();
            foreach (var (column, content) in cells)
            {
                if (content.Key.Kind is "blank" or "formula") continue;
                var index = byColumn ? column : row;
                if (!lines.TryGetValue(index, out var line)) lines[index] = line = [];
                line.Add((byColumn ? row : column, content));
            }
        }

        var signatures = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var (index, line) in lines)
        {
            token.ThrowIfCancellationRequested();
            var signature = string.Concat(line.OrderBy(item => item.Position).Select(item =>
            {
                var key = item.Content.Key;
                return FormattableString.Invariant($"{item.Position}:{key.Kind}:{key.Text.Length}:{key.Text}");
            }));
            if (!signatures.TryAdd(signature, index)) signatures[signature] = 0;
        }
        return signatures;
    }

    internal static IEnumerable<(int Index, XElement Node)> RowNodes(XElement sheet)
    {
        var ns = sheet.Name.Namespace;
        var index = 0;
        foreach (var row in sheet.Element(ns + "sheetData")?.Elements(ns + "row") ?? [])
        {
            index = (int?)row.Attribute("r") ?? index + 1;
            yield return (index, row);
        }
    }

    internal static IEnumerable<(int Index, XElement Node)> CellNodes(XElement row)
    {
        var index = 0;
        foreach (var cell in row.Elements(row.Name.Namespace + "c"))
        {
            index = (string?)cell.Attribute("r") is string address ? Address.Parse(address).Column : index + 1;
            yield return (index, cell);
        }
    }
}
