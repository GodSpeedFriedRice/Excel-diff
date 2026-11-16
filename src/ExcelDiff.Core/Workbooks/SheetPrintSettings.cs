using System.Globalization;
using System.Xml.Linq;

namespace ExcelDiff.Core.Workbooks;

public readonly record struct PrintReference(int? FirstRow, int? FirstColumn, int? LastRow, int? LastColumn);
public readonly record struct PrintBreak(int Before, bool Manual);

public sealed record SheetPrintSettings(IReadOnlyList<PrintReference> Areas, IReadOnlyList<PrintReference> Titles,
    IReadOnlyList<PrintBreak> RowBreaks, IReadOnlyList<PrintBreak> ColumnBreaks, bool HasUnresolvedArea)
{
    internal static SheetPrintSettings Empty { get; } = new([], [], [], [], false);

    internal static SheetPrintSettings Read(XElement sheet, string? area, string? titles)
    {
        var ns = sheet.Name.Namespace;
        var areas = new List<PrintReference>();
        var unresolved = ReadReferences(area, areas);
        var titleRefs = new List<PrintReference>();
        ReadReferences(titles, titleRefs);
        return new(areas.AsReadOnly(), titleRefs.AsReadOnly(), Breaks("rowBreaks", 1048576),
            Breaks("colBreaks", 16384), unresolved);

        PrintBreak[] Breaks(string element, int limit) =>
            (sheet.Element(ns + element)?.Elements(ns + "brk") ?? [])
            .Select(node => (Id: Int((string?)node.Attribute("id")), Manual: XlsxReader.True((string?)node.Attribute("man"))))
            .Where(item => item.Id >= 1 && item.Id < limit)
            .GroupBy(item => item.Id).OrderBy(group => group.Key)
            .Select(group => new PrintBreak(group.Key + 1, group.Any(item => item.Manual))).ToArray();
    }

    private static bool ReadReferences(string? formula, List<PrintReference> references)
    {
        if (string.IsNullOrWhiteSpace(formula)) return false;
        formula = formula.Trim().TrimStart('=');
        var quoted = false;
        var parentheses = 0;
        var start = 0;
        var unresolved = false;
        for (var index = 0; index <= formula.Length; index++)
        {
            if (index < formula.Length && formula[index] == '\'')
            {
                quoted = !quoted;
                continue;
            }
            if (index < formula.Length)
            {
                if (quoted) continue;
                if (formula[index] == '(') parentheses++;
                if (formula[index] == ')') parentheses--;
                if (parentheses != 0 || formula[index] != ',') continue;
            }
            var piece = formula.AsSpan(start, index - start).Trim();
            var qualifier = piece.LastIndexOf('!');
            if (qualifier >= 0) piece = piece[(qualifier + 1)..];
            if (!quoted && parentheses == 0 && TryReference(piece.ToString().Replace("$", "", StringComparison.Ordinal), out var reference))
                references.Add(reference);
            else unresolved = true;
            start = index + 1;
        }
        return unresolved;
    }

    private static bool TryReference(string text, out PrintReference reference)
    {
        reference = default;
        var ends = text.Split(':', 2, StringSplitOptions.TrimEntries);
        var first = Endpoint(ends[0]);
        var last = Endpoint(ends[^1]);
        if (first == default || last == default || (first.Row is null) != (last.Row is null) ||
            (first.Column is null) != (last.Column is null) || first.Row > last.Row || first.Column > last.Column)
            return false;
        reference = new(first.Row, first.Column, last.Row, last.Column);
        return true;
    }

    private static (int? Row, int? Column) Endpoint(string text)
    {
        var letters = 0;
        while (letters < text.Length && char.IsAsciiLetter(text[letters])) letters++;
        var row = Int(text[letters..]);
        var column = Address.TryColumn(text.AsSpan(0, letters), out var number) ? number : (int?)null;
        if (letters > 0 && column is null || letters < text.Length && row is not (>= 1 and <= 1048576)) return default;
        return (letters < text.Length ? row : null, column);
    }

    private static int Int(string? value) => int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var number)
        ? number : 0;
}
