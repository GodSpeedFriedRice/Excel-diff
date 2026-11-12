using System.Xml.Linq;

namespace ExcelDiff.Core.Workbooks;

public sealed record Workbook(IReadOnlyList<Sheet> Sheets, bool Date1904);
public sealed record Sheet(string Name, string Visibility,
    IReadOnlyDictionary<int, IReadOnlyDictionary<int, Cell>> Rows,
    IReadOnlyDictionary<string, string> Settings, IReadOnlyDictionary<string, string> DefaultFormat, XDocument Xml)
{
    public Cell? At(int row, int column) =>
        Rows.TryGetValue(row, out var cells) ? cells.GetValueOrDefault(column) : null;
}

public sealed record Cell(string Kind, string Value, string? Formula, IReadOnlyDictionary<string, string> Format)
{
    public string Display => Kind == "b" ? (Value == "1" ? "TRUE" : "FALSE") : Value;
}

public static class Address
{
    public static bool TryColumn(ReadOnlySpan<char> name, out int column)
    {
        column = 0;
        foreach (var character in name)
        {
            var upper = char.ToUpperInvariant(character);
            if (upper is < 'A' or > 'Z') return false;
            column = column * 26 + upper - 'A' + 1;
            if (column > 16384) return false;
        }
        return column > 0;
    }

    public static (int Row, int Column) Parse(string address)
    {
        var length = 0;
        while (length < address.Length && char.IsAsciiLetter(address[length])) length++;
        if (!TryColumn(address.AsSpan(0, length), out var column) ||
            !int.TryParse(address.AsSpan(length), out var row) || row is < 1 or > 1048576)
            throw new InvalidDataException($"無効なセル番地: {address}");
        return (row, column);
    }

    public static string Column(int column)
    {
        var name = "";
        while (column > 0)
        {
            column--;
            name = (char)('A' + column % 26) + name;
            column /= 26;
        }
        return name;
    }

    public static string Format(int? row, int? column) =>
        row is null ? "—" : column is null ? $"{row} 行" : $"{Column(column.Value)}{row}";
}
