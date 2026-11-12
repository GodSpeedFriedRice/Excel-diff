using System.Xml;
using ExcelDiff.Core.Workbooks;

namespace ExcelDiff.Core.Comparison;

public enum FileStatus { Same, Different, LeftOnly, RightOnly, Unreadable }

public sealed record SheetSummary(string Name, FileStatus Status, int DifferenceCount);

public sealed record FilePair(string RelativePath, FileInfo? Left, FileInfo? Right, FileStatus Status,
    IReadOnlyList<SheetSummary> Sheets, string? Error);

public static class FolderComparer
{
    public static IReadOnlyList<FilePair> Compare(string left, string right, bool recursive,
        CancellationToken token = default)
    {
        var before = Enumerate(left, recursive);
        var after = Enumerate(right, recursive);
        var rows = new List<FilePair>();
        foreach (var name in before.Keys.Union(after.Keys, StringComparer.OrdinalIgnoreCase)
            .OrderBy(path => Path.GetDirectoryName(path), StringComparer.OrdinalIgnoreCase)
            .ThenBy(path => Path.GetFileName(path), StringComparer.OrdinalIgnoreCase))
        {
            token.ThrowIfCancellationRequested();
            var a = before.GetValueOrDefault(name);
            var b = after.GetValueOrDefault(name);
            if (a is null || b is null)
            {
                rows.Add(new FilePair(name, a, b, a is null ? FileStatus.RightOnly : FileStatus.LeftOnly, [], null));
                continue;
            }

            try
            {
                var leftBook = XlsxReader.Read(a.FullName, token);
                var rightBook = XlsxReader.Read(b.FullName, token);
                var sheets = WorkbookComparer.Compare(leftBook, rightBook, token: token)
                    .Select(SummarizeSheet).ToArray();
                var same = leftBook.Date1904 == rightBook.Date1904 && sheets.All(sheet => sheet.Status == FileStatus.Same);
                rows.Add(new FilePair(name, a, b, same ? FileStatus.Same : FileStatus.Different, sheets, null));
            }
            catch (Exception e) when (IsReadError(e))
            {
                rows.Add(new FilePair(name, a, b, FileStatus.Unreadable, [], e.Message));
            }
        }
        return rows;
    }

    public static bool IsReadError(Exception e) => e is IOException or UnauthorizedAccessException or
        InvalidDataException or XmlException or FormatException or ArgumentException or OverflowException;

    private static Dictionary<string, FileInfo> Enumerate(string root, bool recursive) =>
        new DirectoryInfo(root).EnumerateFiles("*", new EnumerationOptions
        {
            RecurseSubdirectories = recursive,
            AttributesToSkip = FileAttributes.ReparsePoint,
            IgnoreInaccessible = false
        }).Where(file => file.Extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) &&
            !file.Name.StartsWith("~$", StringComparison.Ordinal))
          .ToDictionary(file => Path.GetRelativePath(root, file.FullName), StringComparer.OrdinalIgnoreCase);

    private static SheetSummary SummarizeSheet(SheetDiff sheet)
    {
        var status = sheet.Left is null ? FileStatus.RightOnly :
            sheet.Right is null ? FileStatus.LeftOnly :
            sheet.Differences.Count > 0 ? FileStatus.Different : FileStatus.Same;
        return new SheetSummary(sheet.Name, status, sheet.Differences.Count);
    }
}
