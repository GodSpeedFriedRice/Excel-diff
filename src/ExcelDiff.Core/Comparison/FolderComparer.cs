using System.Xml;
using ExcelDiff.Core.Workbooks;

namespace ExcelDiff.Core.Comparison;

public enum FileStatus { Same, Different, LeftOnly, RightOnly, Unreadable }

public sealed record SheetSummary(string Name, FileStatus Status, int DifferenceCount);

public sealed record FilePair(string RelativePath, FileInfo? Left, FileInfo? Right, FileStatus Status,
    IReadOnlyList<SheetSummary> Sheets, string? Error);

public sealed record FolderComparisonProgress(int Completed, int Total);

public static class FolderComparer
{
    public static IReadOnlyList<FilePair> Compare(string left, string right, bool recursive,
        CancellationToken token = default) => Compare(left, right, recursive, token, null);

    public static IReadOnlyList<FilePair> Compare(string left, string right, bool recursive,
        CancellationToken token, Action<FolderComparisonProgress>? reportProgress)
    {
        var before = Enumerate(left, recursive, token);
        var after = Enumerate(right, recursive, token);
        var names = before.Keys.Union(after.Keys, StringComparer.OrdinalIgnoreCase)
            .OrderBy(path => Path.GetDirectoryName(path), StringComparer.OrdinalIgnoreCase)
            .ThenBy(path => Path.GetFileName(path), StringComparer.OrdinalIgnoreCase)
            .ToArray();
        var rows = new List<FilePair>(names.Length);
        token.ThrowIfCancellationRequested();
        reportProgress?.Invoke(new(0, names.Length));
        foreach (var name in names)
        {
            token.ThrowIfCancellationRequested();
            rows.Add(CompareFile(name, before.GetValueOrDefault(name), after.GetValueOrDefault(name), token));
            reportProgress?.Invoke(new(rows.Count, names.Length));
        }
        return rows;
    }

    private static FilePair CompareFile(string name, FileInfo? a, FileInfo? b, CancellationToken token)
    {
        if (a is null || b is null)
            return new FilePair(name, a, b, a is null ? FileStatus.RightOnly : FileStatus.LeftOnly, [], null);

        try
        {
            var leftBook = XlsxReader.Read(a.FullName, token);
            var rightBook = XlsxReader.Read(b.FullName, token);
            var sheets = WorkbookComparer.Compare(leftBook, rightBook, token: token)
                .Select(SummarizeSheet).ToArray();
            var same = leftBook.Date1904 == rightBook.Date1904 && sheets.All(sheet => sheet.Status == FileStatus.Same);
            return new FilePair(name, a, b, same ? FileStatus.Same : FileStatus.Different, sheets, null);
        }
        catch (Exception e) when (IsReadError(e))
        {
            return new FilePair(name, a, b, FileStatus.Unreadable, [], e.Message);
        }
    }

    public static bool IsReadError(Exception e) => e is IOException or UnauthorizedAccessException or
        InvalidDataException or XmlException or FormatException or ArgumentException or OverflowException;

    private static Dictionary<string, FileInfo> Enumerate(string root, bool recursive, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        var files = new Dictionary<string, FileInfo>(StringComparer.OrdinalIgnoreCase);
        var options = new EnumerationOptions
        {
            RecurseSubdirectories = recursive,
            AttributesToSkip = FileAttributes.ReparsePoint,
            IgnoreInaccessible = false
        };
        foreach (var file in new DirectoryInfo(root).EnumerateFiles("*", options))
        {
            token.ThrowIfCancellationRequested();
            if (file.Extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) &&
                !file.Name.StartsWith("~$", StringComparison.Ordinal))
                files.Add(Path.GetRelativePath(root, file.FullName), file);
        }
        return files;
    }

    private static SheetSummary SummarizeSheet(SheetDiff sheet)
    {
        var status = sheet.Left is null ? FileStatus.RightOnly :
            sheet.Right is null ? FileStatus.LeftOnly :
            sheet.Differences.Count > 0 ? FileStatus.Different : FileStatus.Same;
        return new SheetSummary(sheet.Name, status, sheet.Differences.Count);
    }
}
