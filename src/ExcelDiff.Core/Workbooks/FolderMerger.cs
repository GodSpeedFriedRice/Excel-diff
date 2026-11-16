using ExcelDiff.Core.Comparison;

namespace ExcelDiff.Core.Workbooks;

public sealed record FolderMergeBook(string RelativePath, string? BasePath, string? ProductionPath, string? IncomingPath,
    WorkbookMergePlan? Plan, string? Notice);

public sealed class FolderMergePlan
{
    private readonly string[] roots;
    private readonly string[][] inputPaths;
    internal FolderMergePlan(string[] roots, string[][] inputPaths, List<FolderMergeBook> books)
    {
        this.roots = roots;
        this.inputPaths = inputPaths;
        Books = books.AsReadOnly();
    }

    public IReadOnlyList<FolderMergeBook> Books { get; }

    public void Save(string path, IReadOnlyDictionary<string, IReadOnlyDictionary<MergeCellAddress, MergeChoice>> choices,
        CancellationToken token = default)
    {
        var output = Path.TrimEndingDirectorySeparator(Path.GetFullPath(path));
        if (Directory.Exists(output) || File.Exists(output))
            throw new IOException("出力先と同名のフォルダまたはファイルがあります。新しい保存先を指定してください。");
        if (roots.Any(root => Contains(root, output)))
            throw new IOException("入力フォルダの外に保存してください。");
        var parent = Path.GetDirectoryName(output);
        if (parent is null || !Directory.Exists(parent))
            throw new DirectoryNotFoundException("保存先の親フォルダが見つかりません。");

        void VerifyFiles()
        {
            for (var i = 0; i < roots.Length; i++)
            {
                var current = FolderComparer.Enumerate(roots[i], recursive: true, token).Keys;
                if (!new HashSet<string>(inputPaths[i], StringComparer.Ordinal).SetEquals(current))
                    throw new IOException("解析後に入力フォルダのファイル一覧が変わりました。もう一度解析してください。");
            }
        }

        VerifyFiles();
        var staging = Path.Combine(parent, $".excel-diff-{Guid.NewGuid():N}.tmp");
        Directory.CreateDirectory(staging);
        try
        {
            foreach (var book in Books)
            {
                token.ThrowIfCancellationRequested();
                if (book.ProductionPath is null) continue;
                var dest = Path.Combine(staging, book.RelativePath);
                Directory.CreateDirectory(Path.GetDirectoryName(dest)!);
                if (book.Plan is { } plan)
                    plan.Save(dest, choices.GetValueOrDefault(book.RelativePath) ?? NoChoices, token);
                else
                {
                    using var source = new FileStream(book.ProductionPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                    using var target = new FileStream(dest, FileMode.CreateNew, FileAccess.Write, FileShare.None);
                    source.CopyToAsync(target, token).GetAwaiter().GetResult();
                }
            }
            VerifyFiles();
            token.ThrowIfCancellationRequested();
            Directory.Move(staging, output);
        }
        finally
        {
            if (Directory.Exists(staging)) Directory.Delete(staging, recursive: true);
        }
    }

    private static readonly IReadOnlyDictionary<MergeCellAddress, MergeChoice> NoChoices =
        new Dictionary<MergeCellAddress, MergeChoice>();

    private static bool Contains(string root, string output)
    {
        var relative = Path.GetRelativePath(root, output);
        return !Path.IsPathRooted(relative) && relative != ".." &&
            !relative.StartsWith(".." + Path.DirectorySeparatorChar, StringComparison.Ordinal);
    }
}

public static class FolderMerger
{
    public static FolderMergePlan Analyze(string baseFolder, string productionFolder, string incomingFolder,
        CancellationToken token = default, Action<FolderComparisonProgress>? reportProgress = null)
        => AnalyzeCore(baseFolder, productionFolder, incomingFolder, token, reportProgress);

    public static FolderMergePlan AnalyzePair(string productionFolder, string incomingFolder,
        CancellationToken token = default, Action<FolderComparisonProgress>? reportProgress = null)
        => AnalyzeCore(null, productionFolder, incomingFolder, token, reportProgress);

    private static FolderMergePlan AnalyzeCore(string? baseFolder, string productionFolder, string incomingFolder,
        CancellationToken token, Action<FolderComparisonProgress>? reportProgress)
    {
        var productionRoot = Path.GetFullPath(productionFolder);
        var incomingRoot = Path.GetFullPath(incomingFolder);
        var baseRoot = baseFolder is null ? null : Path.GetFullPath(baseFolder);
        var productionFiles = FolderComparer.Enumerate(productionRoot, recursive: true, token);
        var incomingFiles = FolderComparer.Enumerate(incomingRoot, recursive: true, token);
        var baseFiles = baseRoot is null ? null : FolderComparer.Enumerate(baseRoot, recursive: true, token);
        var paths = productionFiles.Keys.Union(baseFiles?.Keys.AsEnumerable() ?? [], StringComparer.OrdinalIgnoreCase)
            .Union(incomingFiles.Keys, StringComparer.OrdinalIgnoreCase).Order(StringComparer.OrdinalIgnoreCase).ToArray();
        var books = new List<FolderMergeBook>(paths.Length);
        reportProgress?.Invoke(new(0, paths.Length));
        foreach (var path in paths)
        {
            token.ThrowIfCancellationRequested();
            var original = baseFiles?.GetValueOrDefault(path)?.FullName;
            var production = productionFiles.GetValueOrDefault(path)?.FullName;
            var incoming = incomingFiles.GetValueOrDefault(path)?.FullName;
            WorkbookMergePlan? plan = null;
            string? notice = null;
            if ((baseFiles is not null && original is null) || production is null || incoming is null)
            {
                var missing = new List<string>(3);
                if (baseFiles is not null && original is null) missing.Add("共通ベース");
                if (production is null) missing.Add("取り込み1");
                if (incoming is null) missing.Add("取り込み2");
                notice = $"ファイルなし（{string.Join("・", missing)}）。追加・削除は手動で確認してください。";
            }
            else
            {
                try
                {
                    plan = original is null ? WorkbookMerger.AnalyzePair(production, incoming, token) :
                        WorkbookMerger.Analyze(original, production, incoming, token);
                }
                catch (Exception error) when (FolderComparer.IsReadError(error)) { notice = error.Message; }
            }

            books.Add(new(path, original, production, incoming, plan, notice));
            reportProgress?.Invoke(new(books.Count, paths.Length));
        }
        return baseRoot is not null && baseFiles is not null
            ? new([baseRoot, productionRoot, incomingRoot], [baseFiles.Keys.ToArray(), productionFiles.Keys.ToArray(), incomingFiles.Keys.ToArray()], books)
            : new([productionRoot, incomingRoot], [productionFiles.Keys.ToArray(), incomingFiles.Keys.ToArray()], books);
    }
}
