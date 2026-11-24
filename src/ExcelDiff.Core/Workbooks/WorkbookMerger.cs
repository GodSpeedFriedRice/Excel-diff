using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using ExcelDiff.Core.Comparison;

namespace ExcelDiff.Core.Workbooks;

public enum MergeDisposition { KeepProduction, TakeIncoming, AlreadyMatches, Conflict, Unsupported }
public enum MergeChoice { Production, Incoming }

public sealed record MergeCellAddress(string Sheet, int Row, int Column)
{
    public string Address => Workbooks.Address.Format(Row, Column);
}

public sealed record MergeCellChange(MergeCellAddress Cell, string Base, string Production, string Incoming,
    MergeDisposition Disposition, string? Reason = null)
{
    internal bool AllowsChoice { get; init; } = true;
    public bool CanChoose => Disposition != MergeDisposition.Unsupported && AllowsChoice;
    public IReadOnlyList<MergeRange> MergeRanges { get; init; } = [];
    public string Address => MergeRanges.Count == 0 ? Cell.Address : string.Join(" / ", MergeRanges);
}

public sealed record MergeNotice(string Sheet, string Reason);

public sealed class WorkbookMergePlan
{
    private readonly Input? baseInput;
    private readonly Input productionInput;
    private readonly Input incomingInput;
    private readonly Edit[] edits;

    internal WorkbookMergePlan(Input? baseInput, Input productionInput, Input incomingInput, List<Edit> edits,
        List<MergeNotice> notices, IReadOnlyList<PackagePartMergeDifference> partChanges, IReadOnlyList<PackagePartDifference> pairParts)
    {
        this.baseInput = baseInput;
        this.productionInput = productionInput;
        this.incomingInput = incomingInput;
        this.edits = edits.ToArray();
        Changes = Array.AsReadOnly(this.edits.Select(edit => edit.Change).ToArray());
        Notices = notices.AsReadOnly();
        RequiresManualResolution = notices.Count > 0 || Changes.Any(change => change.Disposition == MergeDisposition.Unsupported);
        PartChanges = partChanges;
        PairParts = pairParts;
    }

    public IReadOnlyList<MergeCellChange> Changes { get; }
    public IReadOnlyList<MergeNotice> Notices { get; }
    public bool RequiresManualResolution { get; }
    public IReadOnlyList<PackagePartMergeDifference> PartChanges { get; }
    public IReadOnlyList<PackagePartDifference> PairParts { get; }

    public Workbook?[] ReadPreview(CancellationToken token = default) =>
        [baseInput is null ? null : ReadPreview(baseInput, token), ReadPreview(productionInput, token), ReadPreview(incomingInput, token)];

    private static Workbook ReadPreview(Input input, CancellationToken token)
    {
        using var stream = input.Open();
        input.Verify(stream, token);
        return XlsxReader.Read(input.Path, token);
    }

    public void Save(string path, IReadOnlyDictionary<MergeCellAddress, MergeChoice> choices, CancellationToken token = default)
    {
        var output = WorkbookMerger.XlsxPath(path);
        if (new[] { baseInput?.Path, productionInput.Path, incomingInput.Path }.Contains(output, StringComparer.OrdinalIgnoreCase))
            throw new IOException("入力ファイルとは別のファイル名を指定してください。");
        if (File.Exists(output)) throw new IOException("出力先に同名のファイルがあります。");
        if (RequiresManualResolution)
            throw new InvalidOperationException("手動で解消が必要な項目があります。Excel で修正してから再比較してください。");
        var writes = new List<Edit>();
        foreach (var edit in edits)
        {
            token.ThrowIfCancellationRequested();
            var hasChoice = choices.TryGetValue(edit.Change.Cell, out var choice);
            if (hasChoice && (!edit.Change.CanChoose || choice is not (MergeChoice.Production or MergeChoice.Incoming)))
                throw new InvalidOperationException("このセルでは採用する内容を変更できません。");
            if (!hasChoice && edit.Change.Disposition == MergeDisposition.Conflict)
                throw new InvalidOperationException("すべてのコンフリクトで採用する内容を選択してください。");
            if (hasChoice ? choice == MergeChoice.Incoming : edit.Change.Disposition == MergeDisposition.TakeIncoming)
                writes.Add(edit);
        }

        using var original = baseInput?.Open();
        using var production = productionInput.Open();
        using var incoming = incomingInput.Open();
        if (original is not null) baseInput!.Verify(original, token);
        productionInput.Verify(production, token);
        incomingInput.Verify(incoming, token);
        var temporary = Path.Combine(Path.GetDirectoryName(output)!, $".excel-diff-{Guid.NewGuid():N}.tmp");
        var created = false;
        try
        {
            using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
            {
                created = true;
                production.CopyToAsync(stream, token).GetAwaiter().GetResult();
                if (writes.Count > 0)
                {
                    stream.Position = 0;
                    using var zip = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true);
                    XlsxMergeWriter.Apply(zip, writes, token);
                }
            }
            token.ThrowIfCancellationRequested();
            File.Move(temporary, output);
        }
        finally
        {
            if (created) File.Delete(temporary);
        }
    }

    internal sealed record Edit(MergeCellChange Change, string Part, MergeContent Incoming)
    {
        internal LayoutEdit? Layout { get; init; }
    }

    internal sealed record LayoutEdit(MergeIndex Scope, IReadOnlyList<MergeRange> IncomingMerges,
        IReadOnlyDictionary<(int Row, int Column), MergeContent> IncomingCells);
    internal sealed record Input(string Path, string Label, byte[] Hash)
    {
        internal FileStream Open() => new(Path, FileMode.Open, FileAccess.Read, FileShare.Read);

        internal void Verify(Stream stream, CancellationToken token)
        {
            if (!Hash.AsSpan().SequenceEqual(Fingerprint(stream, token)))
                throw new IOException($"{Label}が解析後に変更されました。もう一度解析してください。");
        }

        internal static byte[] Fingerprint(Stream stream, CancellationToken token)
        {
            using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
            var buffer = new byte[64 * 1024];
            int read;
            while ((read = stream.Read(buffer)) != 0)
            {
                token.ThrowIfCancellationRequested();
                hash.AppendData(buffer.AsSpan(0, read));
            }
            stream.Position = 0;
            return hash.GetHashAndReset();
        }
    }
}

public static partial class WorkbookMerger
{
    public static WorkbookMergePlan Analyze(string basePath, string productionPath, string incomingPath,
        CancellationToken token = default) => AnalyzeCore(basePath, productionPath, incomingPath, token);

    public static WorkbookMergePlan AnalyzePair(string productionPath, string incomingPath, CancellationToken token = default)
        => AnalyzeCore(null, productionPath, incomingPath, token);

    private static WorkbookMergePlan AnalyzeCore(string? basePath, string productionPath, string incomingPath,
        CancellationToken token)
    {
        (WorkbookMergePlan.Input? Input, Workbook? Book, Dictionary<string, PackagePart>? Parts) baselineBook =
            basePath is null ? default : Read(basePath, "共通ベース", token);
        var hasBase = baselineBook.Book is not null;
        var (productionInput, production, productionParts) = Read(productionPath, "取り込み1", token);
        var (incomingInput, incoming, incomingParts) = Read(incomingPath, "取り込み2", token);
        var edits = new List<WorkbookMergePlan.Edit>();
        var notices = new List<MergeNotice>();
        var partChanges = hasBase ? PackageComparer.Compare(baselineBook.Parts!, productionParts, incomingParts, token) : [];
        var pairParts = hasBase ? [] : PackageComparer.Compare(productionParts, incomingParts, token);
        WorkbookMergePlan Result()
        {
            if (hasBase && notices.Count > 0 && partChanges.All(change => change.Status == PackagePartMergeStatus.SameChange))
                notices.Clear();
            return new(baselineBook.Input, productionInput, incomingInput, edits, notices, partChanges, pairParts);
        }
        if (production.Date1904 != incoming.Date1904 ||
            (baselineBook.Book is { } baseBook && baseBook.Date1904 != production.Date1904))
        {
            notices.Add(new("ブック", "日付システム（1900年／1904年）が異なるため、セルの取り込みは停止します。"));
            return Result();
        }
        var originalSheets = baselineBook.Book is { } originalBook ? ByName(originalBook) : [];
        var productionSheets = ByName(production);
        var incomingSheets = ByName(incoming);
        foreach (var name in originalSheets.Keys.Union(productionSheets.Keys).Union(incomingSheets.Keys))
        {
            token.ThrowIfCancellationRequested();
            var original = originalSheets.GetValueOrDefault(name);
            var destination = productionSheets.GetValueOrDefault(name);
            var source = incomingSheets.GetValueOrDefault(name);
            if ((hasBase && original is null) || destination is null || source is null)
            {
                if (!hasBase || original is not null || (destination is null) != (source is null))
                    notices.Add(new(name, "シートの追加・削除・名前変更は手動で確認してください。"));
                else
                    notices.Add(new(name, "両側に同名のシートが追加されています。手動で確認してください。"));
                continue;
            }
            var baseline = original is null ? null : new MergeWorksheet(original, token);
            var current = new MergeWorksheet(destination, token);
            var other = new MergeWorksheet(source, token);
            if (baseline is null ? current.HasMovedLines(other, token) :
                baseline.HasMovedLines(current, token) || baseline.HasMovedLines(other, token))
            {
                notices.Add(new(name, "行・列の移動を検出しました。手動で確認してください。"));
                continue;
            }

            var mergeScope = AddMergeEdits(original, destination, source, baseline, current, other, edits, token);

            foreach (var row in (baseline?.Rows.Keys.AsEnumerable() ?? []).Union(current.Rows.Keys).Union(other.Rows.Keys).Order())
            {
                token.ThrowIfCancellationRequested();
                var originalRow = baseline?.Rows.GetValueOrDefault(row);
                var currentRow = current.Rows.GetValueOrDefault(row);
                var incomingRow = other.Rows.GetValueOrDefault(row);
                var merges = destination.Merges.InRows(row, row);
                var changedMerges = mergeScope.InRows(row, row);
                var currentFormulaRanges = current.FormulaRanges.InRows(row, row);
                var incomingFormulaRanges = other.FormulaRanges.InRows(row, row);
                foreach (var column in Columns(originalRow).Union(Columns(currentRow)).Union(Columns(incomingRow)).Order())
                {
                    if (changedMerges.Any(range => range.ContainsColumn(column))) continue;
                    if (merges.Any(range => range.ContainsColumn(column) && !range.IsAnchor(row, column))) continue;
                    var before = originalRow?.GetValueOrDefault(column) ?? MergeContent.Blank;
                    var ours = currentRow?.GetValueOrDefault(column) ?? MergeContent.Blank;
                    var theirs = incomingRow?.GetValueOrDefault(column) ?? MergeContent.Blank;
                    if (hasBase ? theirs.Key == before.Key && ours.Key == before.Key : theirs.Key == ours.Key) continue;

                    string? reason = null;
                    var allowsChoice = !ours.Special && !theirs.Special &&
                        !currentFormulaRanges.Any(range => range.ContainsColumn(column)) &&
                        !incomingFormulaRanges.Any(range => range.ContainsColumn(column)) &&
                        !(theirs.Key.Kind == "formula" && ExternalBookReference().IsMatch(theirs.Key.Text));
                    MergeDisposition disposition;
                    if (hasBase && theirs.Key == before.Key) disposition = MergeDisposition.KeepProduction;
                    else if (hasBase && theirs.Key == ours.Key) disposition = MergeDisposition.AlreadyMatches;
                    else if (!allowsChoice)
                    {
                        disposition = MergeDisposition.Unsupported;
                        reason = "セル内の文字装飾、特殊な数式、外部参照などを含みます。手動で解消してください。";
                    }
                    else disposition = hasBase && ours.Key == before.Key ? MergeDisposition.TakeIncoming : MergeDisposition.Conflict;

                    var change = new MergeCellChange(new(name, row, column), hasBase ? before.Display : "", ours.Display, theirs.Display,
                        disposition, reason) { AllowsChoice = allowsChoice };
                    edits.Add(new(change, destination.XmlPartPath!, theirs));
                }
            }
        }
        if (notices.Count > 0)
            for (var index = 0; index < edits.Count; index++)
            {
                var edit = edits[index];
                var referencesSheet = edit.Layout is { } layout
                    ? layout.IncomingCells.Values.Any(cell => cell.Key.Kind == "formula" && cell.Key.Text.Contains('!'))
                    : edit.Incoming.Key.Kind == "formula" && edit.Incoming.Key.Text.Contains('!');
                if (!referencesSheet) continue;
                var change = edit.Change;
                if (change.Disposition is MergeDisposition.TakeIncoming or MergeDisposition.Conflict)
                    change = change with
                    {
                        Disposition = MergeDisposition.Unsupported,
                        Reason = "構造変更のあるブック内でシートを参照する数式です。手動で解消してください。"
                    };
                edits[index] = edit with { Change = change with { AllowsChoice = false } };
            }
        return Result();
    }

    private static (WorkbookMergePlan.Input Input, Workbook Book, Dictionary<string, PackagePart> Parts) Read(
        string path, string label, CancellationToken token)
    {
        path = XlsxPath(path);
        try
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            var hash = WorkbookMergePlan.Input.Fingerprint(stream, token);
            var book = XlsxReader.Read(stream, token);
            stream.Position = 0;
            return (new(path, label, hash), book, PackageComparer.Read(stream, token));
        }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException or InvalidDataException or System.Xml.XmlException)
        {
            throw new InvalidDataException($"{label}を読み込めません。{error.Message}", error);
        }
    }

    internal static string XlsxPath(string path)
    {
        if (!string.Equals(Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException(".xlsx ファイルを指定してください。");
        return Path.GetFullPath(path);
    }

    private static Dictionary<string, Sheet> ByName(Workbook book) => book.Sheets.ToDictionary(sheet => sheet.Name);
    private static IEnumerable<int> Columns(Dictionary<int, MergeContent>? row) => row?.Keys ?? Enumerable.Empty<int>();

    [GeneratedRegex(@"\[\d+\]")]
    private static partial Regex ExternalBookReference();
}
