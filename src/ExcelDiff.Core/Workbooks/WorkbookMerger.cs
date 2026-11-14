using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.RegularExpressions;

namespace ExcelDiff.Core.Workbooks;

public enum MergeDisposition { KeepProduction, TakeIncoming, AlreadyMatches, Conflict, Unsupported }
public enum MergeChoice { Production, Incoming }

public sealed record MergeCellAddress(string Sheet, int Row, int Column)
{
    public string Address => Workbooks.Address.Format(Row, Column);
}

public sealed record MergeCellChange(MergeCellAddress Cell, string Base, string Production, string Incoming,
    MergeDisposition Disposition, string? Reason = null);

public sealed record MergeNotice(string Sheet, string Reason);

public sealed class WorkbookMergePlan
{
    private readonly Input[] inputs;
    private readonly Edit[] edits;

    internal WorkbookMergePlan(Input[] inputs, List<Edit> edits, List<MergeNotice> notices)
    {
        this.inputs = inputs;
        this.edits = edits.ToArray();
        Changes = Array.AsReadOnly(this.edits.Select(edit => edit.Change).ToArray());
        Notices = notices.AsReadOnly();
    }

    public IReadOnlyList<MergeCellChange> Changes { get; }
    public IReadOnlyList<MergeNotice> Notices { get; }

    public void Save(string path, IReadOnlyDictionary<MergeCellAddress, MergeChoice> choices, CancellationToken token = default)
    {
        var output = WorkbookMerger.XlsxPath(path);
        if (inputs.Any(input => string.Equals(input.Path, output, StringComparison.OrdinalIgnoreCase)))
            throw new IOException("入力ファイルとは別のファイル名を指定してください。");
        if (File.Exists(output)) throw new IOException("出力先に同名のファイルがあります。");
        var writes = new List<Edit>();
        foreach (var edit in edits)
        {
            token.ThrowIfCancellationRequested();
            switch (edit.Change.Disposition)
            {
                case MergeDisposition.TakeIncoming:
                    writes.Add(edit);
                    break;
                case MergeDisposition.Conflict:
                    if (!choices.TryGetValue(edit.Change.Cell, out var choice) ||
                        choice is not (MergeChoice.Production or MergeChoice.Incoming))
                        throw new InvalidOperationException("すべての競合で採用する内容を選択してください。");
                    if (choice == MergeChoice.Incoming) writes.Add(edit);
                    break;
            }
        }

        using var original = inputs[0].Open();
        using var production = inputs[1].Open();
        using var incoming = inputs[2].Open();
        inputs[0].Verify(original, token);
        inputs[1].Verify(production, token);
        inputs[2].Verify(incoming, token);
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

    internal sealed record Edit(MergeCellChange Change, string Part, MergeContent Incoming);
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
        CancellationToken token = default)
    {
        var (baseInput, baseBook) = Read(basePath, "共通ベース", token);
        var (productionInput, production) = Read(productionPath, "本番版", token);
        var (incomingInput, incoming) = Read(incomingPath, "取り込み元", token);
        if (baseBook.Date1904 != production.Date1904 || baseBook.Date1904 != incoming.Date1904)
            throw new InvalidDataException("3冊の日付システム（1900年／1904年）が一致している必要があります。");

        var edits = new List<WorkbookMergePlan.Edit>();
        var notices = new List<MergeNotice>();
        var originalSheets = ByName(baseBook);
        var productionSheets = ByName(production);
        var incomingSheets = ByName(incoming);
        foreach (var name in originalSheets.Keys.Union(productionSheets.Keys).Union(incomingSheets.Keys))
        {
            token.ThrowIfCancellationRequested();
            var original = originalSheets.GetValueOrDefault(name);
            var destination = productionSheets.GetValueOrDefault(name);
            var source = incomingSheets.GetValueOrDefault(name);
            if (original is null || destination is null || source is null)
            {
                if (original is not null || (destination is null) != (source is null))
                    notices.Add(new(name, "シートの追加・削除・名前変更は手動で確認してください。"));
                else
                    notices.Add(new(name, "両側に同名のシートが追加されています。本番版を保持します。"));
                continue;
            }
            if (!original.Merges.Ranges.SequenceEqual(destination.Merges.Ranges) ||
                !original.Merges.Ranges.SequenceEqual(source.Merges.Ranges))
            {
                notices.Add(new(name, "結合範囲が変わったシートです。本番版を保持します。"));
                continue;
            }

            var baseline = new MergeWorksheet(original, token);
            var current = new MergeWorksheet(destination, token);
            var other = new MergeWorksheet(source, token);
            if (baseline.HasMovedLines(current, token) || baseline.HasMovedLines(other, token))
            {
                notices.Add(new(name, "行・列の移動を検出しました。本番版を保持します。"));
                continue;
            }

            foreach (var row in baseline.Rows.Keys.Union(current.Rows.Keys).Union(other.Rows.Keys).Order())
            {
                token.ThrowIfCancellationRequested();
                var originalRow = baseline.Rows.GetValueOrDefault(row);
                var currentRow = current.Rows.GetValueOrDefault(row);
                var incomingRow = other.Rows.GetValueOrDefault(row);
                var merges = original.Merges.InRows(row, row);
                var currentFormulaRanges = current.FormulaRanges.InRows(row, row);
                var incomingFormulaRanges = other.FormulaRanges.InRows(row, row);
                foreach (var column in Columns(originalRow).Union(Columns(currentRow)).Union(Columns(incomingRow)).Order())
                {
                    if (merges.Any(range => range.ContainsColumn(column) && !range.IsAnchor(row, column))) continue;
                    var before = originalRow?.GetValueOrDefault(column) ?? MergeContent.Blank;
                    var ours = currentRow?.GetValueOrDefault(column) ?? MergeContent.Blank;
                    var theirs = incomingRow?.GetValueOrDefault(column) ?? MergeContent.Blank;
                    if (theirs.Key == before.Key && ours.Key == before.Key) continue;

                    string? reason = null;
                    MergeDisposition disposition;
                    if (theirs.Key == before.Key) disposition = MergeDisposition.KeepProduction;
                    else if (theirs.Key == ours.Key) disposition = MergeDisposition.AlreadyMatches;
                    else if (ours.Special || theirs.Special ||
                        currentFormulaRanges.Any(range => range.ContainsColumn(column)) ||
                        incomingFormulaRanges.Any(range => range.ContainsColumn(column)) ||
                        (theirs.Key.Kind == "formula" && ExternalBookReference().IsMatch(theirs.Key.Text)))
                    {
                        disposition = MergeDisposition.Unsupported;
                        reason = "セル内の文字装飾、特殊な数式、外部参照などを含みます。本番版を保持します。";
                    }
                    else disposition = ours.Key == before.Key ? MergeDisposition.TakeIncoming : MergeDisposition.Conflict;

                    var change = new MergeCellChange(new(name, row, column), before.Display, ours.Display, theirs.Display,
                        disposition, reason);
                    edits.Add(new(change, destination.XmlPartPath!, theirs));
                }
            }
        }
        if (notices.Count > 0)
            for (var index = 0; index < edits.Count; index++)
            {
                var edit = edits[index];
                if (edit.Incoming.Key.Kind == "formula" &&
                    edit.Change.Disposition is MergeDisposition.TakeIncoming or MergeDisposition.Conflict &&
                    edit.Incoming.Key.Text.Contains('!'))
                    edits[index] = edit with
                    {
                        Change = edit.Change with
                        {
                            Disposition = MergeDisposition.Unsupported,
                            Reason = "構造変更のあるブック内でシートを参照する数式です。本番版を保持します。"
                        }
                    };
            }
        return new([baseInput, productionInput, incomingInput], edits, notices);
    }

    private static (WorkbookMergePlan.Input Input, Workbook Book) Read(string path, string label, CancellationToken token)
    {
        path = XlsxPath(path);
        try
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            var hash = WorkbookMergePlan.Input.Fingerprint(stream, token);
            return (new(path, label, hash), XlsxReader.Read(stream, token));
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
